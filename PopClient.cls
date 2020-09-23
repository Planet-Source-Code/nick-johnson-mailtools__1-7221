VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Handles communication with a POP server to retrieve mail etc. Simply encapsulates POP server functions for a VB environment. Requires a Winsock control passed to it to function."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This class module is designed to encapsulate for VB the neccessary
'function calls to a POP server, and interpret the responses.
'It still requires a basic knowledge of how to retrieve
'messages off a POP server, but goes something like this:

'Give the PopClient Object a Winsock Object
'Set the ServerName
'Set the ServerPort
'Connect
'User <username>
'Pass <password>
'Stat
'If there are messages:
'   List
'   For each message:
'       Top(<messageno>, 0)
'       If the message passes any rules:
'           Retr <message>
'For a basic retrieval. After connect, user, and pass,
'operations can be performed in any order.

'This module uses 'hard' errors. If an error occurs in your
'input, winsock, or the server returns an unexpected response,
'an error will be raised. I reccommend setting error handling
'to Resume Next, and checking for expected errors after each
'operation.

'local variable(s) to hold property value(s)
Private WithEvents pWsConnection As Winsock 'local copy
Attribute pWsConnection.VB_VarHelpID = -1
Private pTimeout As Integer 'local copy

'Local variables
Private strServerResponse As String
Private colMessageQueue As New Collection 'Stores waiting server messages
Private sngTimer As Single
Private a As Integer 'Loop counter

'local error code constants
Private Const errNotConnected = 1 + vbObjectError
Private Const errTimeout = 2 + vbObjectError
Private Const errNoWinsock = 3 + vbObjectError
Private Const errServerError = 4 + vbObjectError
Private Const errInvalidServerResponse = 5 + vbObjectError
Private Const errInvalidMessage = 6 + vbObjectError

Private Const debugging = False

Public Function Uidl(Optional messageNo As Integer = -1) As Collection
Attribute Uidl.VB_Description = "Returns a 2 dimensional array, 1 row per message. The first column contains the string representation of the message number, the second column contains the messages Unique ID No"
    'Retuns a collection of two element arrays of type string.
    'The first element in each array contains the message
    'number, and the second element contains a unique ID no
    'for that message
    If Not connected Then
        Err.Raise errNotConnected, "PopClient", "Not connected to server!"
    Else
        Dim colResponse As Collection
        Dim strCurrentResult(1 To 2) As String
                
        Set Uidl = New Collection
        If messageNo > 0 Then
            sendStr "UIDL " + Str(messageNo) + vbCrLf
            Set colResponse = New Collection
            colResponse.Add getOnelineResponse
        Else
            sendStr "UIDL " + vbCrLf
            Set colResponse = tokeniseString(removeChr(getMultilineResponse, 10), 13)
        End If
        For a = 1 To colResponse.Count
            strCurrentResult(1) = Left(colResponse(a), InStr(1, colResponse(a), " ") - 1)
            strCurrentResult(2) = Right(colResponse(a), Len(colResponse(a)) - InStr(1, colResponse(a), " "))
            Uidl.Add strCurrentResult
        Next
    End If
End Function

Public Function Top(messageNo As Integer, lines As Integer) As Message
Attribute Top.VB_Description = "Returns the message header, a blank line, then the first <lines> lines of the message."
    'Retrieves the header, a blank line, then the first <lines>
    'lines of the message.
    If Not connected Then
        Err.Raise errNotConnected, "PopClient", "Not connected to server!"
    Else
        sendStr "TOP " + Str(messageNo) + " " + Str(lines) + vbCrLf
        Set Top = strToMessage(getMultilineResponse)
    End If
End Function

Public Sub Quit()
Attribute Quit.VB_Description = "Disconnects from the server."
    'Disconnects from the server
    If Not connected Then
        Err.Raise errNotConnected, "PopClient", "Not connected to server!"
    Else
        sendStr "QUIT" + vbCrLf
        waitForOk
        pWsConnection.Close
    End If
End Sub

Public Sub Reset()
Attribute Reset.VB_Description = "Unmarks all messages marked for deletion."
    'Resets all messages marked for deletion to normal status
    If Not connected Then
        Err.Raise errNotConnected, "PopClient", "Not connected to server!"
    Else
        sendStr "RSET" + vbCrLf
        waitForOk
    End If
End Sub

Public Sub Noop()
Attribute Noop.VB_Description = "Sends a NOOP command to the server, does nothing."
    'Sends a NoOp to the server to keep the connection 'alive'
    If Not connected Then
        Err.Raise errNotConnected, "PopClient", "Not connected to server!"
    Else
        sendStr "NOOP" + vbCrLf
        waitForOk
    End If
End Sub

Public Sub Dele(messageNo As Integer)
Attribute Dele.VB_Description = "Deletes the specified message."
    'Deletes the specified message
    If Not connected Then
        Err.Raise errNotConnected, "PopClient", "Not connected to server!"
    Else
        sendStr "DELE " + Str(messageNo) + vbCrLf
        waitForOk
    End If
End Sub

Public Function Retr(messageNo As Integer) As Message
Attribute Retr.VB_Description = "Returns the specified message, including headers."
    'Retrieves the specified message
    If Not connected Then
        Err.Raise errNotConnected, "PopClient", "Not connected to server!"
    Else
        sendStr "RETR " + Str(messageNo) + vbCrLf
        Set Retr = strToMessage(getMultilineResponse)
    End If
End Function

Public Function List(Optional messageNo As Integer = -1) As Collection
Attribute List.VB_Description = "Returns a 2 dimensional array. Each row contains information on a specific message. Column 1 contains the message number, column 2 contains the size of the message."
    'Returns a collection of two item arrays of type long.
    'Each array represents a message, the first number is the
    'message number, the second is it's size in bytes.
    If Not connected Then
        Err.Raise errNotConnected, "PopClient", "Not connected to server!"
    Else
        Dim colResponse As Collection
        Dim colCurrentLine As Collection
        Dim lngCurrentResult(1 To 2) As Long
                
        Set List = New Collection
        If messageNo > 0 Then
            sendStr "LIST " + Str(messageNo) + vbCrLf
            Set colResponse = New Collection
            colResponse.Add getOnelineResponse
        Else
            sendStr "LIST " + vbCrLf
            Set colResponse = tokeniseString(removeChr(getMultilineResponse, 10), 13)
        End If
        For a = 1 To colResponse.Count
            Set colCurrentLine = tokeniseString(colResponse(a), Asc(" "))
            lngCurrentResult(1) = Val(colCurrentLine(1))
            lngCurrentResult(2) = Val(colCurrentLine(2))
            List.Add lngCurrentResult
        Next
    End If
End Function

Public Function Stat() As Long()
Attribute Stat.VB_Description = "Returns a two element array. Element 1 contains the number of messages, element 2 contains the total size."
    'Returns two longs - the first one contains the number
    'of messages waiting, the second contains the total
    'byte size of the messages.
    If Not connected Then
        Err.Raise errNotConnected, "PopClient", "Not connected to server!"
    Else
        Dim colResponse As Collection
        Dim lngStat(1 To 2) As Long
        
        sendStr "STAT" + vbCrLf
        Set colResponse = tokeniseString(getOnelineResponse, Asc(" "))
        If colResponse.Count < 2 Then
            Err.Raise errInvalidServerResponse, "PopClient", "The server returned an invalid response"
        Else
            lngStat(1) = colResponse(1)
            lngStat(2) = colResponse(2)
            Stat = lngStat
        End If
    End If
End Function

Public Sub Pass(password As String)
Attribute Pass.VB_Description = "Sends the server the specified password"
    'Sends a password to the POP server
    If Not connected Then
        Err.Raise errNotConnected, "PopClient", "Not connected to server!"
    Else
        sendStr "PASS " + password + vbCrLf
        waitForOk
    End If
End Sub

Public Sub User(username As String)
Attribute User.VB_Description = "Sends the server the specified username"
    'Sends a username to the POP server
    If Not connected Then
        Err.Raise errNotConnected, "PopClient", "Not connected to server!"
    Else
        sendStr "USER " + username + vbCrLf
        waitForOk
    End If
End Sub

Public Sub Connect()
Attribute Connect.VB_Description = "Connects to the specified server."
    'Connects to the POP server and waits for an OK response
    pWsConnection.Connect
    waitForOk
End Sub

Public Property Let serverPort(ByVal vData As Integer)
Attribute serverPort.VB_Description = "The port to connect to. Defaults to standard POP (110)"
    If Not pWsConnection Is Nothing Then
        pWsConnection.RemotePort = vData
    Else
        Err.Raise errNoWinsock, "PopClient", "Winsock control not supplied"
    End If
End Property
Public Property Get serverPort() As Integer
    If Not pWsConnection Is Nothing Then
        serverPort = pWsConnection.RemotePort
    Else
        Err.Raise errNoWinsock, "PopClient", "Winsock control not supplied"
    End If
End Property

Public Property Let serverName(ByVal vData As String)
Attribute serverName.VB_Description = "The hostname of the server to connect to. Defaults to localhost."
    If Not pWsConnection Is Nothing Then
        pWsConnection.RemoteHost = vData
    Else
        Err.Raise errNoWinsock, "PopClient", "Winsock control not supplied"
    End If
End Property
Public Property Get serverName() As String
    If Not pWsConnection Is Nothing Then
        serverName = pWsConnection.RemoteHost
    Else
        Err.Raise errNoWinsock, "PopClient", "Winsock control not supplied"
    End If
End Property

Public Property Let timeout(vData As Integer)
    'Sets the time to wait before timing out waiting for a reply.
    'BEWARE! Neccessary for lapsed server, but still times
    'while making a connection. Anyone have a better suggestion?
    pTimeout = vData
End Property
Public Property Get timeout() As Integer
    timeout = pTimeout
End Property

Public Property Set wsConnection(ByVal vData As Winsock)
Attribute wsConnection.VB_Description = "The winsock control PopClient uses to connect to the POP server."
    Set pWsConnection = vData
End Property
Public Property Get wsConnection() As Winsock
    Set pWsConnection = pWsConnection
End Property

Private Sub Class_Initialize()
    pTimeout = 30
End Sub

Private Sub pWsConnection_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    pWsConnection.GetData strData, vbString
    If debugging Then
        Form1.Text1.Text = Form1.Text1.Text + strData
    End If
    For a = 1 To Len(strData)
        pWsConnection_ChrArrival (Asc(Mid(strData, a, 1)))
    Next a
End Sub

Private Sub pWsConnection_ChrArrival(bChr As Byte)
    If bChr = 13 Then
        colMessageQueue.Add strServerResponse
        strServerResponse = ""
    ElseIf bChr <> 10 Then
        strServerResponse = strServerResponse + Chr(bChr)
    End If
End Sub

Private Sub setTimer()
    sngTimer = Timer
End Sub

Private Function getTimer() As Boolean
    If sngTimer + pTimeout <= Timer Then
        getTimer = False
    Else
        getTimer = True
    End If
End Function

Private Sub waitForOk()
    setTimer
    While colMessageQueue.Count < 1 And getTimer
        DoEvents
    Wend
    If getTimer Then
        Dim strServerLine As String
        strServerLine = colMessageQueue(1)
        colMessageQueue.Remove 1
        If UCase(Left(strServerLine, 3)) <> "+OK" Then
            Err.Raise errServerError, "PopClient", strServerLine
        End If
    Else
        Err.Raise errTimeout, "PopClient", "Timed out waiting for server"
    End If
End Sub

Private Function getOnelineResponse() As String
    'Same as waitForOk, but returns all the text after +OK (if any)
    setTimer
    While colMessageQueue.Count < 1 And getTimer
        DoEvents
    Wend
    If getTimer Then
        Dim strServerLine As String
        strServerLine = colMessageQueue(1)
        colMessageQueue.Remove 1
        
        If UCase(Left(strServerLine, 3)) <> "+OK" Then
            Err.Raise errServerError, "PopClient", strServerLine
        Else
            If Len(strServerLine) > 4 Then
                getOnelineResponse = Right$(strServerLine, Len(strServerLine) - 3)
            Else
                getOnelineResponse = ""
            End If
        End If
    Else
        Err.Raise errTimeout, "PopClient", "Timed out waiting for server"
    End If
End Function

Private Function getMultilineResponse() As String
    'Gets a multiline response from the server, starting with
    '+OK (not passed on) and ending with a single . on a line
    '(not passed on)
    setTimer
    While colMessageQueue.Count < 1 And getTimer
        DoEvents
    Wend
    If getTimer Then
        Dim strServerLine As String
        strServerLine = colMessageQueue(1)
        colMessageQueue.Remove 1
        
        If UCase(Left(strServerLine, 3)) <> "+OK" Then
            Err.Raise errServerError, "PopClient", strServerLine
        Else
            Do While strServerLine <> "."
                setTimer
                While colMessageQueue.Count < 1 And getTimer
                    DoEvents
                Wend
                If Not getTimer Then
                    Err.Raise errTimeout, "PopClient", "Timed out waiting for server"
                End If
                
                strServerLine = colMessageQueue(1)
                colMessageQueue.Remove 1
                
                If Left(strServerLine, 2) = ".." Then
                    getMultilineResponse = getMultilineResponse + Right(strServerLine, Len(strServerLine) - 1) + vbCrLf
                ElseIf strServerLine <> "." Then
                    getMultilineResponse = getMultilineResponse + strServerLine + vbCrLf
                End If
            Loop
        End If
    Else
        Err.Raise errTimeout, "PopClient", "Timed out waiting for server"
    End If
End Function

Private Function connected() As Boolean
    If pWsConnection.State = 7 Then
        connected = True
    Else
        connected = False
    End If
End Function

Private Function removeChr(strData As String, bChr As Byte) As String
    For a = 1 To Len(strData)
        If Mid(strData, a, 1) <> Chr(bChr) Then
            removeChr = removeChr + Mid(strData, a, 1)
        End If
    Next a
End Function

Private Function strToMessage(strMessage As String) As Message
    Dim strHeaderName As String
    Dim strHeaderValue As String
    Dim colMessageHeader As Collection
    Dim lngHeaderEnd As Long
    
    lngHeaderEnd = InStr(1, strMessage, vbCrLf + vbCrLf)
    If lngHeaderEnd = 0 Then
        lngHeaderEnd = InStr(1, strMessage, vbCr + vbCr)
    End If
    
    If lngHeaderEnd > 0 Then
        Set strToMessage = New Message
        strToMessage.MessageText = Trim(Right(strMessage, Len(strMessage) - lngHeaderEnd))
        Set colMessageHeader = tokeniseString(removeChr(Left(strMessage, lngHeaderEnd), 10), 13)
        
        For a = 1 To colMessageHeader.Count
            If InStr(1, colMessageHeader(a), ":") <> 0 And InStr(1, colMessageHeader(a), ":") <> Len(Trim(colMessageHeader(a))) Then
                strHeaderName = Left(colMessageHeader(a), InStr(1, colMessageHeader(a), ":") - 1)
                strHeaderValue = Trim(Right(colMessageHeader(a), Len(colMessageHeader(a)) - InStr(1, colMessageHeader(a), ":")))
                Select Case UCase(strHeaderName)
                Case "RETURN-PATH"
                    strToMessage.ReturnPath = strHeaderValue
                Case "MESSAGE-ID"
                    strToMessage.MessageId = strHeaderValue
                Case "SUBJECT"
                    strToMessage.Subject = strHeaderValue
                Case "TO"
                    strToMessage.MessageTo = strHeaderValue
                Case "CC"
                    strToMessage.MessageCC = strHeaderValue
                Case "FROM"
                    strToMessage.MessageFrom = strHeaderValue
                Case "DATE"
                    strToMessage.MessageDate = strHeaderValue
                Case "CONTENT-TYPE"
                    strToMessage.ContentType = strHeaderValue
                End Select
                On Error Resume Next
                If UCase(Left(strHeaderName, 2)) = "X-" Then
                    strToMessage.XHeaders.Add strHeaderName + ":" + strHeaderValue, strHeaderName
                End If
                On Error GoTo 0
            End If
        Next a
    Else
        Err.Raise errInvalidMessage, "PopClient", "The server returned an invalid message"
    End If
End Function

Private Function sendStr(strData As String)
    If debugging Then
        Form1.Text1.Text = Form1.Text1.Text + strData
    End If
    pWsConnection.SendData strData
End Function

Private Function tokeniseString(strData As String, token As Byte) As Collection
    'Takes a string and returns a collection of strings, one
    'for each 'token', or item seperated by a particular character
    Dim currentToken As String
    Set tokeniseString = New Collection
    For a = 1 To Len(strData)
        If Mid(strData, a, 1) = Chr(token) Then
            If currentToken <> "" Then
                tokeniseString.Add currentToken
                currentToken = ""
            End If
        Else
            currentToken = currentToken + Mid(strData, a, 1)
        End If
    Next a
    If currentToken <> "" Then tokeniseString.Add currentToken
End Function

