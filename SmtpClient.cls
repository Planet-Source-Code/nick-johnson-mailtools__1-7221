VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SmtpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class moudle is designed to encapsulate for VB the neccessary
'function calls to a SMTP server, and interpret the responses.
'It still requires a basic knowledge of how to send messages
'via SMTP, but should go something like this:

'Give the SmtpClient Object a Winsock Object
'Set the ServerName
'Set the ServerPort
'Connect
'HELO <domain>
'MAIL <sender> - This will appear as the return address
'RCPT <recepient> - The actual person who recieves it, not
'                   neccessarially what is seen on the TO:
'                   line.
'Repeat RCPT for each recepient.
'DATA <message object> - The actual message. The only headers
'                        that are used are To:, CC:, From:,
'                        Subject:, Content-Type, and XHeaders
'                        Note that the To: and CC: headers
'                        do not acutually affect the recepient.
'Repeat from MAIL for each message to be sent
'QUIT
'Demands for more features? Comments, questions, offers of $?
'Tell me! arachnid@mad.scientist.com
'To really impress me, use this module to send me the message!

'local variable(s) to hold property value(s)
Private WithEvents pWsConnection As Winsock 'local copy
Attribute pWsConnection.VB_VarHelpID = -1
Private pTimeout As Integer 'local copy

Public Enum Reciept
    None = 0
    Success = 1
    Failure = 2
    Delay = 4
End Enum

'Local variables
Private strServerResponse As String
Private colMessageQueue As New Collection 'Not really needed, but better safe than sorry
Private sngTimer As Single
Private a As Integer 'Loop counter

'local error code constants
Private Const errNotConnected = 1 + vbObjectError
Private Const errTimeout = 2 + vbObjectError
Private Const errNoWinsock = 3 + vbObjectError
Private Const errServerError = 4 + vbObjectError
Private Const errInvalidServerResponse = 5 + vbObjectError
Private Const errInvalidMessage = 6 + vbObjectError

'local SMTP reply constants
Private Const smtpStatus = 211
Private Const smtpHelp = 214
Private Const smtpReady = 220
Private Const smtpClosing = 221
Private Const smtpDone = 250
Private Const smtpWillForward = 251
Private Const smtpStartMail = 354
Private Const smtpShuttingDown = 421
Private Const smtpMailboxUnavailable = 450
Private Const smtpLocalError = 451
Private Const smtpNoSpace = 452
Private Const smtpSyntaxError = 500
Private Const smtpArgError = 501
Private Const smtpNoCommand = 502
Private Const smtpBadSequence = 503
Private Const smtpNoParamater = 504
Private Const smtpMailboxUnavailable2 = 550
Private Const smtpUserRejected = 551
Private Const smtpTooBig = 552
Private Const smtpInvalidMailboxName = 553
Private Const smtpFailed = 554

Private Const debugging = True

Private Function sendStr(strData As String)
    If debugging Then
        Form1.Text1.Text = Form1.Text1.Text + strData
    End If
    pWsConnection.SendData strData
End Function

Private Function connected() As Boolean
    If pWsConnection.State = 7 Then
        connected = True
    Else
        connected = False
    End If
End Function

Private Sub getReply(intExpectedCode As Integer, Optional intOtherCode As Integer = -1)
    'Waits for a server reply and checks to see if it is
    'what was expected.
    setTimer
    While colMessageQueue.Count < 1 And getTimer
        DoEvents
    Wend
    If getTimer Then
        Dim strServerLine As String
        strServerLine = colMessageQueue(1)
        colMessageQueue.Remove 1
        If Val(Left(strServerLine, 3)) <> intExpectedCode Then
            If Val(Left(strServerLine, 3)) <> intOtherCode Then
                Err.Raise errServerError, "SmtpClient", strServerLine
            End If
        End If
    Else
        Err.Raise errTimeout, "SmtpClient", "Timed out waiting for server"
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

Private Function removeChr(strData As String, bChr As Byte) As String
    For a = 1 To Len(strData)
        If Mid(strData, a, 1) <> Chr(bChr) Then
            removeChr = removeChr + Mid(strData, a, 1)
        End If
    Next a
End Function

Public Property Let serverPort(ByVal vData As Integer)
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
    Set pWsConnection = vData
End Property
Public Property Get wsConnection() As Winsock
    Set pWsConnection = pWsConnection
End Property

Public Sub Connect()
    'Connects to the SMTP server waits for a 220 response
    Dim strServerResponse As String
    pWsConnection.Connect
    getReply smtpReady
End Sub

Public Sub Helo(strDomain As String)
    'Says 'Hello' and identifies itself to the SMTP server
    sendStr "HELO " + strDomain + vbCrLf
    getReply smtpDone
End Sub

Public Sub Mail(strSender As String)
    'Starts a new message from <strSender>
    sendStr "MAIL FROM:" + strSender + vbCrLf
    getReply smtpDone
End Sub

Public Sub Rcpt(strRecepient As String, bReciept As Byte)
    'Adds a recepient to the current message. Can be done
    'any number of times for each message.
    'bReciept should contain values from the public enum
    'Reciept, indicating when to notify the sender.
    Dim strOptions As String
    If bReciept = 0 Then
        strOptions = "NOTIFY=NONE"
    End If
    If bReciept And 1 Then
        strOptions = "NOTIFY=SUCCESS"
    End If
    If bReciept And 2 Then
        strOptions = strOptions + " NOTIFY=FAILURE"
    End If
    If bReciept And 4 Then
        strOptions = strOptions + " NOTIFY=DELAY"
    End If
    sendStr "RCPT TO:" + strRecepient + " " + strOptions + vbCrLf
    getReply smtpDone, smtpWillForward
End Sub

Public Sub Data(msgMessage As Message)
    Dim colMessageText As Collection
    Dim strCurrentLine As String
    
    sendStr "DATA" + vbCrLf
    getReply smtpStartMail
    sendStr "From:" + msgMessage.MessageFrom + vbCrLf
    sendStr "To:" + msgMessage.MessageTo + vbCrLf
    sendStr "Cc:" + msgMessage.MessageCC + vbCrLf
    sendStr "Subject:" + msgMessage.Subject + vbCrLf
    sendStr "Content-Type:" + msgMessage.ContentType + vbCrLf
    For a = 1 To msgMessage.XHeaders.Count
        sendStr msgMessage.XHeaders(a) + vbCrLf
    Next a
    sendStr vbCrLf 'end of header, beginning of message

    Set colMessageText = tokeniseString(removeChr(msgMessage.MessageText, 10), 13)
    For a = 1 To colMessageText.Count
        strCurrentLine = colMessageText(a)
        If Left(strCurrentLine, 1) = "." Then strCurrentLine = strCurrentLine + "."
        sendStr strCurrentLine + vbCrLf
    Next a
    sendStr "." + vbCrLf
    getReply smtpDone
End Sub

Public Sub Reset()
    'Resets the server, cancelling any messages underway.
    sendStr "RSET" + vbCrLf
    getReply smtpDone
End Sub

Public Sub Noop()
    'Sends a NOOP to the server, to keep the connection alive.
    sendStr "NOOP" + vbCrLf
    getReply smtpDone
End Sub

Public Sub Quit()
    'Tells the server to disconnect, then kills the
    'connection from our end.
    sendStr "QUIT" + vbCrLf
    getReply smtpClosing
    pWsConnection.Close
End Sub
