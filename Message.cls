VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private pMessageCC As String 'local copy
Private pMessageText As String 'local copy
Private pReturnPath As String 'local copy
Private pMessageId As String 'local copy
Private pSubject As String 'local copy
Private pMessageTo As String 'local copy
Private pMessageFrom As String 'local copy
Private pMessageDate As String 'local copy
Private pContentType As String 'local copy
Private pXHeaders As Collection 'local copy

Public Property Set XHeaders(ByVal vData As Collection)
    Set pXHeaders = vData
End Property
Public Property Get XHeaders() As Collection
    Set XHeaders = pXHeaders
End Property

Public Property Let ContentType(ByVal vData As String)
    pContentType = vData
End Property
Public Property Get ContentType() As String
    ContentType = pContentType
End Property

Public Property Let MessageDate(ByVal vData As String)
    pMessageDate = vData
End Property
Public Property Get MessageDate() As String
    MessageDate = pMessageDate
End Property

Public Property Let MessageFrom(ByVal vData As String)
    pMessageFrom = vData
End Property
Public Property Get MessageFrom() As String
    MessageFrom = pMessageFrom
End Property

Public Property Let MessageTo(ByVal vData As String)
    pMessageTo = vData
End Property
Public Property Get MessageTo() As String
    MessageTo = pMessageTo
End Property

Public Property Let Subject(ByVal vData As String)
    pSubject = vData
End Property
Public Property Get Subject() As String
    Subject = pSubject
End Property

Public Property Let MessageId(ByVal vData As String)
    pMessageId = vData
End Property
Public Property Get MessageId() As String
    MessageId = pMessageId
End Property

Public Property Let ReturnPath(ByVal vData As String)
    pReturnPath = vData
End Property
Public Property Get ReturnPath() As String
    ReturnPath = pReturnPath
End Property

Public Property Let MessageText(ByVal vData As String)
    pMessageText = vData
End Property
Public Property Get MessageText() As String
    MessageText = pMessageText
End Property

Public Property Let MessageCC(ByVal vData As String)
    pMessageCC = vData
End Property
Public Property Get MessageCC() As String
    MessageCC = pMessageCC
End Property

Private Sub Class_Initialize()
    Set pXHeaders = New Collection
End Sub
