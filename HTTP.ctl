VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.UserControl HTTP 
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   InvisibleAtRuntime=   -1  'True
   Picture         =   "HTTP.ctx":0000
   ScaleHeight     =   510
   ScaleWidth      =   510
   ToolboxBitmap   =   "HTTP.ctx":030A
   Begin MSWinsockLib.Winsock WS 
      Left            =   1560
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
End
Attribute VB_Name = "HTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_URL = ""
'Property Variables:
Dim m_URL As String
'Event Declarations:
Event Done(HTML As String)
Private RemoteHost As String

Private mstrURL As String
Private mstrResponseDocument As String
Private mblnIsProxyUsed As Boolean
Private mblnIsPicture As Boolean
Private mblnIsHeader As Boolean
Private mstrReturnHeader As String
Private mstrRequestHeader As String
Private mstrLocalFile As String
Event Connected(IPHost As String)
Private HTML As String
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get URL() As String
Attribute URL.VB_Description = "URL for Downloading HTML Page"
    URL = m_URL
End Property

Public Property Let URL(ByVal New_URL As String)
    m_URL = New_URL
    PropertyChanged "URL"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=WS,WS,-1,RemotePort
Public Property Get RemotePort() As Long
Attribute RemotePort.VB_Description = "Returns/Sets the port to be connected to on the remote computer"
    RemotePort = WS.RemotePort
End Property

Public Property Let RemotePort(ByVal New_RemotePort As Long)
    WS.RemotePort() = New_RemotePort
    PropertyChanged "RemotePort"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function DownloadHTML(Optional URL = "http://www.pscode.com/index.html", Optional Port = 80) As Variant






    Dim strPureURL As String
    Dim strServerAddress As String
    Dim strServerHostIP As String
    Dim strDocumentURI As String
    Dim lngStartPos As Long
    Dim lngServerPort As Long
    
    Dim strRequestTemplate As String
     
     mstrURL = URL
             
'    If (optProxy.Value = True) Then
'        mblnIsProxyUsed = True
'    End If

    
    If UCase(Left(mstrURL, 7)) <> "HTTP://" Then
        MsgBox "Please enter url With http://", vbCritical + vbOK
        Exit Function
    End If
    
    ' Note: This section of code (header) is based on code posted
    ' by Tair Abdurman on http://www.planetsourcecode.com
    ' - Thanks for the proxy help Tair
    mstrRequestHeader = ""
    strRequestTemplate = "GET _$-$_$- HTTP/1.0" & Chr(13) & Chr(10) & _
    "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-comet, */*" & Chr(13) & Chr(10) & _
    "Accept-Language: en" & Chr(13) & Chr(10) & _
    "Accept-Encoding: gzip , deflate" & Chr(13) & Chr(10) & _
    "Cache-Control: no-cache" & Chr(13) & Chr(10) & _
    "Proxy-Connection: Keep-Alive" & Chr(13) & Chr(10) & _
    "User-Agent: SSM Agent 1.0" & Chr(13) & Chr(10) & _
    "Host: @$@@$@" & Chr(13) & Chr(10)
    
    ' Remove "http://"
    strPureURL = Right(mstrURL, Len(mstrURL) - 7)
    lngStartPos = InStr(1, strPureURL, "/")
    
    If lngStartPos < 1 Then
        strServerAddress = strPureURL
        strDocumentURI = "/"
    Else
        strServerAddress = Left(strPureURL, lngStartPos - 1)
        strDocumentURI = Right(strPureURL, Len(strPureURL) - lngStartPos + 1)
        mstrLocalFile = App.Path & "\" & Right(strPureURL, Len(strPureURL) - InStrRev(strPureURL, "/"))
    End If
            
    If strServerAddress = "" Or strDocumentURI = "" Then
        'msgbox "Unable To detect target page!", vbCritical + vbOK
        Exit Function
    End If
            
    If mblnIsProxyUsed Then
        strServerHostIP = txtProxy.Text
        mstrRequestHeader = strRequestTemplate
        mstrRequestHeader = Replace(mstrRequestHeader, "_$-$_$-", mstrURL)
        lngServerPort = 80
    Else
        strServerHostIP = strServerAddress
        lngServerPort = 80
        mstrRequestHeader = strRequestTemplate
        mstrRequestHeader = Replace(mstrRequestHeader, "_$-$_$-", strDocumentURI)
    End If
            
    mstrRequestHeader = Replace(mstrRequestHeader, "@$@@$@", strServerAddress)
    mstrRequestHeader = mstrRequestHeader & Chr(13) & Chr(10)
'    txtStatus.Text = "Connecting To server ..." & vbCrLf
'    txtStatus.Refresh
    
    ' Are we retreiving a picture
    If (UCase(Right(mstrURL, 3)) = "GIF" Or _
        UCase(Right(mstrURL, 3)) = "JPG") Then
        mblnIsPicture = True
        On Error Resume Next
        Kill mstrLocalFile
        
        ' Open mstrLocalFile For Binary As #1
        Open mstrLocalFile For Binary Access Write As #1
        mblnIsHeader = True
    Else
        mblnIsHeader = False
        mblnIsPicture = False
    End If
           ''msgbox mstrRequestHeader
    WS.Connect strServerHostIP, Port






End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_URL = m_def_URL
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_URL = PropBag.ReadProperty("URL", m_def_URL)
    WS.RemotePort = PropBag.ReadProperty("RemotePort", 0)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Width = 510
Height = 510
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("URL", m_URL, m_def_URL)
    Call PropBag.WriteProperty("RemotePort", WS.RemotePort, 0)
End Sub

Private Sub WS_Close()
WS.Close

RaiseEvent Done(Mid(HTML, InStr(1, UCase(HTML), "<"), Len(HTML) - InStr(1, UCase(HTML), "<")))
HTML = ""
End Sub

Private Sub WS_Connect()
    HTML = ""
    'msgbox mstrRequestHeader
    WS.Tag = ""
    WS.SendData mstrRequestHeader
    RaiseEvent Connected(WS.RemoteHostIP)
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
   Dim strTemp As String
    Dim lngBytes As Long
    Dim blnFoundHeadEndByte As Boolean
    Dim b() As Byte
    Dim b2() As Byte
    Dim aryMyArray As Variant
    Dim i As Long
    Dim j As Long
    Dim strChr As String
    
    WS.GetData strTemp, vbString
    HTML = HTML & strTemp
    
    
   

End Sub

Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
WS.Close
End Sub


