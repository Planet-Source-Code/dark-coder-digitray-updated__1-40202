VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Begin VB.Form frmTray 
   BackColor       =   &H00000000&
   Caption         =   "Archon DigiTray"
   ClientHeight    =   2625
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7080
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin DigiTray.DownloadCtl DownloadCtl1 
      Left            =   3240
      Top             =   1080
      _ExtentX        =   1085
      _ExtentY        =   873
   End
   Begin VB.Timer Timer3 
      Interval        =   30000
      Left            =   4320
      Top             =   1080
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3240
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   3840
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   2760
      Top             =   1080
   End
   Begin MSComctlLib.ImageList ilWeather 
      Left            =   3240
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":116E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin DigiTray.HTTP HTTP 
      Left            =   120
      Top             =   -2220
      _ExtentX        =   900
      _ExtentY        =   900
      RemotePort      =   80
   End
   Begin DigiTray.TrayControl tc 
      Left            =   120
      Top             =   -2220
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private DontBlockCount As Long
Dim allowPops As Boolean, BlockHim As Boolean
Dim animateStep As Integer
Dim CTRLDown As Boolean
Private lMinHeight As Long
Private lMinWidth As Long
Private bResizeOff As Boolean
Private colMessages As String
Private Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hWnd As Long) As Long
      Private Type Quote
    Symbol As String
    TradeDate As String
    TradeTime As String
    Change As Double
    OpenPrice As Double
    HighPrice As Double
    LowPrice As Double
    LastPrice As Double
    Volume As Long
End Type

Private QuoteData As Quote


Private Declare Function Shell_NotifyIcon Lib "shell32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'constants required by Shell_NotifyIcon API call:
Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4
Const WM_MOUSEMOVE = &H200
Const WM_RBUTTONDBLCLK = &H206
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_MBUTTONDBLCLK = &H209
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private nid As NOTIFYICONDATA
Const TimeServerUrl = "http://tycho.usno.navy.mil/cgi-bin/timer.pl"
Public Sub GetAtomicTime() 'This sub retreives the raw time data file from the USNO atomic time server
    On Error GoTo ErrRtn 'Traps the "RequestTimeout" property of the Internet Transfer Control
    Dim tempData As String 'Holds the data received from the atomic time server
    
'    lblStatus.Caption = "Connecting to USNO Atomic Time server..." 'Updates the status label
    DoEvents
    'Clipboard.SetText TimeServerURL
    tempData = Inet1.OpenURL(TimeServerUrl) 'Request time data from USNO atomic time server
    Call SetAtomicTime(tempData) 'Call SetAtomicTime sub
    Exit Sub
    
ErrRtn: 'This routine is run only if the attempted network request fails
'    lblStatus.ForeColor = vbRed
'    lblStatus.Caption = "Network request failed!"
'    InProcess = False
'    Unload Me 'Exit program
End Sub
Private Sub SetAtomicTime(RawData As String) 'Extrapolates the UTC time from the raw data received from the USNO
                                             'atomic time server, and sets the local system's time to the time-zone
                                             'adjusted UTC atomic time
                                             
    Dim X As Integer 'Holds found character positions
    Dim Y As Integer 'Holds found character positions
    Dim tempTime As Variant 'Holds the extrapolated UTC and adjusted times
    
    X = InStr(1, RawData, GetSetting(App.Title, "Clock", "Zone", "Pacific")) 'Find "Universal" in the raw data ("Universal" indicates UTC time)
    If X > 0 Then 'If "Universal" was found in the raw data
        tempTime = Left$(RawData, X) 'Set "tempTime" equal to the section of the raw data we're interested in
        Y = InStrRev(tempTime, ",") 'Find the first comma in the tempTime data, starting from the back
        If Y > 0 Then 'If a comma was found in the "tempTime" data
            tempTime = CDate(Trim(Mid$(RawData, Y + 1, (X - (Y + 1))))) 'Cast the "tempTime" variable into a date containing the extracted actual UTC time
            Time = tempTime ' - AdjustTimeForTimeZone 'Set the local system time to the time-zone adjusted UTC atomic time
            'lblStatus.ForeColor = RGB(127, 255, 127) 'Change the status label's forecolor to light green
            'lblStatus.Caption = "Your system time has been changed to: " & Time & "..." 'Update the status label
            'InProcess = False
            DoEvents
            'Unload Me 'Exit the program
        Else 'If no comma was found in the "tempTime" data
            'lblStatus.ForeColor = vbRed
            'lblStatus.Caption = "Received bad data!"
            'InProcess = False
            'DoEvents
            'Unload Me 'Exit the program
        End If
    Else 'If "Universal" was not found in the raw data
        'lblStatus.ForeColor = vbRed
        'lblStatus.Caption = "Received bad data!"
        'InProcess = False
        'DoEvents
        'Unload Me 'Exit the program
    End If
End Sub


Public Function WeatherIcon()
On Error Resume Next
HTTP.DownloadHTML "http://www.uspntech.com/cgi-bin/apexec.pl?template=freeweather.htm&etype=weather&search=" & GetSetting(App.Title, "Weather", "Zipcode", "")
End Function


Private Sub DownloadCtl1_DownloadComplete(TempFileName As String)
   
    'Store data in a Quote UDT
    Dim Record As String
    Dim Found As Integer
    Dim fNum As Long
    
    On Error GoTo ErrHndl
    
    fNum = FreeFile
    Open TempFileName For Input Lock Read As #fNum
        
    Line Input #fNum, Record 'Parse record
    Record = Replace(Record, """", "")
    
    If InStr(Record, "N/A") Then 'Invalid symbol
        QuoteData.TradeDate = "N/A"
        Exit Sub
    End If
                
    'Symbol
    Found = InStr(Record, ",")
    If Found <> 0 Then
        QuoteData.Symbol = Mid$(Record, 1, Found - 1)
    End If
    Record = Mid$(Record, Found + 1)
                
    'Last price
    Found = InStr(Record, ",")
    If Found <> 0 Then
        QuoteData.LastPrice = CDbl(Mid$(Record, 1, Found - 1))
    End If
    Record = Mid$(Record, Found + 1)
                
    'Trade date
    Found = InStr(Record, ",")
    If Found <> 0 Then
        QuoteData.TradeDate = Mid$(Record, 1, Found - 1)
    End If
    Record = Mid$(Record, Found + 1)
                
    'Trade time
    Found = InStr(Record, ",")
    If Found <> 0 Then
        QuoteData.TradeTime = Mid$(Record, 1, Found - 1)
    End If
    Record = Mid$(Record, Found + 1)
                
    'Change
    Found = InStr(Record, ",")
    If Found <> 0 Then
        QuoteData.Change = CDbl(Mid$(Record, 1, Found - 1))
    End If
    Record = Mid$(Record, Found + 1)
                
    'Open price
    Found = InStr(Record, ",")
    If Found <> 0 Then
        QuoteData.OpenPrice = CDbl(Mid$(Record, 1, Found - 1))
    End If
    Record = Mid$(Record, Found + 1)
                
    'High price
    Found = InStr(Record, ",")
    If Found <> 0 Then
        QuoteData.HighPrice = CDbl(Mid$(Record, 1, Found - 1))
    End If
    Record = Mid$(Record, Found + 1)
                
    'Low price
    Found = InStr(Record, ",")
    If Found <> 0 Then
        QuoteData.LowPrice = CDbl(Mid$(Record, 1, Found - 1))
    End If
    Record = Mid$(Record, Found + 1)
                
    'Volume
    QuoteData.Volume = CDbl(Record)
          
    Close #fNum
    
    'Remove temp file
    Kill TempFileName

    Exit Sub
    
ErrHndl:
    
    MsgBox Err.Description, vbExclamation, "Error: " & Err.Number
End Sub

Private Sub Form_Load()
tc.AddIcon Me, 0, Me.Icon, "Syncronizing... - DigiTray Atomic Clock"
tc.AddIcon frmWeather, 1, ilWeather.ListImages(6).ExtractIcon, "Unknown - DigiTray Weather"
tc.AddIcon frmStock, 2, frmStock.Icon, "StockQuoter - DigiTray Stocks"
WeatherIcon
GetAtomicTime
Timer3_Timer
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
   Dim msg As Long
       
   'really interesting stuff here...i got it from MSDN
   If Me.ScaleMode = vbPixels Then
      msg = X
   Else
      msg = X / Screen.TwipsPerPixelX
   End If

   'handles mouse events when form is minimized, hidden and icon is in the system tray
   Select Case msg
      Case 513

          frmAtomic.Show
      Case WM_RBUTTONDOWN
      Case WM_RBUTTONUP
         PopupMenu mnuTray
      Case WM_LBUTTONDBLCLK
        'UpdateIcon NIM_DELETE
        'bResizeOff = True
        'Me.WindowState = vbNormal
        'Result = SetForegroundWindow(Me.hWnd)
        'Me.Show
        'bResizeOff = False
        'Me.Tag = ""
      Case WM_LBUTTONDOWN
      Case WM_LBUTTONUP
      Case WM_MBUTTONDBLCLK
      Case WM_MBUTTONDOWN
      Case WM_MBUTTONUP
      Case WM_MOUSEMOVE
      Case Else
   End Select
End Sub


Private Sub HTTP_Done(HTML As String)
'tc.modIcon Me, 1, ilWeather.ListImages(6).ExtractIcon, "Current Temperature: ?"
TranslateWeather HTML
End Sub

Public Sub TranslateWeather(HTML As String)
Dim Weather As String, Weathertemp As String
On Error Resume Next
X = InStr(X + 1, HTML, ";F</font></div></td><td>&nbsp;&nbsp;</td>") + Len(";F</font></div></td><td>&nbsp;&nbsp;</td>")
'X = InStr(X + 1, HTML, ";F</font></div></td><td>&nbsp;&nbsp;</td>") + Len(";F</font></div></td><td>&nbsp;&nbsp;</td>")
'X = InStr(X + 1, HTML, ";F</font></div></td><td>&nbsp;&nbsp;</td>") + Len(";F</font></div></td><td>&nbsp;&nbsp;</td>")
Y = InStr(X, HTML, "</td>")
Weather = Replace(Replace(Replace(Replace(Mid(HTML, X, Y - (X)), "<td>", ""), vbCrLf, ""), Chr(10), ""), Chr(9), "")
If InStr(1, UCase(Weather), "RAIN") Or InStr(1, Weather, "DRIZZLE") Then
frmWeather.Icon = ilWeather.ListImages(5).ExtractIcon
GoTo 1
End If
If InStr(1, UCase(Weather), "CLOUDY") Or InStr(1, UCase(Weather), "CLOUDS") Then
frmWeather.Icon = ilWeather.ListImages(2).ExtractIcon
GoTo 1
End If
If InStr(1, UCase(Weather), "SUNNY") Then
frmWeather.Icon = ilWeather.ListImages(1).ExtractIcon
GoTo 1
End If
If InStr(1, UCase(Weather), "STORM") Then
frmWeather.Icon = ilWeather.ListImages(3).ExtractIcon
GoTo 1
End If
frmWeather.Icon = ilWeather.ListImages(6).ExtractIcon
1:
If InStr(1, Weather, "Highs") Then
j = InStr(1, Weather, "Highs")
Else
j = InStr(1, Weather, "Lows")
End If
K = Len(Weather)
Weathertemp = Mid(Weather, j, K - j)
If Weathertemp = "" Then Weathertemp = "ZipCode Invalid"
frmWeather.txtWeather = Weather
tc.modIcon frmWeather, 1, frmWeather.Icon, Weathertemp & " - DigiTray Weather"
End Sub

Private Sub mnuTraySettingsStocks_Click()
SaveSetting App.Title, "Stockmarket", "Quote", InputBox("Enter a quote please", "MSFT")
End Sub

Private Sub mnuTraySettingsTime_Click()
frmAtomic.Show
End Sub

Private Sub mnuTraySettingsWeather_Click()
frmWeather.Show
End Sub

Private Sub mnuExit_Click()
tc.delIcon 0
tc.delIcon 1
tc.delIcon 2
End
End Sub

Private Sub Timer1_Timer()
WeatherIcon
End Sub


Private Sub Timer2_Timer()
GetAtomicTime
End Sub


Private Sub Timer3_Timer()
   txtSymbol = "KKD"
    Dim Q As Quote
    On Error Resume Next

    Q = GetQuote(GetSetting(App.Title, "Stockmarket", "Quote", "MSFT"))
    If Not Q.LastPrice = LastPrice Then
    StockUpdate GetSetting(App.Title, "Stockmarket", "Quote", "MSFT"), GetSetting(App.Title, "Stockmarket", "Quote", "MSFT") & " last price at " & _
                        Q.TradeTime & " was " & Q.LastPrice & _
                        " with volume of " & Q.Volume, 0
    'That's it!
    End If
    LastPrice = Q.LastPrice
End Sub

Private Function GetQuote(Symbol As String) As Quote
    'Downloads data from finance.yahoo.com
    QuoteData.TradeDate = ""
    DownloadCtl1.BeginDownload "http://finance.yahoo.com/d/quotes.csv?s=" & _
                               UCase(Symbol) & "&f=sl1d1t1c1ohgv&e=.csv"
    Do While QuoteData.TradeDate = ""
        DoEvents
    Loop
    GetQuote = QuoteData
End Function

