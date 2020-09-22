VERSION 5.00
Begin VB.Form frmStock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DigiTray Stock Quoter"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   Icon            =   "frmStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Quote Stock:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Sub Command1_Click()
SaveSetting App.Title, "StockMarket", "Quote", Text1
Hide
End Sub

Private Sub Command2_Click()
Hide
End Sub

Private Sub Form_Activate()
Text1 = GetSetting(App.Title, "StockMarket", "Quote", "MSFT")
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

          frmStock.Show
      Case WM_RBUTTONDOWN
      Case WM_RBUTTONUP
         'PopupMenu mnuTray
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
