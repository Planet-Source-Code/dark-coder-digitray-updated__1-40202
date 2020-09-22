VERSION 5.00
Begin VB.UserControl TrayControl 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   82
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   80
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   0
      Top             =   80
      Width           =   270
   End
End
Attribute VB_Name = "TrayControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_TipText As String             'text to be displayed on system tray icon
Private Const def_TipText = ""


Public frm As Form
Public IconObject As Object
Public lngPrevWndProc As Long 'Original WNDPROC address.
Public lngWndID As Long 'Our unique icon identifier.
Public lngHwnd As Long 'The hwnd of frmTray.
Private Notify As NOTIFYICONDATA
Private BarData As APPBARDATA

Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GWL_WNDPROC = (-4)
Private Const IDANI_OPEN = &H1
Private Const IDANI_CLOSE = &H2
Private Const IDANI_CAPTION = &H3
Private Const NIF_TIP = &H4
Private Const NIM_ADD = 0&
Private Const NIM_DELETE = 2&
Private Const NIM_MODIFY = 1&
Private Const NIF_ICON = 2&
Private Const NIF_MESSAGE = 1&
Private Const ABM_GETTASKBARPOS = &H5&
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_USER = &H400

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, _
    ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, _
    lprcTo As RECT) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

Public Enum ZoomTypes
    ZOOM_FROM_TRAY
    ZOOM_TO_TRAY
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Private Type APPBARDATA
        cbSize As Long
        hWnd As Long
        uCallbackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long
End Type

Public Function SendToTray()

    'Create effect where form minimizes into the tray
    Dim lngRetVal As Long

    ZoomForm ZOOM_TO_TRAY, frm.hWnd
    frm.Visible = False 'hide the form from view
    Picture2.Picture = frm.Icon 'store original icon from restoration on terminate
    
    'take the specified initial image
    Set IconObject = frm.Icon
    
    'create the new icon on the system tray
    AddIcon frm, IconObject.Handle, IconObject, m_TipText
    
End Function

Public Function RestoreFromTray()

    'Create the effect that original window is expanding from system tray
    On Error Resume Next
    delIcon IconObject.Handle   'remove icon from tray
    frm.Icon = Picture2.Picture 'restore original icon
    ZoomForm ZOOM_FROM_TRAY, frm.hWnd
    frm.Visible = True          'make original form visible
    
End Function

Public Property Get TipText() As String

    TipText = m_TipText

End Property

Public Property Let TipText(ByVal New_TipText As String)
    
    m_TipText = New_TipText
    PropertyChanged "TipText"

End Property

Private Sub UserControl_InitProperties()
    
    'Initialize Properties for User Control
    m_TipText = def_TipText

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    'Load property values from storage
    Set frm = Parent
    m_TipText = PropBag.ReadProperty("TipText", def_TipText)
    
End Sub

Private Sub UserControl_Resize()
    
    With UserControl
        .Height = 450
        .Width = 450
    End With

    Line (0, 0)-(ScaleWidth, 0), vb3DHighlight 'Lightest Shadow
    Line (2, 2)-(ScaleWidth - 2, 2), vb3DDKShadow 'Darkest Shadow
    Line (0, 0)-(0, ScaleHeight), vb3DHighlight 'Lightest Shadow
    Line (2, 2)-(2, ScaleHeight - 2), vb3DDKShadow 'Darkest Shadow
    Line (ScaleWidth - 1, 1)-(ScaleWidth - 1, ScaleHeight), vb3DDKShadow 'Darkest Shadow
    Line (ScaleWidth - 3, 3)-(ScaleWidth - 3, ScaleHeight - 3), vb3DHighlight 'Lightest Shadow
    Line (ScaleWidth - 1, ScaleHeight - 1)-(0, ScaleHeight - 1), vb3DDKShadow 'Darkest Shadow
    Line (ScaleWidth - 3, ScaleHeight - 3)-(1, ScaleHeight - 3), vb3DHighlight 'Lightest Shadow
    Refresh

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("TipText", m_TipText, def_TipText)
    
End Sub

Public Function ZoomForm(zoomToWhere As ZoomTypes, hWnd As Long) As Boolean
    
    'This function 'zooms' a window.
    Dim rctFrom As RECT
    Dim rctTo As RECT
    Dim lngTrayHand As Long
    Dim lngStartMenuHand As Long
    Dim lngChildHand As Long
    Dim strClass As String * 255
    Dim lngClassNameLen As Long
    Dim lngRetVal As Long

    'Select the type of zoom to do.
    Select Case zoomToWhere
        'Zoom the window into the tray.
        Case ZOOM_FROM_TRAY
            'Get the handle to the start menu.
            lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)

            'Get the handle to the first child window of the start menu.
            lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)

            'Loop through all siblings until we find the 'System Tray' (A.K.A. --> TrayNotifyWnd)
            Do
                
                lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))

                'If it is the tray then store the handle.
                If InStr(1, strClass, "TrayNotifyWnd") Then
                    lngTrayHand = lngChildHand
                    Exit Do
                End If
                'If we didn't find it, go to the next sibling.
                lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
            
            Loop

            'Get the RECT of  our form.
            lngRetVal = GetWindowRect(hWnd, rctFrom)

            'Get the RECT of the Tray.
            lngRetVal = GetWindowRect(lngTrayHand, rctTo)

            'Zoom from the tray to where our form is.
            lngRetVal = DrawAnimatedRects(frm.hWnd, IDANI_CLOSE Or IDANI_CAPTION, rctTo, rctFrom)

        Case ZOOM_TO_TRAY

            'Get the handle to the start menu.
            lngStartMenuHand = FindWindow("Shell_TrayWnd", vbNullString)

            'Get the handle to the first child window of the start menu.
            lngChildHand = GetWindow(lngStartMenuHand, GW_CHILD)

            'Loop through all siblings until we find the 'System Tray' (A.K.A. --> TrayNotifyWnd)
            Do
                
                lngClassNameLen = GetClassName(lngChildHand, strClass, Len(strClass))
                'If it is the tray then store the handle.
                If InStr(1, strClass, "TrayNotifyWnd") Then
                    lngTrayHand = lngChildHand
                    Exit Do
                End If
                'If we didn't find it, go to the next sibling.
                lngChildHand = GetWindow(lngChildHand, GW_HWNDNEXT)
            
            Loop
            'Get the RECT of  our form.
            lngRetVal = GetWindowRect(hWnd, rctFrom)

            'Get the RECT of the Tray.
            lngRetVal = GetWindowRect(lngTrayHand, rctTo)

            'Zoom from where our form is to the tray .
            lngRetVal = DrawAnimatedRects(frm.hWnd, IDANI_OPEN Or IDANI_CAPTION, rctFrom, rctTo)
    
    End Select

End Function

Public Sub modIcon(Form1 As Form, IconID As Long, Icon, ToolTip As String)
    
    'Modify an existing icon on the system tray
    Dim Result As Long
    Notify.cbSize = 88&
    Notify.hWnd = Form1.hWnd
    Notify.uID = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = ToolTip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_MODIFY, Notify)

End Sub

Public Sub AddIcon(Form1 As Form, IconID As Long, Icon As Object, ToolTip As String)
    
    'Create an icon on the system tray
    Dim Result As Long
    BarData.cbSize = 36&
    Result = SHAppBarMessage(ABM_GETTASKBARPOS, BarData)
    Notify.cbSize = 88&
    Notify.hWnd = Form1.hWnd
    Notify.uID = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = ToolTip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_ADD, Notify)

End Sub

Public Sub delIcon(IconID As Long)
    
    'Remove an icon from the system tray
    Dim Result As Long
    Notify.uID = IconID
    Result = Shell_NotifyIcon(NIM_DELETE, Notify)

End Sub
