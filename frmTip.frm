VERSION 5.00
Begin VB.Form frmBalloon 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   ControlBox      =   0   'False
   ForeColor       =   &H80000017&
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   275
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timAutoClose 
      Enabled         =   0   'False
      Left            =   3960
      Top             =   1200
   End
   Begin VB.Image imgX_Up 
      Height          =   240
      Left            =   4380
      Picture         =   "frmTip.frx":000C
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgX_Dn 
      Height          =   240
      Left            =   4320
      Picture         =   "frmTip.frx":034E
      Top             =   390
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4305
      TabIndex        =   2
      ToolTipText     =   "Close"
      Top             =   135
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgX 
      Height          =   240
      Left            =   4275
      Picture         =   "frmTip.frx":0690
      Top             =   90
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   1
      Left            =   1200
      Picture         =   "frmTip.frx":09D2
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   0
      Left            =   960
      Picture         =   "frmTip.frx":0F5C
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   2
      Left            =   1440
      Picture         =   "frmTip.frx":14E6
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   2
      Left            =   720
      Picture         =   "frmTip.frx":1A70
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "frmTip.frx":1FFA
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIconXP 
      Height          =   240
      Index           =   0
      Left            =   240
      Picture         =   "frmTip.frx":2584
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape shpBorder 
      BorderWidth     =   2
      Height          =   1800
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.Image imgDisplayIcon 
      Height          =   240
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "<Title>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lblText 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "<Caption>"
      ForeColor       =   &H80000017&
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   4410
   End
End
Attribute VB_Name = "frmBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'All variables must be declared

Dim XY() As POINTAPI

Dim sTahomaOrMsSansSerif As String

Private Declare Function CreateEllipticRgn Lib "gdi32" _
    (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long 'Used to round the corners of the form
    
Private Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As POINTAPI, ByVal nCount As Long, _
    ByVal nPolyFillMode As Long) As Long 'Used to round corners of form

'SetWindowRgn is used when setting the form's shape (rounded corners) so
'Windows knows what the window's region is. That's the area in the window
'where Windows permits drawing, and it won't show any part of the window
'that is outside the window region. hWnd is the handle of the window we're
'working with, hRgn is the region's handle, and bRedraw is the redraw flag.
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Sub RoundCorners()
Attribute RoundCorners.VB_Description = "Rounds the corners of the form via API to create the tooltip effect"
Dim hRgn   As Long
Dim lRes   As Long
Dim XY(55) As POINTAPI

With Me
    .ScaleMode = vbPixels
    mlWidth = Me.ScaleWidth
    mlHeight = Me.ScaleHeight

    'Top Left Corner
    XY(0).X = 0 * 2
    XY(0).Y = 12 * 2
    XY(1).X = 1 * 2
    XY(1).Y = 11 * 2
    XY(2).X = 1 * 2
    XY(2).Y = 10 * 2
    XY(3).X = 2 * 2
    XY(3).Y = 9 * 2
    XY(4).X = 2 * 2
    XY(4).Y = 8 * 2
    XY(5).X = 3 * 2
    XY(5).Y = 6 * 2
    XY(6).X = 4 * 2
    XY(6).Y = 5 * 2
    XY(7).X = 5 * 2
    XY(7).Y = 4 * 2
    XY(8).X = 6 * 2
    XY(8).Y = 3 * 2
    XY(9).X = 8 * 2
    XY(9).Y = 2 * 2
    XY(10).X = 9 * 2
    XY(10).Y = 2 * 2
    XY(11).X = 10 * 2
    XY(11).Y = 1 * 2
    XY(12).X = 11 * 2
    XY(12).Y = 1 * 2
    XY(13).X = 12 * 2
    XY(13).Y = 0 * 2

    'Top Right Corner
    XY(14).X = mlWidth - 12 * 2
    XY(14).Y = 0 * 2
    XY(15).X = mlWidth - 1 * 2
    XY(15).Y = 1 * 2
    XY(16).X = mlWidth - 10 * 2
    XY(16).Y = 1 * 2
    XY(17).X = mlWidth - 9 * 2
    XY(17).Y = 2 * 2
    XY(18).X = mlWidth - 8 * 2
    XY(18).Y = 2 * 2
    XY(19).X = mlWidth - 6 * 2
    XY(19).Y = 3 * 2
    XY(20).X = mlWidth - 5 * 2
    XY(20).Y = 4 * 2
    XY(21).X = mlWidth - 4 * 2
    XY(21).Y = 5 * 2
    XY(22).X = mlWidth - 3 * 2
    XY(22).Y = 6 * 2
    XY(23).X = mlWidth - 2 * 2
    XY(23).Y = 8 * 2
    XY(24).X = mlWidth - 2 * 2
    XY(24).Y = 9 * 2
    XY(25).X = mlWidth - 1 * 2
    XY(25).Y = 10 * 2
    XY(26).X = mlWidth - 1 * 2
    XY(26).Y = 11 * 2
    XY(27).X = mlWidth - 0 * 2
    XY(27).Y = 12 * 2

    'Bottom Right Corner
    XY(28).X = mlWidth - 0 * 2
    XY(28).Y = mlHeight - 12 * 2
    XY(29).X = mlWidth - 1 * 2
    XY(29).Y = mlHeight - 11 * 2
    XY(30).X = mlWidth - 1 * 2
    XY(30).Y = mlHeight - 10 * 2
    XY(31).X = mlWidth - 2 * 2
    XY(31).Y = mlHeight - 9 * 2
    XY(32).X = mlWidth - 2 * 2
    XY(32).Y = mlHeight - 8 * 2
    XY(33).X = mlWidth - 3 * 2
    XY(33).Y = mlHeight - 6 * 2
    XY(34).X = mlWidth - 4 * 2
    XY(34).Y = mlHeight - 5 * 2
    XY(35).X = mlWidth - 5 * 2
    XY(35).Y = mlHeight - 4 * 2
    XY(36).X = mlWidth - 6 * 2
    XY(36).Y = mlHeight - 3 * 2
    XY(37).X = mlWidth - 8 * 2
    XY(37).Y = mlHeight - 2 * 2
    XY(38).X = mlWidth - 9 * 2
    XY(38).Y = mlHeight - 2 * 2
    XY(39).X = mlWidth - 10 * 2
    XY(39).Y = mlHeight - 1 * 2
    XY(40).X = mlWidth - 11 * 2
    XY(40).Y = mlHeight - 1 * 2
    XY(41).X = mlWidth - 12 * 2
    XY(41).Y = mlHeight - 0 * 2

    'Bottom Left Corner
    XY(42).X = 12 * 2
    XY(42).Y = mlHeight - 0 * 2
    XY(43).X = 11 * 2
    XY(43).Y = mlHeight - 1 * 2
    XY(44).X = 10 * 2
    XY(44).Y = mlHeight - 1 * 2
    XY(45).X = 9 * 2
    XY(45).Y = mlHeight - 2 * 2
    XY(46).X = 8 * 2
    XY(46).Y = mlHeight - 2 * 2
    XY(47).X = 6 * 2
    XY(47).Y = mlHeight - 3 * 2
    XY(48).X = 5 * 2
    XY(48).Y = mlHeight - 4 * 2
    XY(49).X = 4 * 2
    XY(49).Y = mlHeight - 5 * 2
    XY(50).X = 3 * 2
    XY(50).Y = mlHeight - 6 * 2
    XY(51).X = 2 * 2
    XY(51).Y = mlHeight - 8 * 2
    XY(52).X = 2 * 2
    XY(52).Y = mlHeight - 9 * 2
    XY(53).X = 1 * 2
    XY(53).Y = mlHeight - 10 * 2
    XY(54).X = 1 * 2
    XY(54).Y = mlHeight - 11 * 2
    XY(55).X = 0 * 2
    XY(55).Y = mlHeight - 12 * 2

    'Pass in the address of the first point and
    'the number of points.

    hRgn = CreatePolygonRgn(XY(0), (UBound(XY) + 1), 2)
    lRes = SetWindowRgn(.hwnd, hRgn, True)
End With


'Resize the border to fit:
shpBorder.Height = Me.ScaleHeight
shpBorder.Width = Me.ScaleWidth

'This does make the border two (as opposed to one, as on the other sides)
'pixels thick on the right and bottom sides, but that can sort of look like
'a shadow and not ugly ... right? If we add +1 to the end of both statements
'above, it's only one pixel thick and looks good, except it won't completely
'cover the corners -- and we don't want that! In the future, I plan to pick
'at my form-shaping code to make it match the shape control better

End Sub
Private Sub Form_Click()
'Hide me after I'm clicked on
HideBalloon
End Sub
Private Sub Form_Load()
RoundCorners ' Round the corners of this form to make it look "tool-tippy"
End Sub
Private Sub Form_Resize()
lblText.Height = Me.ScaleHeight - lblText.Top - 10 'Resize the balloon's text label height
'to fit correctly, no matter what the size of the balloon -- since that can be changed
'The - 10 is to give it a little room on the bottom; without it, it would touch it
'without any space (between the end of the balloon's text label and the bottom)

'Do the same as before, now with the width:
lblText.Width = Me.ScaleWidth - 2 * lblText.Left

'Now, resize the title label's width to fit the balloon size:
lblTitle.Width = Me.ScaleWidth - 2 * lblTitle.Left

'Move the X button
lblX.Left = Me.ScaleWidth - (1.5 * lblX.Width) - 1
imgX.Left = Me.ScaleWidth - (1.5 * imgX.Width)
'lblX.Move (Me.ScaleWidth - lblX.Width) - 13, 5
'imgX.Move (Me.ScaleWidth - lblX.Width) - 15, 2
'imgX_Dn.Move (Me.ScaleWidth - lblX.Width) - 15, 2
'imgX_Up.Move (Me.ScaleWidth - lblX.Width) - 15, 2

RoundCorners ' Round the corners of this form to make it look "tool-tippy"

'Resize the border to fit:

shpBorder.Height = Me.ScaleHeight
shpBorder.Width = Me.ScaleWidth

'This does make the border two (as opposed to one, as on the other sides)
'pixels thick on the right and bottom sides, but that can sort of look like
'a shadow and not ugly ... right? If we add +1 to the end of both statements
'above, it's only one pixel thick and looks good, except it won't completely
'cover the corners -- and we don't want that! In the future, I plan to pick
'at my form-shaping code to make it match the shape control better
 
End Sub

Private Sub imgDisplayIcon_Click()
' Hide this balloon if I'm clicked
HideBalloon
End Sub

Private Sub imgX_Click()
HideBalloon
End Sub

Private Sub imgX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then imgX.Picture = imgX_Dn.Picture
End Sub

Private Sub imgX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then imgX.Picture = imgX_Up.Picture
End Sub

Private Sub lblText_Click()
'Hide me after I'm clicked on
HideBalloon
End Sub
Private Sub lblTitle_Click()
'Hide me after I'm clicked on
HideBalloon
End Sub
Public Sub SetBalloon(sTitle, sText, lPosX, lPosY _
    , Optional sIcon, Optional bShowClose As Boolean = False, _
    Optional lAutoCloseAfter = 0, Optional lHeight = 1620, _
     Optional lWidth = 4680, Optional sFont _
     = "Tahoma, if exists", Optional iFontSize = 8)

'Arguments for this Sub are explained below. What this Sub does is
'set the properties for the balloon to be displayed--text, title, etc.
'After setting the properties, you must show the
'balloon yourself by calling <form_name>.Show
'For example, if this "template" form is frmBalloon, you can create a new
'instance of frmBalloon by doing:
'   Dim frmMyBalloon as New frmBalloon
'and then calling frmMyBalloon.SetBalloon using the values you want, as in:
'   frmMyBalloon.SetBalloon "Sample Title", "Sample Text"
'and going on with the arguments as needed (see below and the declaration
'for this Sub above).
'Then, to show the balloon, call
'   frmMyBalloon.Show , , Me
'The Me argument is needed so that the balloon becomes an owner of
'the form it's called from. Otherwise, it'd be possible for the form to
'get in front of the balloon, still appear when its parent is minimized, etc.,
'which we don't want.

'Here's what the arguments for this Sub do:

'sTitle: The bold title to appear above the text on the balloon (Required)

'sText: Text of balloon (Required)

'lPosX and lPosY: The horizontal and vertical, respectively, positions to
'                 show the ballon at (Required)

'sIcon: The icon to be displayed on the balloon, similar to the messagebox's.
'       They're an "i", "x", or "!". (No question mark here; you can't ask
'       on a balloon, can you?) To specifiy, pass either "i", "x", or "!" as
'       the argument, e.g., SetBalloon("Title", "Text", "!" ...
'       For none, don't pass anything. And, they'll use the XP-style icons
'       by default; to use 9x-looking icons instead, specify "i9", "x9", or "!9"
'       Look at the "template" form (frmBalloon, in my example project) to see what
'       they look like; you should see the difference, but they're quite similar--
'       the XP ones just look more colorful and 3D-ish (Optional)

'bShowClose: Whether or not to show the "X" close button the user can
'            press to close the balloon. If there, click to close the
'            balloon; if it's not there (or if it is) clicking anywhere
'            in the balloon will close it. (Optional)

'lAutoCloseAfter: Specifies the amount of time (in milliseconds) after
'                 which to automatically close the balloon. Setting it
'                 to 0 will make it not automatically close.
'                 E.g., 10,000 is ten seconds. (Optional)

'lHeight and lWidth: The width and height that you want the balloon to have.
'                    It's optional, and it will default to a "normal" size.
'                    If you have a long message, increasing the height should
'                    be good, although you can increase the width if you want, too
'                    (Optional)
                     
'
'sFont: The font the text will appear in, defaulting to MS Sans Serif or
'       Tahoma. By default, it will automatically check to see if Tahoma
'       exists. If so, it will use it; if not, MS Sans Serif will be used.
'       Tahoma gives it a "new" look, but some early Windows 9x versions
'       may not have it. You can specify any font you want using this
'       argument, however. (Optional)


'Setting TITLE AND CAPTION on balloon:
Me.lblTitle.Caption = sTitle
Me.lblText.Caption = sText

'Setting the X AND Y POSITIONS:
Me.Top = lPosY
Me.Left = lPosX

'Setting the ICON:
'First, convert the case to all lower; that way, since all Select Case
'statements below use lowercase for identification
sIcon = LCase(sIcon)

Select Case sIcon
    Case "i": 'The "i" icon, XP-style (default)
        Me.imgDisplayIcon.Picture = Me.imgIconXP(0).Picture
        
    Case "i9": 'The "i" icon, 9x/Me-style
        imgDisplayIcon.Picture = imgIcon(0).Picture
        
    Case "x": 'The "x" icon, XP-style
        imgDisplayIcon.Picture = imgIconXP(1).Picture
        
    Case "x9": 'The "x" icon, 9x/Me-style
        imgDisplayIcon.Picture = imgIcon(1).Picture
        
    Case "!": 'The "!" icon, XP-style
        imgDisplayIcon.Picture = imgIconXP(2).Picture
        
    Case "!9": 'The "!" icon, 9x-style
        imgDisplayIcon.Picture = imgIcon(2).Picture
        
    Case Else: 'Use no icon
        Me.imgDisplayIcon.Visible = False
        Me.lblTitle.Left = imgDisplayIcon.Left 'Move title over so it looks right
End Select
        
'Showing/not showing THE X BUTTON:
If bShowClose = False Then ' Then don't show the X button
    Me.imgX.Visible = False
    Me.lblX.Visible = False
End If
If bShowClose = True Then ' Then make the X button visible
    Me.imgX.Visible = True
    Me.lblX.Visible = True
End If

'Enabling/disabling AUTO-CLOSE:
If lAutoCloseAfter = 0 Then ' Then we don't need to auto-close, so ...
    Me.timAutoClose.Enabled = False ' Just make sure the auto-close timer
                                    ' is disabled, since we shouldn't auto-close
Else    ' Then we DO need to auto-close
    Me.timAutoClose.Interval = lAutoCloseAfter ' Set timer's interval so it will
                                               ' auto-close at the right time, and...
    Me.timAutoClose.Enabled = True 'Enable the timer, so it will go off and auto-close
End If

'Setting HEIGHT AND WIDTH:
Me.Width = lWidth
Me.Height = lHeight
RoundCorners

'Setting the FONT AND FONT SIZE:
'If no font specified and using default value, then
If sFont = "Tahoma, if exists" Then
' Check to see if Tahoma exists; if not, use MS Sans Serif; the DoesTahomaExist
'function will return True or False, depending on if Tahoma exists
    If DoesTahomaExist = False Then
        sFont = "MS Sans Serif"
    Else
        sFont = "Tahoma"
    End If
End If

Me.Font = sFont
Me.lblText.Font = sFont
Me.lblTitle.Font = sFont
Me.lblText.FontSize = iFontSize
Me.lblTitle.FontSize = iFontSize

End Sub
Private Sub lblX_Click()
HideBalloon
End Sub
Private Sub lblX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then imgX.Picture = imgX_Dn.Picture
End Sub
Private Sub lblX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then imgX.Picture = imgX_Up.Picture
End Sub
Private Sub timAutoClose_Timer()
' This timer is used to automatically close the balloon, if needed,
' after the specified number of milliseconds

HideBalloon  'Calls HideBalloon(), which hides the balloon
End Sub
Public Sub HideBalloon()
'HideBalloon() is used to manually hide the balloon and by the
'balloon itself to hide itself when needed
Unload Me
End Sub
