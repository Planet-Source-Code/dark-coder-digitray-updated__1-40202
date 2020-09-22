VERSION 5.00
Begin VB.Form frmStocks 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   2640
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   250
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   80
      Picture         =   "frmStocks.frx":0000
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Quote Update : MSFT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ontop As New clsOnTop

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
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, _
ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
        lpRect As RECT) As Long 'Used for getting positions of objects/forms
                                'to place balloons correctly

Private Type RECT   'Also used to store values for positions of balloons
   Left As Long    'after using the API to determine where
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, _
        lpPoint As POINTAPI) As Long 'Also used for getting positions of
                                     'objects/forms we want to place the
                                     'balloons by
                                     
'Used to draw the ellipse on the form
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, _
    ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, _
    ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'Used to create the regiod around the form to shape it
Private Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
    ByVal RectY2 As Long, ByVal EllipseWidth As Long, _
    ByVal EllipseHeight As Long) As Long
                                     
Private mlWidth As Long
Private mlHeight As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type BalloonCoords 'Used to store X and Y coordinates of balloon
    X As Long 'after using API and math operations to figure exact
    Y As Long 'coordinates regarding where to place itself
End Type


Private Sub Command1_Click()
Hide
End Sub

Private Sub Form_Activate()
Ontop.MakeTopMost hWnd
Timer1.Enabled = True
Left = Screen.Width - Width
Top = Screen.Height - Height - 500
End Sub
Private Sub Form_Load()
RoundCorners
End Sub
Private Sub RoundCorners()
Exit Sub
Dim hRgn   As Long
Dim lRes   As Long
Dim XY(55) As POINTAPI

With Me
    .ScaleMode = vbPixels
    mlWidth = Me.ScaleWidth
    mlHeight = Me.ScaleHeight

    'Top Left Corner
    XY(0).X = 0
    XY(0).Y = 12
    XY(1).X = 1
    XY(1).Y = 11
    XY(2).X = 1
    XY(2).Y = 10
    XY(3).X = 2
    XY(3).Y = 9
    XY(4).X = 2
    XY(4).Y = 8
    XY(5).X = 3
    XY(5).Y = 6
    XY(6).X = 4
    XY(6).Y = 5
    XY(7).X = 5
    XY(7).Y = 4
    XY(8).X = 6
    XY(8).Y = 3
    XY(9).X = 8
    XY(9).Y = 2
    XY(10).X = 9
    XY(10).Y = 2
    XY(11).X = 10
    XY(11).Y = 1
    XY(12).X = 11
    XY(12).Y = 1
    XY(13).X = 12
    XY(13).Y = 0

    'Top Right Corner
    XY(14).X = mlWidth - 12
    XY(14).Y = 0
    XY(15).X = mlWidth - 1
    XY(15).Y = 1
    XY(16).X = mlWidth - 10
    XY(16).Y = 1
    XY(17).X = mlWidth - 9
    XY(17).Y = 2
    XY(18).X = mlWidth - 8
    XY(18).Y = 2
    XY(19).X = mlWidth - 6
    XY(19).Y = 3
    XY(20).X = mlWidth - 5
    XY(20).Y = 4
    XY(21).X = mlWidth - 4
    XY(21).Y = 5
    XY(22).X = mlWidth - 3
    XY(22).Y = 6
    XY(23).X = mlWidth - 2
    XY(23).Y = 8
    XY(24).X = mlWidth - 2
    XY(24).Y = 9
    XY(25).X = mlWidth - 1
    XY(25).Y = 10
    XY(26).X = mlWidth - 1
    XY(26).Y = 11
    XY(27).X = mlWidth - 0
    XY(27).Y = 12

    'Bottom Right Corner
    XY(28).X = mlWidth - 0
    XY(28).Y = mlHeight - 12
    XY(29).X = mlWidth - 1
    XY(29).Y = mlHeight - 11
    XY(30).X = mlWidth - 1
    XY(30).Y = mlHeight - 10
    XY(31).X = mlWidth - 2
    XY(31).Y = mlHeight - 9
    XY(32).X = mlWidth - 2
    XY(32).Y = mlHeight - 8
    XY(33).X = mlWidth - 3
    XY(33).Y = mlHeight - 6
    XY(34).X = mlWidth - 4
    XY(34).Y = mlHeight - 5
    XY(35).X = mlWidth - 5
    XY(35).Y = mlHeight - 4
    XY(36).X = mlWidth - 6
    XY(36).Y = mlHeight - 3
    XY(37).X = mlWidth - 8
    XY(37).Y = mlHeight - 2
    XY(38).X = mlWidth - 9
    XY(38).Y = mlHeight - 2
    XY(39).X = mlWidth - 10
    XY(39).Y = mlHeight - 1
    XY(40).X = mlWidth - 11
    XY(40).Y = mlHeight - 1
    XY(41).X = mlWidth - 12
    XY(41).Y = mlHeight - 0

    'Bottom Left Corner
    XY(42).X = 12
    XY(42).Y = mlHeight - 0
    XY(43).X = 11
    XY(43).Y = mlHeight - 1
    XY(44).X = 10
    XY(44).Y = mlHeight - 1
    XY(45).X = 9
    XY(45).Y = mlHeight - 2
    XY(46).X = 8
    XY(46).Y = mlHeight - 2
    XY(47).X = 6
    XY(47).Y = mlHeight - 3
    XY(48).X = 5
    XY(48).Y = mlHeight - 4
    XY(49).X = 4
    XY(49).Y = mlHeight - 5
    XY(50).X = 3
    XY(50).Y = mlHeight - 6
    XY(51).X = 2
    XY(51).Y = mlHeight - 8
    XY(52).X = 2
    XY(52).Y = mlHeight - 9
    XY(53).X = 1
    XY(53).Y = mlHeight - 10
    XY(54).X = 1
    XY(54).Y = mlHeight - 11
    XY(55).X = 0
    XY(55).Y = mlHeight - 12

    'Pass in the address of the first point and
    'the number of points.

    hRgn = CreatePolygonRgn(XY(0), (UBound(XY) + 1), 2)
    lRes = SetWindowRgn(.hWnd, hRgn, True)
End With


'Resize the border to fit:
'shpBorder.Height = Me.ScaleHeight
'shpBorder.Width = Me.ScaleWidth

'This does make the border two (as opposed to one, as on the other sides)
'pixels thick on the right and bottom sides, but that can sort of look like
'a shadow and not ugly ... right? If we add +1 to the end of both statements
'above, it's only one pixel thick and looks good, except it won't completely
'cover the corners -- and we don't want that! In the future, I plan to pick
'at my form-shaping code to make it match the shape control better

End Sub

Private Sub Timer1_Timer()
Hide
Timer1.Enabled = False
End Sub
