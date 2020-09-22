Attribute VB_Name = "modBalloon"
Option Explicit

'Public Const SW_SHOWNOACTIVATE = 4
'Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
'ByVal nCmdShow As Long) As Long 'WILL BE Used to show balloon without "stealing"
                                'focus from window it's called from ... in a
                                'future release of this project

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
        lpRect As RECT) As Long 'Used for getting positions of objects/forms
                                'to place balloons correctly

Public Type RECT   'Also used to store values for positions of balloons
   Left As Long    'after using the API to determine where
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
        lpPoint As POINTAPI) As Long 'Also used for getting positions of
                                     'objects/forms we want to place the
                                     'balloons by
                                     
'Used to draw the ellipse on the form
Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, _
    ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, _
    ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'Used to create the regiod around the form to shape it
Public Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
    ByVal RectY2 As Long, ByVal EllipseWidth As Long, _
    ByVal EllipseHeight As Long) As Long
                                     
Public mlWidth As Long
Public mlHeight As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type BalloonCoords 'Used to store X and Y coordinates of balloon
    X As Long 'after using API and math operations to figure exact
    Y As Long 'coordinates regarding where to place itself
End Type
Public Function DoesTahomaExist()
'This function is is an easy way to determine whether or not a
'font exists by creating a standard font object, assigning a font
'name to it, and checking to see if it does keep that font name (which
'means it does exist, otherwise it'll use a different font or a close
'match).

'This function is hard-coded to check for Tahoma, but if you want to use
'it in another project for something else, you should easily be able to
'modify it.

Dim TestFont As New StdFont 'Create a new standard font object, and ...
TestFont.Name = "Tahoma" 'Assign the in-question font name (Tahoma) to it

'Check to see if the font object's name matches that which we are
'questioning exists (Tahoma); if it does match, it exists, and if not,
'it doesn't. Then return the correct value from this function.
If TestFont.Name = "Tahoma" Then
    DoesTahomaExist = True
Else
    DoesTahomaExist = False
End If
End Function
Public Sub PopBalloon(MYfrm As Form, ObjectHwnd As Long, Text, Title, Optional Icon, Optional ShowClose As Boolean = True, Optional AutoCloseMs)
Dim WinRect As RECT
Dim WinPoint As POINTAPI
Dim BalloonXY As BalloonCoords
On Error Resume Next

Call GetWindowRect(ObjectHwnd, WinRect)
BalloonXY.X = WinRect.Left * Screen.TwipsPerPixelX
BalloonXY.Y = WinRect.Bottom * Screen.TwipsPerPixelY

Dim frmPopUpBalloon As New frmBalloon
frmPopUpBalloon.SetBalloon Title, Text, BalloonXY.X, BalloonXY.Y, Icon, ShowClose, _
    AutoCloseMs
frmPopUpBalloon.Show , MYfrm
frmPopUpBalloon.SetFocus
End Sub


