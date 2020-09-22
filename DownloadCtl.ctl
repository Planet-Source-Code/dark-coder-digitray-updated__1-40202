VERSION 5.00
Begin VB.UserControl DownloadCtl 
   BackColor       =   &H00000000&
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   525
   ScaleWidth      =   510
End
Attribute VB_Name = "DownloadCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyleft 2002 by Modulus Financial Engineering
'http://www.modulusfe.com
'under GNU - General Public License:
'http://www.gnu.org/copyleft/gpl.html
'This copyleft notice must remain intact.

Option Explicit

Public Event Progress(BytesRead As Long)
Public Event DownloadComplete(TempFileName As String)

'This is just a simple usercontrol and it provides
'all the functionality we need to perform downloading.

Public Function BeginDownload(URL As String) As String
    UserControl.AsyncRead URL, 1
End Function

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    RaiseEvent DownloadComplete(AsyncProp.Value)
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    RaiseEvent Progress(AsyncProp.BytesRead)
End Sub
