VERSION 5.00
Begin VB.Form frmAtomic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DigiTray - Atomic Clock"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   Icon            =   "frmAtomic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmAtomic.frx":08CA
      Left            =   120
      List            =   "frmAtomic.frx":08E3
      TabIndex        =   2
      Text            =   "Pacific"
      Top             =   120
      Width           =   5895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmAtomic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub


Private Sub Command1_Click()
frmTray.GetAtomicTime
SaveSetting App.Title, "Clock", "Zone", Combo1.Text
Hide
End Sub

Private Sub Command2_Click()
Hide
End Sub


Private Sub Form_Activate()
Combo1.Text = GetSetting(App.Title, "Clock", "Zone", "Pacific")
End Sub

