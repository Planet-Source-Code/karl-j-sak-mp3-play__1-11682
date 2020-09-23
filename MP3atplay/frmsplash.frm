VERSION 5.00
Begin VB.Form frmsplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   ScaleHeight     =   960
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   4680
      Top             =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "MP3@Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmsplash.frx":0000
      Top             =   120
      Width           =   1350
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'On Error Resume Next
Frmmain.Data1.DatabaseName = CurDir & "\mp3atplay.mdb"
Frmmain.Data1.RecordSource = "select * from listname"
Frmmain.Data1.Refresh
Frmmain.Data2.DatabaseName = CurDir & "\mp3atplay.mdb"
Frmmain.Data2.RecordSource = "select * from playlist"
Frmmain.Data2.Refresh
frmplaylist.Data1.DatabaseName = CurDir & "\mp3atplay.mdb"
frmplaylist.Data1.RecordSource = "select * from playlist"
frmplaylist.Data1.Refresh
frmplaylist.Data2.DatabaseName = CurDir & "\mp3atplay.mdb"
frmplaylist.Data2.RecordSource = "select * from listname"
frmplaylist.Data2.Refresh

End Sub

Private Sub Timer1_Timer()
Frmmain.Show
Unload Me
End Sub

