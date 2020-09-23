VERSION 5.00
Begin VB.Form frmopen 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   4230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Add Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   0
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   3360
      Width           =   4215
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmopen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
' add all the songs from the currently select directory, doesn't include sub-dirs
For a = 0 To File1.ListCount - 1
    song = File1.Path & "\" & File1.List(a)
    Frmmain.List1.AddItem song
Next a
If Frmmain.List1.ListIndex = -1 Then
    Frmmain.List1.ListIndex = 0
    song = Frmmain.List1.List(0)
End If
lastdir = File1.Path
lastdrive = Drive1.Drive
Frmmain.Image6.Visible = False
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
'automatically plays song in queue is empty
song = File1.Path & "\" & File1.FileName
Frmmain.List1.AddItem song
If Frmmain.List1.ListIndex = -1 Then
    Frmmain.List1.ListIndex = 0
    song = Frmmain.List1.List(0)
End If
lastdir = File1.Path
lastdrive = Drive1.Drive
Frmmain.Image6.Visible = False
Unload Me

End Sub

Private Sub Form_Load()
'keeps track of the last directory the user was viewing
If Len(lastdir) > 0 Then
    Drive1.Drive = lastdrive
    Dir1.Path = lastdir
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Frmmain.Image6.Visible = False
Unload Me

End Sub
