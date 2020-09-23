VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About "
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "About webdog"
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Karl J. Sak Copyright 2000"
      Top             =   840
      Width           =   3135
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   360
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   720
      ScaleMode       =   0  'User
      ScaleWidth      =   1350
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   1350
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Email:  ksak@email.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "MP3@Play"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   525
      Left            =   2040
      TabIndex        =   1
      Tag             =   "Product"
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
        Unload Me
End Sub

Private Sub Label1_Click()

End Sub
