VERSION 5.00
Object = "{2C1EC115-F1BA-11D3-BF43-00A0CC32BE58}#9.0#0"; "DMC2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9BBAA498-5614-11D2-B037-444553540000}#1.0#0"; "JSCROLLH.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frmmain 
   BackColor       =   &H00000000&
   Caption         =   "MP3@Play"
   ClientHeight    =   5820
   ClientLeft      =   1650
   ClientTop       =   3270
   ClientWidth     =   8700
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      BackColor       =   &H00000000&
      Caption         =   "Repeat"
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
      Left            =   6960
      TabIndex        =   29
      ToolTipText     =   "Repeat Current Song"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Loop"
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
      Left            =   6960
      TabIndex        =   28
      ToolTipText     =   "Continous Play of Song Queue"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Random"
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
      Left            =   6960
      TabIndex        =   27
      ToolTipText     =   "Randomly Play Song Queue"
      Top             =   4320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1080
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "Frmmain.frx":0442
      DataField       =   "listname"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   4680
      TabIndex        =   24
      ToolTipText     =   "Playlists"
      Top             =   5400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   0
      ForeColor       =   16777088
      ListField       =   "listname"
      Text            =   "DBCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "playlist"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ListName"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   6720
      TabIndex        =   21
      ToolTipText     =   "Song Position"
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin Project1.JScrollH JScrollH4 
      Height          =   255
      Left            =   6720
      TabIndex        =   20
      ToolTipText     =   "Adjust Song Position"
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Value           =   0
      Min             =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   6195
      TabIndex        =   12
      Top             =   240
      Width           =   6255
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Song:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   615
      End
      Begin VB.Label songname 
         BackColor       =   &H00404040&
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
         Height          =   495
         Left            =   720
         TabIndex        =   17
         Top             =   120
         Width           =   5415
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "Time:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
      Begin VB.Label songlength 
         BackColor       =   &H00404040&
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
         Left            =   720
         TabIndex        =   15
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   "Info:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
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
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1560
      Top             =   2640
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00404040&
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
      Height          =   2790
      ItemData        =   "Frmmain.frx":0456
      Left            =   240
      List            =   "Frmmain.frx":0458
      TabIndex        =   10
      Top             =   2520
      Width           =   6255
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   6195
      TabIndex        =   7
      Top             =   1560
      Width           =   6255
      Begin VB.Shape rightmeter 
         BackColor       =   &H00404040&
         BorderColor     =   &H00404040&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   840
         Top             =   240
         Width           =   15
      End
      Begin VB.Shape leftmeter 
         BackColor       =   &H00404040&
         BorderColor     =   &H00404040&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   840
         Top             =   0
         Width           =   15
      End
      Begin VB.Label Label11 
         BackColor       =   &H00404040&
         Caption         =   "Right:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404040&
         Caption         =   "Left:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   615
      End
   End
   Begin Project1.JScrollH JScrollH3 
      Height          =   255
      Left            =   6720
      TabIndex        =   6
      ToolTipText     =   "Adjust Frequency of Song"
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Value           =   10000
      Min             =   25
      Max             =   25000
      Step            =   250
   End
   Begin Project1.JScrollH JScrollH2 
      Height          =   255
      Left            =   6720
      TabIndex        =   1
      ToolTipText     =   "Adjust Balance"
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Value           =   0
      Min             =   -100
      Step            =   25
   End
   Begin Project1.JScrollH JScrollH1 
      Height          =   255
      Left            =   6720
      TabIndex        =   0
      ToolTipText     =   "Adjust Volume"
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Value           =   100
      Min             =   0
      Step            =   5
   End
   Begin DMC2.DMC DMC1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About"
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
      Left            =   7680
      TabIndex        =   32
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clear Queue"
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
      Left            =   1800
      TabIndex        =   31
      ToolTipText     =   "Clear Song Queue"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remove Song"
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
      Left            =   360
      TabIndex        =   30
      ToolTipText     =   "Remove Song From Song Queue"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   6720
      Picture         =   "Frmmain.frx":045A
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7920
      Picture         =   "Frmmain.frx":0D24
      ToolTipText     =   "Paused"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7320
      Picture         =   "Frmmain.frx":15EE
      ToolTipText     =   "Playing"
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   7320
      Picture         =   "Frmmain.frx":1EB8
      ToolTipText     =   "Stopped"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Caption         =   "Position Control:"
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
      Left            =   6720
      TabIndex        =   26
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Color"
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
      Left            =   6720
      TabIndex        =   25
      ToolTipText     =   "Change Display Color"
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select List"
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
      Left            =   3240
      TabIndex        =   23
      ToolTipText     =   "Select Playlist"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Play List"
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
      Left            =   6720
      TabIndex        =   22
      ToolTipText     =   "Create/Edit Play Lists"
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
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
      Left            =   8040
      TabIndex        =   19
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "Song Queue:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Image pause 
      Enabled         =   0   'False
      Height          =   480
      Left            =   7920
      Picture         =   "Frmmain.frx":2782
      ToolTipText     =   "Pause"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image next 
      Height          =   480
      Left            =   7920
      Picture         =   "Frmmain.frx":304C
      ToolTipText     =   "Next Track"
      Top             =   720
      Width           =   480
   End
   Begin VB.Image play 
      Height          =   480
      Left            =   7320
      Picture         =   "Frmmain.frx":3916
      ToolTipText     =   "Play"
      Top             =   720
      Width           =   480
   End
   Begin VB.Image last 
      Height          =   480
      Left            =   6720
      Picture         =   "Frmmain.frx":41E0
      ToolTipText     =   "Previous track"
      Top             =   720
      Width           =   480
   End
   Begin VB.Image stop 
      Height          =   480
      Left            =   7320
      Picture         =   "Frmmain.frx":4AAA
      ToolTipText     =   "Stop"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image open 
      Height          =   480
      Left            =   6720
      Picture         =   "Frmmain.frx":5374
      ToolTipText     =   "Open"
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Frequency:"
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
      Left            =   6720
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
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
      Left            =   7920
      TabIndex        =   4
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Balance:"
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
      Left            =   6720
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Volume:"
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
      Left            =   6720
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = vbChecked Then
    Check2.Value = vbUnchecked
    Check3.Value = vbUnchecked
End If
End Sub
Private Sub Check2_Click()
If Check2.Value = vbChecked Then
    Check3.Value = vbUnchecked
    Check1.Value = vbUnchecked
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = vbChecked Then
    Check2.Value = vbUnchecked
    Check1.Value = vbUnchecked
End If
End Sub

Private Sub Form_Load()
'initialize digital music control

DMC1.DeviceToUse = 1 ' defautl sound card
DMC1.InitBASS Frmmain.hWnd, 44100, False, False
DMC1.BufferLenInSeconds = 1#
Label4 = "100%"
Label5 = "Center"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'must do this routine when leaving app
DMC1.StopStream
DMC1.CloseStream
DMC1.TerminateBASS
Unload Me
End
End Sub

Private Sub Image1_Click()
DMC1.ResumeStream
Timer1.Enabled = True
Timer2.Enabled = True
Image1.Visible = False
Image2.Visible = True
Image3.Visible = False

End Sub

Private Sub JScrollH1_Change()
'controls volume
DMC1.StreamVol = JScrollH1.Value
Label4 = JScrollH1.Value & "%"

End Sub

Private Sub JScrollH2_Change()
' contols speaker balance
DMC1.StreamPan = JScrollH2.Value
If JScrollH2.Value <= -25 Then
    Label5 = "Left"
End If
If JScrollH2.Value >= 25 Then
    Label5 = "Right"
End If
If JScrollH2.Value > -25 And JScrollH2.Value < 25 Then
    Label5 = "Center"
End If

End Sub

Private Sub JScrollH3_Change()
' sets the frequency at which the song is played and resets the time based on frequency
Dim time As Double
Dim Minutes As Integer, Seconds As Integer
Dim strMinutes As String, strSeconds As String
DMC1.StreamFreq = JScrollH3.Value * 4&
time = DMC1.StreamLenInSeconds
Minutes = time / 60
Seconds = (time Mod 60)
If Seconds > 30 Then
    Minutes = Minutes - 1
End If
If Minutes < 10 Then strMinutes = "0" & Minutes Else strMinutes = Minutes
If Seconds < 10 Then strSeconds = "0" & Seconds Else strSeconds = Seconds
totaltime = strMinutes & ":" & strSeconds    'save it here
End Sub

Private Sub JScrollH4_Change()
'contols position of song
Dim pos As Long

pos = DMC1.StreamLen / 100 * JScrollH4.Value

If pos > 0 Then
    DMC1.StreamPos = pos
Else
    DMC1.StreamPos = 0
End If

End Sub

Private Sub Label14_Click()
'adds songs to song queue based on the play list that is selected
On Error Resume Next
Data1.Recordset.MoveFirst
Data2.RecordSource = "Select * from playlist where listname='" & DBCombo1.Text & "'"
Data2.Refresh
Data2.Recordset.MoveLast
Data2.Recordset.MoveFirst
For a = 1 To Data2.Recordset.RecordCount
    List1.AddItem Data2.Recordset.Fields("song").Value
    Data2.Recordset.MoveNext
Next a
End Sub

Private Sub Label15_Click()
'sets color scheme
cd1.ShowColor
If cd1.CancelError = False Then
    Label3.ForeColor = cd1.Color
    songname.ForeColor = cd1.Color
    Label6.ForeColor = cd1.Color
    songlength.ForeColor = cd1.Color
    Label8.ForeColor = cd1.Color
    Label9.ForeColor = cd1.Color
    Label10.ForeColor = cd1.Color
    Label11.ForeColor = cd1.Color
    Label13.ForeColor = cd1.Color
    Label1.ForeColor = cd1.Color
    Label4.ForeColor = cd1.Color
    Label2.ForeColor = cd1.Color
    Label5.ForeColor = cd1.Color
    Label12.ForeColor = cd1.Color
    Label7.ForeColor = cd1.Color
    Check1.ForeColor = cd1.Color
    Check2.ForeColor = cd1.Color
    Check3.ForeColor = cd1.Color
    Label14.ForeColor = cd1.Color
    DBCombo1.ForeColor = cd1.Color
    List1.ForeColor = cd1.Color
    Label15.ForeColor = cd1.Color
    Label17.ForeColor = cd1.Color
    Label16.ForeColor = cd1.Color
    Label18.ForeColor = cd1.Color
End If
End Sub



Private Sub Label16_Click()
'remove song from playlist
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Label18_Click()
'clear song queue
List1.Clear
End Sub

Private Sub Label19_Click()
frmAbout.Show modal, Me
End Sub

Private Sub Label7_Click()
frmplaylist.Show modal, Me

End Sub

Private Sub last_Click()
' replays the last song in queue
If List1.ListIndex = 0 Then
    Exit Sub
End If
List1.ListIndex = List1.ListIndex - 1
song = List1.List(List1.ListIndex)
play_Click
End Sub



Private Sub List1_DblClick()
' plays a song that is selected in the song queue
song = List1.List(List1.ListIndex)
play_Click

End Sub

Private Sub next_Click()
'plays the next available song in the song queue
If List1.ListIndex = List1.ListCount - 1 Then
    Exit Sub
End If
List1.ListIndex = List1.ListIndex + 1
song = List1.List(List1.ListIndex)
play_Click
End Sub

Private Sub open_Click()
Image6.Visible = True
frmopen.Show modal, Me

End Sub

Private Sub pause_Click()
DMC1.PauseStream
Timer1.Enabled = False
Timer2.Enabled = False
Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
End Sub

Private Sub play_Click()
Dim time As Double
Dim Minutes As Integer, Seconds As Integer
Dim strMinutes As String, strSeconds As String
pause.Enabled = True
songname = song
DMC1.OpenStream song
DMC1.PlayStream False
Timer1.Enabled = True
Timer2.Enabled = True
time = DMC1.StreamLenInSeconds
Minutes = time / 60
Seconds = (time Mod 60)
If Seconds > 30 Then
    Minutes = Minutes - 1
End If
If Minutes < 10 Then strMinutes = "0" & Minutes Else strMinutes = Minutes
If Seconds < 10 Then strSeconds = "0" & Seconds Else strSeconds = Seconds
totaltime = strMinutes & ":" & strSeconds    'save it here
Image1.Visible = False
Image2.Visible = True
Image3.Visible = False
If DMC1.StreamIsMono = True Then
    Label9.Caption = "Mono"
Else
    Label9.Caption = "Stereo"
End If
If DMC1.StreamIs8bit = True Then
    Label9.Caption = Label9.Caption & " - 8 bit - " & DMC1.StreamFreq & " Hz"
Else
    Label9.Caption = Label9.Caption & " - 16 bit - " & DMC1.StreamFreq & " Hz"
End If
End Sub


Private Sub stop_Click()
DMC1.StopStream
Image1.Visible = False
Image2.Visible = False
Image3.Visible = True
Timer1.Enabled = False
Timer2.Enabled = False

End Sub

Private Sub Timer1_Timer()
If DMC1.StreamIsActive = True Then
    leftmeter.Width = DMC1.StreamLeftLevel * 10
    rightmeter.Width = DMC1.StreamRightLevel * 10
End If
        
End Sub

Private Sub Timer2_Timer()
Dim time As Double
Dim Minutes As Integer, Seconds As Integer
Dim strMinutes As String, strSeconds As String

If DMC1.StreamIsActive Then
    time = DMC1.StreamPosInSeconds
    Minutes = time / 60
    Seconds = time Mod 60
    If Seconds > 30 Then
        Minutes = Minutes - 1
    End If
    
    'much faster way
    If Minutes < 10 Then strMinutes = "0" & Minutes Else strMinutes = Minutes
    If Seconds < 10 Then strSeconds = "0" & Seconds Else strSeconds = Seconds
    
    songlength.Caption = strMinutes & ":" & strSeconds & " -- " & totaltime
    ProgressBar1.Value = DMC1.StreamPos / (DMC1.StreamLen / 100)
Else
    DMC1.StopStream
    Image2.Visible = False
    time = 0
    Timer1.Enabled = False
    Timer2.Enabled = False
    songlength = ""
    ProgressBar1.Value = 0
    JScrollH4.Value = 0
    Label9 = ""
    ' if random play is checked
    If Check1.Value = vbChecked Then
        Randomize
        MsgBox List1.ListCount
        selectedsong = Int((List1.ListCount - 0) * Rnd + 0)
        MsgBox selectedsong
        List1.ListIndex = selectedsong
        song = List1.List(selectedsong)
        play_Click
        Exit Sub
    End If
    ' if repeat is checked
    If Check3.Value = vbChecked Then
        play_Click
        Exit Sub
    End If
    ' if loop is checked
    If List1.ListIndex = List1.ListCount - 1 And Check2.Value = vbChecked Then
        List1.ListIndex = 0
        song = List1.List(0)
        play_Click
        Exit Sub
    End If
    ' automatically plays next song in queue
    If List1.ListCount > 0 And List1.ListIndex <> List1.ListCount - 1 Then
        List1.ListIndex = List1.ListIndex + 1
        song = List1.List(List1.ListIndex)
        play_Click
    End If
End If

End Sub
