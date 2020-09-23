VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmplaylist 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play Lists"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8685
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   5055
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ListName"
      Top             =   240
      Width           =   1215
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "frmplaylist.frx":0000
      DataField       =   "listname"
      DataSource      =   "Data2"
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Listname"
      Text            =   ""
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   4215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplaylist.frx":0014
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplaylist.frx":0558
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplaylist.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplaylist.frx":0BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplaylist.frx":10F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplaylist.frx":11EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmplaylist.frx":12EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   556
      ButtonWidth     =   609
      ButtonHeight    =   556
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New play list"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "edit"
            Object.ToolTipText     =   "Modify play list"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save current play list"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete selected playlist"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Add a song to the play list"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "remove"
            Object.ToolTipText     =   "Remove a song from the play list"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancel"
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "playlist"
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   5280
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   2640
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   5280
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   5280
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Song Location:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Play List:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmplaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
List1.AddItem File1.Path & "\" & File1.FileName

End Sub

Private Sub Form_Load()

End Sub

Private Sub List1_Click()
Text2.Text = List2.List(List1.ListIndex)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Key
    Case "new"
        DBCombo1.Enabled = False
        Text1.Text = ""
        List1.Clear
        Text1.SetFocus
    Case "edit"
        List1.Clear
        Text1.Text = ""
        Data1.RecordSource = "select * from playlist where listname = '" & DBCombo1.Text & "'"
        Data1.Refresh
        Text1.Text = DBCombo1.Text
        Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            List1.AddItem Data1.Recordset.Fields("song").Value
            Data1.Recordset.MoveNext
        Loop
        Text1.Enabled = False
        editlist = True
    Case "save"
        If editlist = True Then
        
        Else
            If Len(Text1.Text) > 0 Then
                Data2.Recordset.AddNew
                Data2.Recordset.Fields("listname").Value = Text1.Text
                Data2.Recordset.Update
                For a = 0 To List1.ListCount - 1
                    Data1.Recordset.AddNew
                    Data1.Recordset.Fields("listname").Value = Text1.Text
                    Data1.Recordset.Fields("song").Value = List1.List(a)
                    Data1.Recordset.Update
                Next a
            Else
                MsgBox "Please enter a name for the Play list."
            End If
        End If
    Case "add"
        List1.AddItem File1.Path & "\" & File1.FileName
    Case "remove"
        List1.RemoveItem (List1.ListIndex)
    Case "cancel"
    Case "delete"
        Msg = "Are you sure?"
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Delete Selected Playlist"
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            List1.Clear
            Text1.Text = ""
            Text2.Text = ""
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
                Data1.Recordset.Delete
                Data1.Recordset.MoveNext
            Loop
            Data2.Recordset.Delete
            Data2.Refresh
            Data1.Refresh
            Frmmain.Data2.Refresh
        End If
End Select
    
End Sub
