VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Priyan's Directory Search Class [Test App]"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chksearchsubfolders 
      Caption         =   "Search in sub folders"
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdpause 
      Caption         =   "Pause"
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Path"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtpattern 
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Text            =   "*.exe|*.com"
      Top             =   240
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "seperate each pattern by |"
      Height          =   195
      Left            =   5520
      TabIndex        =   9
      Top             =   240
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pattern"
      Height          =   195
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   510
   End
   Begin VB.Label lblstatus 
      Caption         =   "Label1"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   4440
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents search As clsdirsearch
Attribute search.VB_VarHelpID = -1
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Private Sub chksearchsubfolders_Click()
search.searchinsubfolders = chksearchsubfolders.Value
End Sub

Private Sub cmdpause_Click()
If cmdpause.Caption = "Pause" Then
    cmdsearch.Enabled = False
    search.pause = True
    cmdpause.Caption = "Resume"
Else
    search.pause = False
    cmdpause.Caption = "Pause"
    cmdsearch.Enabled = True
End If
End Sub

Private Sub cmdsearch_Click()
Dim start_time As Double
If cmdsearch.Caption = "Search" Then
    cmdsearch.Caption = "Cancel"
    Me.ListView1.ListItems.Clear
    chksearchsubfolders.Enabled = False
    cmdpause.Visible = True
    start_time = Timer
    search.findfiles Dir1.path, txtpattern.Text
    lblstatus.Caption = search.filesfound.Count & " Files Found in " & Round(Timer - start_time, 1) & " Sec"
    MsgBox lblstatus.Caption, vbInformation
    cmdsearch.Caption = "Search"
    cmdpause.Visible = False
    chksearchsubfolders.Enabled = True
Else
    search.Cancel
    cmdpause.Visible = False
    cmdsearch.Caption = "Search"
    chksearchsubfolders.Enabled = True
End If
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.path = Drive1.Drive
End Sub

Private Sub Form_Load()
Set search = New clsdirsearch
lblstatus.Caption = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub





Private Sub ListView1_DblClick()
On Error Resume Next
If MsgBox("Open?", vbYesNo + vbQuestion) = vbYes Then
    ShellExecute 0, "open", ListView1.SelectedItem.Text & ListView1.SelectedItem.SubItems(1), "", "", 1
End If
End Sub

Private Sub search_found(ByVal path$, file As String)
Dim litem As ListItem
Set litem = Me.ListView1.ListItems.Add(, , path)
litem.ListSubItems.Add , , file
litem.ListSubItems.Add , , Round(FileLen(path & file) / 1024, 2) & " Kb"
End Sub

Private Sub search_searching(ByVal path As String)
lblstatus.Caption = "Searching in " & path
End Sub
