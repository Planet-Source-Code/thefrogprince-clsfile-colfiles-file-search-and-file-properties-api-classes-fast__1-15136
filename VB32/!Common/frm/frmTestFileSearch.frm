VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestFileSearch 
   Caption         =   "File Classes Test Harness"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "frmTestFileSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin ControlTestHarness.jeffFrame frmSearch 
      Height          =   1185
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   2090
      BorderStyle     =   10
      BorderWidth     =   2
      Begin VB.TextBox txtPath 
         Height          =   345
         Left            =   2100
         TabIndex        =   9
         Text            =   "D:\"
         Top             =   480
         Width           =   3015
      End
      Begin VB.CheckBox chkSubDirs 
         Caption         =   "Search In Subdirectories?"
         Height          =   195
         Left            =   2100
         TabIndex        =   8
         Top             =   870
         Width           =   2265
      End
      Begin VB.TextBox txtFileSpec 
         Height          =   345
         Left            =   2100
         TabIndex        =   6
         Text            =   "*.mp3"
         Top             =   90
         Width           =   2625
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   675
         Left            =   120
         Picture         =   "frmTestFileSearch.frx":2832
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   390
         Width           =   855
      End
      Begin VB.Label lblLookIn 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Look In:"
         Height          =   195
         Left            =   1440
         TabIndex        =   7
         Top             =   510
         Width           =   585
      End
      Begin VB.Label lblSearchFor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Files Named:"
         Height          =   195
         Left            =   1140
         TabIndex        =   5
         Top             =   150
         Width           =   915
      End
   End
   Begin MSComctlLib.ProgressBar pbMain 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   5085
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listFiles 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      Top             =   2430
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4154
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FILENAME"
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "FILEPATH"
         Text            =   "FilePath"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "FILESIZE"
         Text            =   "FileSize"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "DATELASTUPDATED"
         Text            =   "Last Updated"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "DATECREATED"
         Text            =   "Created"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmTestFileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' This code was written by The Frog Prince
'
' If you have questions or comments, I can be reached at
'        TheFrogPrince@hotmail.com
' If you wanna see more cool vb user controls, classes, code,
' and add-ins like this one, or updates to this code, go to
' my web page at
'        http://members.tripod.com/the__frog__prince/
' You are free to use, re-write, or otherwise do as you wish
' with this code.  However, if you do a cool enhancement, I
' would appreciate it if you could e-mail it to me.  I like
' to see what people do with my stuff.  =)
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Option Explicit

Public cFiles As New colFiles



Private Sub cmdSearch_Click()
    Dim dStart As Date
    Dim dFinish As Date
    Dim lScan As Long
    dStart = Now
    cFiles.Clear
    cFiles.LoadFiles ts.sAppend(Me.txtPath, "\") & Me.txtFileSpec, Me.chkSubDirs.Value, Me.pbMain
    lScan = DateDiff("s", dStart, Now)
    
    Me.pbMain.Max = cFiles.Count + 1
    Me.pbMain.Value = 0
    
    dStart = Now
    With Me.listFiles
        .ListItems.Clear
        Dim l As Long
        For l = 1 To cFiles.Count
            Me.sbMain.SimpleText = "Loading grid (" & l & ")..."
            Me.sbMain.Refresh
            Me.pbMain.Value = Me.pbMain.Value + 1
            Me.pbMain.Refresh
            .ListItems.Add , "KEY" & CStr(l), cFiles(l).sNameAndExtension
            .ListItems(.ListItems.Count).ListSubItems.Add , , cFiles(l).sPath
            .ListItems(.ListItems.Count).ListSubItems.Add , , cFiles(l).lSize
            .ListItems(.ListItems.Count).ListSubItems.Add , , cFiles(l).dLastModified
            .ListItems(.ListItems.Count).ListSubItems.Add , , cFiles(l).dCreated
        Next l
    End With
    Me.sbMain.SimpleText = "Found " & cFiles.Count & " files in " & lScan & " second(s).  Loaded grid in " & DateDiff("s", dStart, Now) & " second(s)."
    
End Sub

Private Sub Form_Resize()
    Me.frmSearch.Width = Me.ScaleWidth
    With Me.listFiles
        .Move 0, (Me.frmSearch.Top + Me.frmSearch.Height), Me.ScaleWidth, Me.ScaleHeight - (Me.frmSearch.Top + Me.frmSearch.Height) - Me.pbMain.Height - Me.sbMain.Height
    End With
End Sub

Private Sub frmSearch_Resize()
    With Me.txtFileSpec
        .Move .Left, .Top, frmSearch.Width - cmdSearch.Left - .Left
    End With
    With Me.txtPath
        .Move .Left, .Top, txtFileSpec.Width
    End With
End Sub

Private Sub listFiles_DblClick()
    Dim frm As frmTestFileProperties
    Set frm = New frmTestFileProperties
    Set frm.cFile = cFiles(Val(Replace(Me.listFiles.SelectedItem.Key, "KEY", "")))
    frm.Show vbModal, Me
    Set frm = Nothing
    
End Sub
