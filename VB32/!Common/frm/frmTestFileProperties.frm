VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTestFileProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Properties Class Test"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   Icon            =   "frmTestFileProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPathRoot 
      Height          =   315
      Left            =   1170
      TabIndex        =   41
      Text            =   "txtPathRoot"
      Top             =   690
      Width           =   5205
   End
   Begin VB.TextBox txtShortFullFilename 
      Height          =   315
      Left            =   1170
      TabIndex        =   37
      Text            =   "txtShortFullFilename"
      Top             =   360
      Width           =   5205
   End
   Begin VB.TextBox txtFileSize 
      Height          =   315
      Left            =   2250
      TabIndex        =   35
      Text            =   "txtFileSize"
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Frame frmAttributes 
      Caption         =   " Attributes "
      Height          =   1215
      Left            =   150
      TabIndex        =   24
      Top             =   1740
      Width           =   2895
      Begin VB.CheckBox chkTemporary 
         Caption         =   "Temporary?"
         Height          =   195
         Left            =   1560
         TabIndex        =   32
         Top             =   900
         Width           =   1275
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "System?"
         Height          =   195
         Left            =   1560
         TabIndex        =   31
         Top             =   690
         Width           =   1275
      End
      Begin VB.CheckBox chkReadonly 
         Caption         =   "Read Only?"
         Height          =   195
         Left            =   1560
         TabIndex        =   30
         Top             =   480
         Width           =   1275
      End
      Begin VB.CheckBox chkNormal 
         Caption         =   "Normal?"
         Height          =   195
         Left            =   1560
         TabIndex        =   29
         Top             =   270
         Width           =   1275
      End
      Begin VB.CheckBox chkHidden 
         Caption         =   "Hidden?"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   900
         Width           =   1425
      End
      Begin VB.CheckBox chkDirectory 
         Caption         =   "Directory?"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   690
         Width           =   1425
      End
      Begin VB.CheckBox chkCompressed 
         Caption         =   "Compressed?"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   480
         Width           =   1425
      End
      Begin VB.CheckBox chkArchive 
         Caption         =   "Archive?"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   270
         Width           =   1425
      End
   End
   Begin VB.Frame frmFileDates 
      Caption         =   " Dates "
      Height          =   1215
      Left            =   3150
      TabIndex        =   17
      Top             =   1740
      Width           =   3225
      Begin MSMask.MaskEdBox meCreated 
         Height          =   285
         Left            =   1260
         TabIndex        =   18
         Top             =   210
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "mm/dd/yyyy hh:mm AM/PM"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meAccessed 
         Height          =   285
         Left            =   1260
         TabIndex        =   19
         Top             =   510
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "mm/dd/yyyy hh:mm AM/PM"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meUpdated 
         Height          =   285
         Left            =   1260
         TabIndex        =   20
         Top             =   810
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         Format          =   "mm/dd/yyyy hh:mm AM/PM"
         PromptChar      =   "_"
      End
      Begin VB.Label lblModified 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Modified:"
         Height          =   195
         Left            =   570
         TabIndex        =   23
         Top             =   840
         Width           =   645
      End
      Begin VB.Label lblAccessed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Last Accessed:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label lblCreated 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Created:"
         Height          =   195
         Left            =   615
         TabIndex        =   21
         Top             =   270
         Width           =   600
      End
   End
   Begin VB.Frame frmVolumeInfo 
      Caption         =   " Volume Info "
      Height          =   2025
      Left            =   150
      TabIndex        =   3
      Top             =   3060
      Width           =   3225
      Begin VB.TextBox txtVolLabel 
         Height          =   285
         Left            =   930
         TabIndex        =   39
         Text            =   "txtVolLabel"
         Top             =   180
         Width           =   2175
      End
      Begin VB.TextBox txtVolSerialNumber 
         Height          =   285
         Left            =   930
         TabIndex        =   33
         Text            =   "txtVolSerialNumber"
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox chkFixedDisk 
         Caption         =   "Fixed Disk?"
         Height          =   195
         Left            =   1410
         TabIndex        =   13
         Top             =   870
         Width           =   1785
      End
      Begin VB.CheckBox chkNetworkPath 
         Caption         =   "Network Path?"
         Height          =   195
         Left            =   1410
         TabIndex        =   12
         Top             =   1080
         Width           =   1785
      End
      Begin VB.CheckBox chkUNC 
         Caption         =   "UNC?"
         Height          =   195
         Left            =   1410
         TabIndex        =   11
         Top             =   1290
         Width           =   1785
      End
      Begin VB.CheckBox chkUNCServer 
         Caption         =   "UNC Server?"
         Height          =   195
         Left            =   1410
         TabIndex        =   10
         Top             =   1500
         Width           =   1785
      End
      Begin VB.CheckBox chkUNCServerShare 
         Caption         =   "UNC Server Share?"
         Height          =   195
         Left            =   1410
         TabIndex        =   9
         Top             =   1710
         Width           =   1785
      End
      Begin VB.CheckBox chkERemote 
         Caption         =   "Remote?"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   1500
         Width           =   1365
      End
      Begin VB.CheckBox chkERamDisk 
         Caption         =   "Ram Disk?"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   1290
         Width           =   1365
      End
      Begin VB.CheckBox chkERemovable 
         Caption         =   "Removable?"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   1710
         Width           =   1365
      End
      Begin VB.CheckBox chkEFixedDisk 
         Caption         =   "Fixed Disk?"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   1080
         Width           =   1365
      End
      Begin VB.CheckBox chkECDRom 
         Caption         =   "CD Rom?"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   870
         Width           =   1365
      End
      Begin VB.Label lblVolLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Label:"
         Height          =   195
         Left            =   465
         TabIndex        =   40
         Top             =   210
         Width           =   435
      End
      Begin VB.Label lblVolSerialNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Serial No:"
         Height          =   195
         Left            =   210
         TabIndex        =   34
         Top             =   510
         Width           =   690
      End
   End
   Begin VB.TextBox txtFileExtension 
      Height          =   315
      Left            =   1170
      TabIndex        =   2
      Text            =   "txtFileExtension"
      Top             =   1350
      Width           =   585
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Text            =   "txtFileName"
      Top             =   1020
      Width           =   5205
   End
   Begin VB.TextBox txtFullFileName 
      Height          =   315
      Left            =   1170
      TabIndex        =   0
      Text            =   "txtFullFileName"
      Top             =   30
      Width           =   5205
   End
   Begin VB.Label lblComments 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmTestFileProperties.frx":27A2
      Height          =   1995
      Left            =   3540
      TabIndex        =   43
      Top             =   3090
      Width           =   2775
   End
   Begin VB.Label lblPathRoot 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Path Root:"
      Height          =   195
      Left            =   360
      TabIndex        =   42
      Top             =   750
      Width           =   765
   End
   Begin VB.Label lblShortFullFilename 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Short Filename:"
      Height          =   195
      Left            =   30
      TabIndex        =   38
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label lblFileSize 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Size:"
      Height          =   195
      Left            =   1830
      TabIndex        =   36
      Top             =   1410
      Width           =   345
   End
   Begin VB.Label lblExtension 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Ext:"
      Height          =   195
      Left            =   855
      TabIndex        =   16
      Top             =   1440
      Width           =   270
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   660
      TabIndex        =   15
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label lblFullFilename 
      Alignment       =   1  'Right Justify
      Caption         =   "&Full Filename:"
      Height          =   225
      Left            =   150
      TabIndex        =   14
      Top             =   90
      Width           =   975
   End
End
Attribute VB_Name = "frmTestFileProperties"
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

Public cFile As clsFile


Private Sub Form_Activate()
    Me.txtFileExtension = cFile.sExtension
    Me.txtFileName = cFile.sName
    Me.txtFullFileName = cFile.sFilename
    Me.meAccessed = cFile.dLastAccessed
    Me.meCreated = cFile.dCreated
    Me.meUpdated = cFile.dLastModified
    Me.txtFileSize = cFile.lSize
    Me.txtShortFullFilename = cFile.sShortName
    Me.txtPathRoot = cFile.sPathRoot
    
    Me.chkArchive = Abs(CBool(cFile.eAttributes And efaARCHIVE))
    Me.chkCompressed = Abs(CBool(cFile.eAttributes And efaCOMPRESSED))
    Me.chkDirectory = Abs(CBool(cFile.eAttributes And efaDIRECTORY))
    Me.chkHidden = Abs(CBool(cFile.eAttributes And efaHIDDEN))
    Me.chkNormal = Abs(CBool(cFile.eAttributes And efaNORMAL))
    Me.chkReadonly = Abs(CBool(cFile.eAttributes And efaREADONLY))
    Me.chkSystem = Abs(CBool(cFile.eAttributes And efaSYSTEM))
    Me.chkTemporary = Abs(CBool(cFile.eAttributes And efaTEMPORARY))
    
    Me.txtVolLabel = cFile.sVolumeName
    Me.txtVolSerialNumber = cFile.lVolumeSerialNo
    Me.chkECDRom = Abs(cFile.eVolumeType = DRIVE_CDROM)
    Me.chkEFixedDisk = Abs(cFile.eVolumeType = DRIVE_FIXED)
    Me.chkERamDisk = Abs(cFile.eVolumeType = DRIVE_RAMDISK)
    Me.chkERemote = Abs(cFile.eVolumeType = DRIVE_REMOTE)
    Me.chkERemovable = Abs(cFile.eVolumeType = DRIVE_REMOVABLE)
    
    Me.chkFixedDisk = Abs(cFile.bFixedDisk)
    Me.chkNetworkPath = Abs(cFile.bNetworkPath)
    Me.chkUNC = Abs(cFile.bUNC)
    Me.chkUNCServer = Abs(cFile.bUNCServer)
    Me.chkUNCServerShare = Abs(cFile.bUNCServerShare)
    
    
End Sub

