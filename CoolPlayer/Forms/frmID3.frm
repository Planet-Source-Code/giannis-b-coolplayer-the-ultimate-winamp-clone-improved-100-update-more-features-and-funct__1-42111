VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmID3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MPEG File Info & ID3 Tag Editor."
   ClientHeight    =   4935
   ClientLeft      =   4950
   ClientTop       =   1965
   ClientWidth     =   5175
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmID3.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleMode       =   0  'User
   ScaleWidth      =   5148.798
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   4680
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3678
            MinWidth        =   3678
            Text            =   "Info:"
            TextSave        =   "Info:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5450
            MinWidth        =   5450
            Text            =   "Filename:"
            TextSave        =   "Filename:"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRem 
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      ToolTipText     =   "Remove ID3 Tag"
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      ToolTipText     =   "Close the editor"
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      ToolTipText     =   "Open file"
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Save ID3 Tag"
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame fraID3 
      Caption         =   "ID3 v1."
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "ID3 v1 Tag information"
      Top             =   0
      Width           =   4935
      Begin VB.ComboBox lstGen 
         Height          =   315
         ItemData        =   "frmID3.frx":030A
         Left            =   2280
         List            =   "frmID3.frx":030C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   0
         ToolTipText     =   "Title"
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox txtArt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   1
         ToolTipText     =   "Artist"
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtAlb 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Album"
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txtYear 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   3
         ToolTipText     =   "Year"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtCom 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   4
         ToolTipText     =   "Comments"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label lblTit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Title:"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblArt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Artist:"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblAlb 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Album:"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Year:"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblCom 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Comments:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblGenre 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Genre:"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   1350
         Width           =   495
      End
   End
   Begin VB.Frame fraMPEG 
      Caption         =   "MPEG Information"
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "MPEG Ingormation"
      Top             =   2640
      Width           =   4935
      Begin VB.Label lblInfoS 
         Height          =   1335
         Left            =   2535
         TabIndex        =   18
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         Height          =   1335
         Left            =   130
         TabIndex        =   17
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Menu mnuT 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuTCl 
         Caption         =   "&Close Editor..."
      End
      Begin VB.Menu mnuTS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTex 
         Caption         =   "&Exit CoolPlayer"
      End
   End
End
Attribute VB_Name = "frmID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    Call CloseEditor
End Sub
Private Sub cmdOpen_Click()

    On Error GoTo OError
    With ODialog
     .CancelError = True
     .hOwner = frmID3.hwnd
     .DialogTitle = "Open a MP3 file for ID3 Tag reading..."
     .Filter = "MP3 Files (*.mp3)|*.mp3|All Files (*.*)|*.*"
     .FileName = ""
     Call .ShowOpen(CI.bTop)
    End With
    Call GetIt(ODialog.FileName)

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub cmdRem_Click()

    With frmID3
     .Status.Panels(1).Text = IDx.RemoveTag(IDx.FileName)
     .txtTitle.Text = IDx.Title
     .txtArt.Text = IDx.Artist
     .txtAlb.Text = IDx.Album
     .txtYear.Text = IDx.Year
     .txtCom.Text = IDx.Comments
     .lstGen.ListIndex = GetGenre(CInt(IDx.Genre))
    End With

End Sub
Private Sub cmdSave_Click()

    With frmID3
     IDx.Title = .txtTitle.Text
     IDx.Album = .txtAlb.Text
     IDx.Artist = .txtArt.Text
     IDx.Year = .txtYear.Text
     IDx.Comments = .txtCom.Text
     IDx.Genre = .lstGen.ItemData(.lstGen.ListIndex)
     .Status.Panels(1).Text = IDx.writetag(IDx.FileName)
    End With

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseEditor
End Sub
Private Sub mnutCl_Click()
    Call CloseEditor
End Sub
Private Sub mnutEx_Click()
    Call ProgramExit
End Sub
