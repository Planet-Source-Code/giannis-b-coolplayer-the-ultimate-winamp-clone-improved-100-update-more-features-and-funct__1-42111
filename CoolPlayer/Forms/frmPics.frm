VERSION 5.00
Begin VB.Form frmSkn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skin browser"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPics.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox Files 
      Height          =   480
      Left            =   2280
      Pattern         =   "*.txt"
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.DirListBox Dirs 
      Height          =   540
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close browser"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Frame fraSkins 
      Caption         =   "Available skins..."
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton cmdRef 
         Caption         =   "&Refresh"
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox txtInfo 
         BackColor       =   &H8000000F&
         Height          =   975
         HideSelection   =   0   'False
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2520
         Width           =   4815
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "Skins &directory..."
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   3600
         Width           =   1455
      End
      Begin VB.ListBox lstSkins 
         Height          =   2010
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblInfo 
         Caption         =   "CoolPlayer uses Winamp's skins, so the credit goes to Winamp..."
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   2280
         Width           =   4815
      End
   End
   Begin VB.Label lblMore 
      Caption         =   "Get more Winamp skins..."
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmPics.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Menu mnuT 
      Caption         =   "&Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuTCl 
         Caption         =   "&Close skin browser..."
      End
      Begin VB.Menu mnuTS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTex 
         Caption         =   "&Exit CoolPlayer"
      End
   End
End
Attribute VB_Name = "frmSkn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()

    Call DisableForms(True)
    Call SaveIniSettings(True)
    Unload frmSkn

End Sub
Private Sub cmdDir_Click()

    On Error GoTo AError
    Dim fFolder As String

    fFolder = File.BrowseDir(frmSkn.hwnd, "Select a path...", CI.bTop)
    If Len(fFolder) = 0 Then Exit Sub
    Call CheckPath(fFolder, True)
    Call SaveIniSettings(True)

AError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub cmdRef_Click()

    On Error GoTo LError
    Call LoadSkinNumber

LError:
    If Err.Number <> 0 Then Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DisableForms(True)
    Call SaveIniSettings(True)
End Sub

Private Sub lblMore_Click()
    Call Shell("Start.exe " & "http://www.winamp.com", 0)
End Sub
Private Sub lstSkins_Click()

    On Error GoTo CError
    If GL.bClick = False Then Exit Sub
    Call SkinIt
    Call SaveIniSettings(True)
    GL.bClick = False

CError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub lstSkins_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error GoTo PError
    Call SkinIt
    Call SaveIniSettings(True)

PError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub lstSkins_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GL.bClick = IIf(Button = 1, True, False)
End Sub
Private Sub mnutCl_Click()

    Call DisableForms(True)
    frmSkn.Hide

End Sub
Private Sub mnutEx_Click()
    Call ProgramExit
End Sub
