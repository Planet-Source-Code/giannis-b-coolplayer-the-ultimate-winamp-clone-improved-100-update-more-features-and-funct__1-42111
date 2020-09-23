VERSION 5.00
Begin VB.Form frmTim 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jump to time"
   ClientHeight    =   1335
   ClientLeft      =   5670
   ClientTop       =   3120
   ClientWidth     =   3135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "Close the dialog"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdJump 
      Caption         =   "&Jump"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Jump to time"
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame fraMain 
      Caption         =   "Jump to:"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   130
         MaxLength       =   5
         TabIndex        =   0
         Text            =   "00:00"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLen 
         Caption         =   "Track length:"
         Height          =   210
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Trach length"
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblMS 
         Caption         =   "Minutes : Seconds"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Menu mnuT 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuTCl 
         Caption         =   "&Close dialog..."
      End
      Begin VB.Menu mnuTS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTEx 
         Caption         =   "&Exit CoolPlayer"
      End
   End
   Begin VB.Menu mnuTo 
      Caption         =   "To time"
      Visible         =   0   'False
      Begin VB.Menu mnuToJ 
         Caption         =   "&Jump to time..."
      End
      Begin VB.Menu mnuToJS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToF 
         Caption         =   "&Forward 5 seconds"
      End
      Begin VB.Menu mnuToB 
         Caption         =   "&Rewind 5 seconds"
      End
   End
End
Attribute VB_Name = "frmTim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()

    Call DisableForms(True)
    Unload frmTim

End Sub
Private Sub cmdJump_Click()
    Call ToTime(MPlay.duration)
End Sub
Private Sub mnuToB_Click()
    Call GotoTime(-5)
End Sub
Private Sub mnuToF_Click()
    Call GotoTime(5)
End Sub
Private Sub mnuToJ_Click()

    On Error GoTo TError
    With frmTim
     Load frmTim
     Call DisableForms(False)
     Call Graph.Ontop(.hwnd, CI.bTop)
     .txtTime.Text = TimePosition(MPlay.CurrentPosition)
     .lblLen.Caption = "Track length: " & TimePosition(MPlay.duration) & " mins"
     .Show
     .txtTime.SetFocus
    End With

TError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call DisableForms(True)
End Sub
Private Sub txtTime_KeyPress(KeyAscii As Integer)

    On Error GoTo SError
    If KeyAscii < 48 And Not KeyAscii = 8 And Not KeyAscii = 13 And Not KeyAscii = 27 Then
     KeyAscii = 0
    ElseIf KeyAscii > 58 And Not KeyAscii = 8 And Not KeyAscii = 13 And Not KeyAscii = 27 Then
     KeyAscii = 0
    End If

    If KeyAscii = 13 Then
     Call ToTime(MPlay.duration)
    ElseIf KeyAscii = 27 Then
     Call DisableForms(True)
     Unload frmTim
    End If

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub mnutCl_Click()
     
    Call DisableForms(True)
    Unload frmTim

End Sub
Private Sub mnutEx_Click()
    Call ProgramExit
End Sub
