VERSION 5.00
Begin VB.Form frmAb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About CoolPlayer"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4935
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMain 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4155
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   480
      Width           =   4695
      Begin VB.PictureBox picPar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   240
         ScaleHeight     =   3825
         ScaleWidth      =   4065
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label lblFull 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "This proggy is made by John. Now, finally added a balance bar and some fixes... Enjoy it..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   3375
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Label lblMAb 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "    CoolPlayer     By John"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2535
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
      End
   End
   Begin VB.Timer tmrSc 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4440
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Close"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Close the dialog"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label lblAb 
      Caption         =   "About"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "About CoolPlayer"
      Top             =   120
      Width           =   615
   End
   Begin VB.Line linA 
      X1              =   2400
      X2              =   2400
      Y1              =   360
      Y2              =   120
   End
   Begin VB.Label lblMail 
      Caption         =   "Mail me at: Giannis@usa.com"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2640
      MouseIcon       =   "frmAbout.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Mail me"
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label lblCP 
      Caption         =   "CoolPlayer"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "About CoolPlayer"
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuT 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuTOk 
         Caption         =   "&OK John"
      End
      Begin VB.Menu mnuTS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTex 
         Caption         =   "&Exit CoolPlayer"
      End
   End
End
Attribute VB_Name = "frmAb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mnutEx_Click()
    Call ProgramExit
End Sub
Private Sub mnutOK_Click()

    Call DisableForms(True)
    Unload frmAb

End Sub
Private Sub cmdOk_Click()
    
    Call DisableForms(True)
    Unload frmAb

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call DisableForms(True)
End Sub
Private Sub lblAb_Click()

    On Error Resume Next
    If GL.bClick = False Then Exit Sub
    If tmrSc.Enabled = True Then Exit Sub
    Call SHLabels(False)
    GL.bClick = False

End Sub
Private Sub lblAb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     lblAb.ForeColor = vbBlue
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub lblAb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAb.ForeColor = IIf(Allow(lblAb, Button, X, Y), vbBlue, vbBlack)
End Sub
Private Sub lblAb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then lblAb.ForeColor = vbBlack
End Sub
Private Sub lblCP_Click()

    On Error Resume Next
    If GL.bClick = False Then Exit Sub
    Call SHLabels(True)
    GL.bClick = False

End Sub
Private Sub lblCP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     lblCP.ForeColor = vbBlue
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub lblCP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCP.ForeColor = IIf(Allow(lblCP, Button, X, Y), vbBlue, vbBlack)
End Sub
Private Sub lblCP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then lblCP.ForeColor = vbBlack
End Sub
Private Sub lblMail_Click()

    Call ShellExecute(hwnd, "open", "mailto:Giannis@usa.com", vbNullString, vbNullString, 5)
    Call HideForms(False)
    Call DisableForms(True)
    Unload frmAb

End Sub
Private Sub tmrSc_Timer()

    On Error Resume Next
    With picPar
     If .Top < picMain.Height - picMain.Height - .Height Then
      .Top = .Height - 1
      .Top = picMain.Height - 5
     Else
      .Top = .Top - 5
     End If
    End With

End Sub
