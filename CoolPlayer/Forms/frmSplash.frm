VERSION 5.00
Begin VB.Form frmSp 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   3135
   ClientLeft      =   4905
   ClientTop       =   3840
   ClientWidth     =   4575
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
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSt 
      Interval        =   1000
      Left            =   0
      Top             =   3120
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Date"
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label lblSp 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "CoolPlayer by John..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Made by John"
      Top             =   2280
      Width           =   4335
   End
End
Attribute VB_Name = "frmSp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cnt As Integer
Private Sub Form_Click()
    frmSp.Hide
End Sub
Private Sub Form_Initialize()
    Call CreateKey
End Sub
Private Sub Form_Load()

    On Error GoTo LError
    Call CheckSplash
    Call Graph.LoadElliptic(frmSp)
    Call Graph.Ontop(frmSp.hwnd, True)
    Call AddCommands(Command$)

LError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub lblDate_Click()
    frmSp.Hide
End Sub
Private Sub lblSp_Click()
    frmSp.Hide
End Sub
Private Sub tmrSt_Timer()

    On Error GoTo TError
    Cnt = Cnt + 1
    If Cnt = 3 Then Call LoadMain: Unload frmSp

TError:
    If Err.Number <> 0 Then Call LoadMain: Unload frmSp: Exit Sub

End Sub
