VERSION 5.00
Begin VB.Form frmMn 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "CoolPlayer Main Window"
   ClientHeight    =   1740
   ClientLeft      =   3120
   ClientTop       =   2100
   ClientWidth     =   4125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1740
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMSli 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   240
      ScaleHeight     =   150
      ScaleWidth      =   435
      TabIndex        =   39
      Top             =   1080
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picMBal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   2840
      ScaleHeight     =   165
      ScaleWidth      =   210
      TabIndex        =   38
      Top             =   870
      Width           =   210
   End
   Begin VB.PictureBox picMVol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   2280
      ScaleHeight     =   165
      ScaleWidth      =   210
      TabIndex        =   37
      Top             =   870
      Width           =   210
   End
   Begin VB.PictureBox picEx 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3960
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   36
      ToolTipText     =   "Exit"
      Top             =   50
      Width           =   140
   End
   Begin VB.PictureBox picCl 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3810
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   35
      ToolTipText     =   "Expand\Collapse"
      Top             =   50
      Width           =   140
   End
   Begin VB.PictureBox picMin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3660
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   34
      ToolTipText     =   "Minimize"
      Top             =   50
      Width           =   140
   End
   Begin VB.PictureBox picOpt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   90
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   33
      ToolTipText     =   "Options"
      Top             =   50
      Width           =   135
   End
   Begin VB.PictureBox D 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1350
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   32
      Top             =   390
      Width           =   135
   End
   Begin VB.PictureBox C 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1170
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   31
      Top             =   390
      Width           =   135
   End
   Begin VB.PictureBox B 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   900
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   30
      Top             =   390
      Width           =   135
   End
   Begin VB.PictureBox A 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   135
      TabIndex        =   29
      Top             =   390
      Width           =   135
   End
   Begin VB.PictureBox pSc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   1660
      ScaleHeight     =   90
      ScaleWidth      =   2310
      TabIndex        =   28
      Top             =   400
      Width           =   2310
   End
   Begin VB.PictureBox imgPlaylist 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3630
      ScaleHeight     =   180
      ScaleWidth      =   345
      TabIndex        =   27
      ToolTipText     =   "Playlist"
      Top             =   870
      Width           =   345
   End
   Begin VB.PictureBox picBal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2660
      ScaleHeight     =   195
      ScaleWidth      =   570
      TabIndex        =   26
      Top             =   850
      Width           =   570
   End
   Begin VB.PictureBox picVol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1600
      ScaleHeight     =   195
      ScaleWidth      =   1020
      TabIndex        =   25
      Top             =   850
      Width           =   1020
   End
   Begin VB.PictureBox picNe 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1620
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   24
      ToolTipText     =   "Next"
      Top             =   1320
      Width           =   330
   End
   Begin VB.PictureBox picSt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1270
      ScaleHeight     =   270
      ScaleWidth      =   345
      TabIndex        =   23
      ToolTipText     =   "Stop"
      Top             =   1320
      Width           =   345
   End
   Begin VB.PictureBox picPa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   930
      ScaleHeight     =   270
      ScaleWidth      =   345
      TabIndex        =   22
      ToolTipText     =   "Pause\Resume"
      Top             =   1320
      Width           =   345
   End
   Begin VB.PictureBox picPl 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   580
      ScaleHeight     =   270
      ScaleWidth      =   345
      TabIndex        =   21
      ToolTipText     =   "Play\Resume"
      Top             =   1320
      Width           =   345
   End
   Begin VB.PictureBox picRv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      ScaleHeight     =   270
      ScaleWidth      =   345
      TabIndex        =   20
      ToolTipText     =   "Previus"
      Top             =   1320
      Width           =   345
   End
   Begin VB.PictureBox picOp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2040
      ScaleHeight     =   240
      ScaleWidth      =   330
      TabIndex        =   19
      ToolTipText     =   "Open"
      Top             =   1330
      Width           =   330
   End
   Begin VB.PictureBox picRe 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3150
      ScaleHeight     =   225
      ScaleWidth      =   420
      TabIndex        =   18
      ToolTipText     =   "Repeat On\Off"
      Top             =   1330
      Width           =   420
   End
   Begin VB.PictureBox picSh 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2470
      ScaleHeight     =   225
      ScaleWidth      =   675
      TabIndex        =   17
      ToolTipText     =   "Shuffle On\Off"
      Top             =   1330
      Width           =   680
   End
   Begin VB.PictureBox imgEq 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3280
      ScaleHeight     =   180
      ScaleWidth      =   345
      TabIndex        =   16
      ToolTipText     =   "Equalizer"
      Top             =   870
      Width           =   345
   End
   Begin VB.PictureBox picM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3180
      ScaleHeight     =   180
      ScaleWidth      =   405
      TabIndex        =   15
      Top             =   610
      Width           =   405
   End
   Begin VB.PictureBox picS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3580
      ScaleHeight     =   180
      ScaleWidth      =   420
      TabIndex        =   14
      Top             =   610
      Width           =   420
   End
   Begin VB.PictureBox Bit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   1660
      ScaleHeight     =   90
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   650
      Width           =   225
   End
   Begin VB.PictureBox Hrz 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   2340
      ScaleHeight     =   90
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   650
      Width           =   225
   End
   Begin VB.PictureBox Titlebar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1300
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   3960
      TabIndex        =   11
      Top             =   3000
      Width           =   3960
   End
   Begin VB.PictureBox Cbuttons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   1680
      ScaleHeight     =   540
      ScaleWidth      =   2400
      TabIndex        =   10
      Top             =   6360
      Width           =   2400
   End
   Begin VB.PictureBox Shufrep 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   1380
      TabIndex        =   9
      Top             =   6360
      Width           =   1380
   End
   Begin VB.PictureBox Balance 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   4200
      ScaleHeight     =   3855
      ScaleWidth      =   1020
      TabIndex        =   8
      Top             =   3000
      Width           =   1020
   End
   Begin VB.PictureBox Volume 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   5400
      ScaleHeight     =   3855
      ScaleWidth      =   1020
      TabIndex        =   7
      Top             =   3000
      Width           =   1020
   End
   Begin VB.PictureBox Posbar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   120
      ScaleHeight     =   150
      ScaleMode       =   0  'User
      ScaleWidth      =   4005
      TabIndex        =   6
      Top             =   4440
      Width           =   4005
   End
   Begin VB.PictureBox Pledit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1095
      ScaleWidth      =   3960
      TabIndex        =   5
      Top             =   5160
      Width           =   3960
   End
   Begin VB.PictureBox Monoster 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3000
      ScaleHeight     =   360
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   4680
      Width           =   1110
   End
   Begin VB.PictureBox Numbers 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1335
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.PictureBox Text 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      ScaleHeight     =   375
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.FileListBox Files 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000012&
      Height          =   225
      Left            =   600
      Pattern         =   "*.mp3;*.wav;*.mid;*.midi;*.wma;*.midi;*.mid"
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Timer tmrScroller 
      Interval        =   200
      Left            =   120
      Top             =   6960
   End
   Begin VB.TextBox Ini 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image imgCP 
      Height          =   230
      Left            =   3650
      ToolTipText     =   "About CoolPlayer"
      Top             =   1330
      Width           =   300
   End
End
Attribute VB_Name = "frmMn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 37 Then Call GotoTime(-5)
    If KeyCode = 39 Then Call GotoTime(5)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    GL.bClick = True
    Call GetMainPar(Button, X, Y)

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo MError
    If GL.bClick = True Then Call MoveForm(frmMn, Button, GL.X, GL.Y, X, Y, True)
    Call TEvent(Tray.TrayEvent(X))

MError:
    If Err.Number <> 0 Then GL.bClick = False: Exit Sub

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo UError
    Call Mpop(Button, X, Y, frmMn, frmMnu.mnuPM)
    Call SaveIniSettings(False)
    GL.bClick = False

UError:
    If Err.Number <> 0 Then GL.bClick = False: Exit Sub

End Sub

Private Sub Form_Terminate()

    On Error Resume Next
    Call Tray.RemoveTrayIcon
    Call MPlay.StopT

End Sub

Private Sub picMBal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bDrag = True: GL.xSli = X
     Call BalDown
     Call DefineBalance
    End If

End Sub
Private Sub picMBal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo MError
    If GL.bDrag = True Then Call MoveBalance(X)

MError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub picMBal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bDrag = False
     Call BalUp
    End If

End Sub
Private Sub picEx_Click()

    If GL.bClick = True Then Call ProgramExit
    GL.bClick = False

End Sub
Private Sub picEx_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picEx, Button, X, Y) = True Then
     Call ExitDown
    Else
     Call ExitUp
    End If

End Sub
Private Sub imgCP_Click()

    If GL.bClick = False Then Exit Sub
    Call LoadfrmAbout
    GL.bClick = False

End Sub
Private Sub imgCP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GL.bClick = IIf(Button = 1, True, False)
End Sub

Private Sub picRe_Click()

    If GL.bClick = False Then Exit Sub
    If CI.bLoop = True Then
     CI.bLoop = False
    ElseIf CI.bLoop = False Then
     CI.bLoop = True
    End If
    GL.bClick = False
    Call SaveIniSettings(False)

End Sub
Private Sub picRe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call LoopDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picRe_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picRe, Button, X, Y) = True Then
     Call LoopDown
    Else
     Call LoopUp
    End If

End Sub
Private Sub picRe_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call LoopUp
End Sub
Private Sub picMin_Click()

    If GL.bClick = False Then Exit Sub
    Call HideForms(False)
    GL.bClick = False

End Sub
Private Sub picMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picMin, Button, X, Y) = True Then
     Call MinDown
    Else
     Call MinUp
    End If

End Sub
Private Sub picNe_Click()

    If GL.bClick = False Then Exit Sub
    Call NextP
    GL.bClick = False

End Sub
Private Sub picNe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call NextDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picNe_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picNe, Button, X, Y) = True Then
     Call NextDown
    Else
     Call NextUp
    End If

End Sub
Private Sub picNe_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call NextUp
End Sub
Private Sub picOp_Click()

    On Error GoTo OError
    If GL.bClick = False Then Exit Sub
    Call OpenForFile(frmMn)
    Call OpenUp
    GL.bClick = False

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub picOp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call OpenDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picOp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picOp, Button, X, Y) = True Then
     Call OpenDown
    Else
     Call OpenUp
    End If

End Sub
Private Sub picOp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call OpenUp
End Sub
Private Sub picOpt_Click()

    If GL.bClick = True Then Call PopupMenu(frmMnu.mnuO, , picOpt.Left + 200, picOpt.Top)
    GL.bClick = False

End Sub
Private Sub picOpt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call AboutDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picOpt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picOpt, Button, X, Y) = True Then
     Call AboutDown
    Else
     Call AboutUp
    End If

End Sub
Private Sub picOpt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call AboutUp
End Sub
Private Sub picPa_Click()

    If GL.bClick = False Then Exit Sub
    Call Pause
    GL.bClick = False

End Sub
Private Sub picPa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call PauseDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picPa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picPa, Button, X, Y) = True Then
     Call PauseDown
    Else
     Call PauseUp
    End If

End Sub
Private Sub picPa_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call PauseUp
End Sub
Private Sub picPl_Click()

    If GL.bClick = True Then Call GetPlay(True)
    GL.bClick = False

End Sub
Private Sub picPl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call PlayDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picPl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picPl, Button, X, Y) = True Then
     Call PlayDown
    Else
     Call PlayUp
    End If

End Sub
Private Sub picPl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call PlayUp
End Sub
Private Sub imgPlaylist_Click()

    If GL.bClick = False Then Exit Sub
    With frmPl
     If CI.BList = True Then
      CI.BList = False
      .Visible = False
     ElseIf CI.BList = False Then
      CI.BList = True
      .Visible = True
     End If
     Call SaveIniSettings(False)
     GL.bClick = False
    End With

End Sub
Private Sub imgPlaylist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call PlDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub imgPlaylist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(imgPlaylist, Button, X, Y) = True Then
     Call PlDown
    Else
     Call PlUp
    End If

End Sub
Private Sub imgPlaylist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call PlUp
End Sub
Private Sub picRv_Click()

    If GL.bClick = False Then Exit Sub
    Call PrevP
    GL.bClick = False

End Sub
Private Sub picRv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call PrevDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picRv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picRv, Button, X, Y) = True Then
     Call PrevDown
    Else
     Call PrevUp
    End If

End Sub
Private Sub picRv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call PrevUp
End Sub
Private Sub picSh_Click()

    If GL.bClick = False Then Exit Sub
    If CI.bRand = True Then
     CI.bRand = False
    ElseIf CI.bRand = False Then
     CI.bRand = True
    End If
    GL.bClick = False
    Call SaveIniSettings(False)

End Sub
Private Sub picSh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call ShuffDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picSh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picSh, Button, X, Y) = True Then
     Call ShuffDown
    Else
     Call ShuffUp
    End If

End Sub
Private Sub picSh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call ShuffUp
End Sub
Private Sub picMSli_Click()

    If GL.bClick = True Then Call PopupMenu(frmTim.mnuTo, , picMSli.Left, picMSli.Top - 1070)
    GL.bClick = False

End Sub
Private Sub picSt_Click()

    On Error GoTo SError
    If GL.bClick = True Then Call StopPlay
    GL.bClick = False

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub picSt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call StopDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picSt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picSt, Button, X, Y) = True Then
     Call StopDown
    Else
     Call StopUp
    End If

End Sub
Private Sub picSt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call StopUp
End Sub
Private Sub picCl_Click()

    If GL.bClick = False Then Exit Sub
    Call MinMain
    GL.bClick = False

End Sub
Private Sub picCl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True: Call TopDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picCl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picCl, Button, X, Y) = True Then
     Call TopDown
    Else
     Call TopUp
    End If

End Sub
Private Sub picCl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call TopUp
End Sub
Private Sub picMVol_Click()

    If GL.bClick = True Then Call PopupMenu(frmMnu.mnuMute, , picMVol.Left, picMVol.Top - 420)
    GL.bClick = False
    Call SaveIniSettings(False)

End Sub
Private Sub Form_Resize()
    Call ListLeft
End Sub
Private Sub picEx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call ExitDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picEx_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call ExitUp
End Sub
Private Sub picMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call MinDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call MinUp
End Sub
Private Sub picMSli_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.sDrag = True: GL.xSli = X
     Call SliderDown
    ElseIf Button = 2 Then
     GL.bClick = True
    End If

End Sub
Private Sub picMSli_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo SError
    If GL.sDrag = True Then Call MoveSlider(X, MPlay.duration)

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub picMSli_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     Call PlaySlider(MoveSlider(X, MPlay.duration))
     Call SliderUp: GL.sDrag = False
    End If

End Sub
Private Sub picMVol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.vDrag = True: GL.xSli = X
     Call VolDown
     Call DefineVolume
    End If

End Sub
Private Sub picMVol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GL.vDrag = True Then Call MoveVolume(X)
End Sub
Private Sub picMVol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.vDrag = False: GL.bClick = False
    ElseIf Button = 2 Then
     GL.bClick = True
    End If
    Call VolUp

End Sub
Private Sub imgEq_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call EqDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub imgEq_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(imgEq, Button, X, Y) = True Then
     Call EqDown
    Else
     Call EqUp
    End If

End Sub
Private Sub imgEq_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     Call EqUp: GL.bClick = False
    End If

End Sub

Private Sub tmrScroller_Timer()

    On Error GoTo TError
    Call ScrollText(frmMn.pSc)
    Call ShowTime
    Call BarMove
    If MPlay.EndOfStream = True Then Call CheckOnEnd

TError:
    If Err.Number <> 0 Then Exit Sub

End Sub
