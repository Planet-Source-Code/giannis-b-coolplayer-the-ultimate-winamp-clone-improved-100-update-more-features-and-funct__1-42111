VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPl 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "CoolPlayer Playlist Editor"
   ClientHeight    =   6960
   ClientLeft      =   9900
   ClientTop       =   1305
   ClientWidth     =   4125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleMode       =   0  'User
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Cn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   1980
      ScaleHeight     =   90
      ScaleWidth      =   225
      TabIndex        =   25
      Top             =   6540
      Width           =   225
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3470
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   24
      Top             =   6510
      Width           =   330
   End
   Begin VB.PictureBox picOpt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1515
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   23
      Top             =   6510
      Width           =   330
   End
   Begin VB.PictureBox picTrack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1080
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   22
      Top             =   6510
      Width           =   330
   End
   Begin VB.PictureBox picRem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   645
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   21
      Top             =   6510
      Width           =   330
   End
   Begin VB.PictureBox picFile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   210
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   20
      Top             =   6510
      Width           =   330
   End
   Begin VB.PictureBox picExp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3960
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   19
      ToolTipText     =   "Close"
      Top             =   50
      Width           =   140
   End
   Begin VB.PictureBox picCl 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3820
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   18
      ToolTipText     =   "Expand\Collapse"
      Top             =   50
      Width           =   135
   End
   Begin VB.PictureBox picNew 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3470
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   15
      Top             =   6030
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picLoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3470
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   14
      Top             =   6300
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picSort 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1520
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   13
      Top             =   6030
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1520
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   12
      Top             =   6300
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1080
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   11
      Top             =   6030
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picSelZero 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1080
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   10
      Top             =   6300
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picRemMisc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   640
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picRemAll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   640
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   8
      Top             =   6030
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picCrop 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   640
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   7
      Top             =   6300
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picAddUrl 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   210
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   6
      Top             =   6030
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picAddDir 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   210
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   5
      Top             =   6300
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox ListBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   3420
      ScaleHeight     =   810
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   6030
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox MisBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   1470
      ScaleHeight     =   810
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   6030
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox SelBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   1030
      ScaleHeight     =   810
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   6030
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox RemBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   600
      ScaleHeight     =   1080
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox AddBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   160
      ScaleHeight     =   810
      ScaleWidth      =   45
      TabIndex        =   0
      Top             =   6030
      Visible         =   0   'False
      Width           =   45
   End
   Begin MSComctlLib.ListView l 
      Height          =   6090
      Left            =   180
      TabIndex        =   16
      Top             =   300
      Width           =   3650
      _ExtentX        =   6429
      _ExtentY        =   10742
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   65280
      BackColor       =   -2147483641
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5186
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Object.Width           =   1235
      EndProperty
   End
   Begin vbpCoolPlayer.ctlScroller Scroller 
      Height          =   6090
      Left            =   3830
      TabIndex        =   17
      Top             =   300
      Width           =   300
      _extentx        =   529
      _extenty        =   10742
   End
   Begin VB.Image imgOp 
      Height          =   135
      Left            =   2670
      Top             =   6740
      Width           =   135
   End
   Begin VB.Image imgNe 
      Height          =   135
      Left            =   2520
      Top             =   6740
      Width           =   135
   End
   Begin VB.Image imgSt 
      Height          =   135
      Left            =   2370
      Top             =   6740
      Width           =   135
   End
   Begin VB.Image imgPa 
      Height          =   135
      Left            =   2220
      Top             =   6740
      Width           =   135
   End
   Begin VB.Image imgPl 
      Height          =   135
      Left            =   2070
      Top             =   6740
      Width           =   135
   End
   Begin VB.Image imgPr 
      Height          =   135
      Left            =   1920
      Top             =   6740
      Width           =   135
   End
End
Attribute VB_Name = "frmPl"
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

    If Button = 1 Then GL.bClick = True
    Call GetMainPar(Button, X, Y)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GL.bClick = True Then Call MoveForm(frmPl, Button, GL.X, GL.Y, X, Y)
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo MError
    If Button = 1 Then GL.bClick = False
    If Button = 2 Then Call Mpop(Button, X, Y, frmPl, frmMnu.mnuPP)
    Call SaveIniSettings(False)

MError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub imgNe_Click()
    Call NextP
End Sub
Private Sub imgOp_Click()
    Call OpenForFile(frmMn)
End Sub
Private Sub imgPa_Click()
    Call Pause
End Sub
Private Sub imgPl_Click()
    Call GetPlay(True)
End Sub
Private Sub imgPr_Click()
    Call PrevP
End Sub
Private Sub l_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Scroller.Value = Item.Index
End Sub
Private Sub l_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 37 Then Call GotoTime(-5)
    If KeyCode = 39 Then Call GotoTime(5)

End Sub

Private Sub l_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GL.bClick = False
End Sub
Private Sub picRem_Click()

    If GL.bClick = True Then Call PopupMenu(frmMnu.mnuMisc, , picRem.Left - 50, picRem.Top - 800)
    'Call SHSel(True)
    GL.bClick = False
    
End Sub
Private Sub picRem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call PlRemDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picRem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picRem, Button, X, Y) = True Then
     Call PlRemDown
    Else
     Call PlRemUp
    End If

End Sub
Private Sub picRem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call PlRemUp
End Sub
Private Sub imgSt_Click()
    Call StopPlay
End Sub
Private Sub Scroller_Scroll()
    l.ListItems.Item(Scroller.Value).EnsureVisible
End Sub
Private Sub picExp_Click()

    If GL.bClick = False Then Exit Sub
    If CI.BList = True Then
     frmPl.Visible = False
     CI.BList = False
    End If
    Call SaveIniSettings(False)
    GL.bClick = False

End Sub
Private Sub picExp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call PExitDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picExp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picExp, Button, X, Y) = True Then
     Call PExitDown
    Else
     Call PExitUp
    End If

End Sub
Private Sub picExp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call PExitUp
End Sub
Private Sub picFile_Click()

    If GL.bClick = True Then Call PopupMenu(frmMnu.mnuF, , picFile.Left - 50, picFile.Top - 800)
    GL.bClick = False

End Sub
Private Sub picFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call PlFileDown
     'Call SHFile(True)
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picFile, Button, X, Y) = True Then
     Call PlFileDown
    Else
     Call PlFileUp
    End If

End Sub
Private Sub picFile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then Call PlFileUp
     'Call SHFile(False)

End Sub
Private Sub picList_Click()

    If GL.bClick = True Then Call PopupMenu(frmMnu.mnuL, , picList.Left - 1230, picList.Top - 1310)
    'Call SHList(True)
    GL.bClick = False

End Sub
Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call PlListDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picList, Button, X, Y) = True Then
     Call PlListDown
    Else
     Call PlListUp
    End If

End Sub
Private Sub picList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call PlListUp
End Sub
Private Sub picOpt_Click()

    If GL.bClick = True Then Call PopupMenu(frmMnu.mnuP, , picOpt.Left - 50, picOpt.Top - 1440)
    'Call SHMisc(True)
    GL.bClick = False

End Sub
Private Sub picOpt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call PlOptDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picOpt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picOpt, Button, X, Y) = True Then
     Call PlOptDown
    Else
     Call PlOptUp
    End If

End Sub
Private Sub picOpt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call PlOptUp
End Sub
Private Sub picCl_Click()

    If GL.bClick = False Then Exit Sub
    Call MinPlaylist
    GL.bClick = False

End Sub
Private Sub picCl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True: Call TExitDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picCl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picCl, Button, X, Y) = True Then
     Call TExitDown
    Else
     Call TExitUp
    End If

End Sub
Private Sub picCl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call TExitUp
End Sub
Private Sub picTrack_Click()

    If GL.bClick = True Then Call PopupMenu(frmMnu.mnuT, , picTrack.Left - 50, picTrack.Top - 2080)
    'Call SHRem(True)
    GL.bClick = False

End Sub

Private Sub picTrack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
     GL.bClick = True
     Call PlTrackDown
    ElseIf Button = 2 Then
     GL.bClick = False
    End If

End Sub
Private Sub picTrack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Allow(picTrack, Button, X, Y) = True Then
     Call PlTrackDown
    Else
     Call PlTrackUp
    End If

End Sub
Private Sub picTrack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call PlTrackUp
End Sub
Private Sub l_KeyDown(KeyCode As Integer, Shift As Integer)

    Call SetMax
    If KeyCode = 46 Then Call RemoveItem

End Sub
Private Sub l_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
     Call Mpop(Button, X, Y, frmPl, frmMnu.mnuPP)
    ElseIf Button = 1 Then
     If X >= 0 And X <= l.Width And Y >= 0 And Y <= l.Height Then
       If CI.bClick Then Call GetPlay(True)
     End If
    End If

End Sub
Private Sub l_DblClick()

    On Error GoTo DError
    Call GetPlay(True)

DError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub l_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Call GetPlay(True)
    If KeyAscii = 27 Then Call StopPlay

End Sub
Private Sub l_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo DError
    Dim i As Integer, s As String

    For i = 1 To Data.Files.Count
     s = Data.Files(i)
     If Lst.getext(s) = "m3u" Or Lst.getext(s) = "pls" Then
      Call LoadList(s, False): Exit For
     Else
      Call AddFile(s, Right(s, Len(s) - InStrRev(s, "\")), False, False)
     End If
    Next i
    Call SetScroller(l.ListItems.Count)
    Call Lst.saveM3U(App.Path & Def, frmPl.l)

DError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub l_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    On Error GoTo OError
    If Data.GetFormat(vbCFFiles) And Button = 1 Then
     Effect = vbDropEffectCopy
    Else
     Effect = vbDropEffectNone
    End If

OError:
    If Err.Number <> 0 Then Exit Sub

End Sub
