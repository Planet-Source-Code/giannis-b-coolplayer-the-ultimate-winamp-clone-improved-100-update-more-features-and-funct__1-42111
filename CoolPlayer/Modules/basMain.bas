Attribute VB_Name = "basMn"
Option Explicit

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Type Dimension
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private lngi As Long
Public Scr As Dimension
Public Sub DefineBalance()

    On Error GoTo DError
    Dim s As Long

    s = CInt(100 - (GL.sOff * 200))
    With frmMn
     If (GL.sOff * 100) = 50 Then
      MPlay.Balance = 0
      Call SkinString(.pSc, "Balance: Center")
     ElseIf (GL.sOff * 100) < 50 Then
      MPlay.Balance = -s
      Call SkinString(.pSc, "Balance: " & s & "% left")
     ElseIf (GL.sOff * 100) > 50 Then
      MPlay.Balance = -s
      Call SkinString(.pSc, "Balance: " & -s & "% right")
     End If
    End With

DError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub GetMainPar(Button As Integer, X As Single, Y As Single)

    If Button = 1 Then
     Call SystemParametersInfo(48, 0&, Scr, 0&)
     GL.X = X
     GL.Y = Y
    End If

End Sub
Public Sub MoveForm(Frm As Form, Button As Integer, sX As Single, sY As Single, X As Single, Y As Single, Optional OK As Boolean)

    If Button = 1 Then
     If CI.bSnap = True Then
      Call MoveFormX(Frm, sX, X)
      Call MoveFormY(Frm, sY, Y)
     Else
      Call Frm.Move(Frm.Left + X - sX, Frm.Top + Y - sY)
     End If
     If OK = True Then
      Call ListLeft
     End If
    End If

End Sub
Public Sub DefineVolume()

    On Error GoTo DError
    With frmMn
     If CI.bMute = False Then
      Call SkinString(.pSc, "Volume: " & MPlay.Volume & "%")
     Else
      Call SkinString(.pSc, "Mute is on")
     End If
    End With

DError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub GotoTime(Value As Integer)

    On Error Resume Next
    With MPlay
     If .CurrentPosition + Value >= .duration Then Exit Sub
     If .CurrentPosition + Value <= 0 Then Exit Sub
     .CurrentPosition = .CurrentPosition + Value
    End With

End Sub
Public Sub LoadMain()

    On Error GoTo MError
    Dim intS As Integer

    With frmMn
     Load frmMn
     Call LoadIniSettings(True, .hwnd)

     If CI.bAss = True Then
      For intS = 1 To 8
       Call Reg.PublicReg(intS, App.Path, App.EXEName)
      Next intS
     End If
     Call Tray.AddTrayIcon(.Icon, .hwnd, "CoolPlayer by John")
     Call LoadPictures
     .Show
    End With

MError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub ButtonChoice(BP As String)

    With frmMnu
     Select Case BP
      Case "paused"
       .mnuTPa.Caption = "&Resume"

      Case "resumed"
       .mnuTPa.Caption = "Pa&use"

      Case "played"
       .mnuTPa.Enabled = True
       .mnuTPa.Caption = "Pa&use"

      Case "stopped"
       .mnuTPa.Caption = "Pa&use"
       .mnuTPa.Enabled = False

       With frmMn
        Call SkinString(.Bit, "0")
        Call SkinString(.Hrz, "0")
        Call NoMode
       End With
     End Select
    End With

End Sub
Public Sub MoveBalance(X As Single)

    On Error GoTo MError
    Call BalDown
    Call MoveOBJ(X, frmMn.picMBal, 2660, 570, 0)
    Call DefineBalance

MError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub MoveOBJ(X As Single, M As Object, Pleft As Long, PWidth As Long, D As Long)

     Dim Pos As Long
     Pos = M.Left + X - GL.xSli
     Pos = IIf(Pos < Pleft, Pleft, Pos)
     M.Left = IIf(Pos > PWidth + Pleft - D - M.Width, PWidth + Pleft - D - M.Width, Pos)
     GL.sOff = ((M.Left - Pleft) / (PWidth - D - M.Width))

End Sub
Public Function MoveSlider(X As Single, Dur As Long) As Integer

    On Error GoTo MError

    If Dur = 0 Then Exit Function
    Call MoveOBJ(X, frmMn.picMSli, 240, 3720, 0)
    Call SkinString(frmMn.pSc, "Seek to:  " & TimePosition(CInt(Dur * GL.sOff)) & _
                   "/" & File.gettime(CStr(Dur)) & "  (" & CInt(GL.sOff * 100) & "%) ")
    MoveSlider = CInt(Dur * GL.sOff)
    Call SliderDown

MError:
    If Err.Number <> 0 Then Exit Function

End Function
Public Sub MoveVolume(X As Single)

    On Error GoTo MError
    Call VolDown
    Call MoveOBJ(X, frmMn.picMVol, 1600, 1020, 40)
    MPlay.Volume = CInt(100 - (100 * GL.sOff))
    If GL.vDrag = True Then Call DefineVolume '(CInt((100 * GL.sOff) / 4.34))

MError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub NameLabels(N As String, D As Long)

    On Error GoTo NError
    If D = 0 Then
     Call Tray.ToolTip(SkinString(frmMn.pSc, "CoolPlayer by John"))
    ElseIf D <> 0 Then
     Call SkinString(frmMn.pSc, SetText(N, D, CI.bScroll))
    End If

NError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Function SetText(N As String, l As Long, C As Boolean)

    Dim s As String, i As Long
    s = N & " (" & File.gettime(CStr(l)) & ")"

    Call Tray.ToolTip(s)
    If Len(s) <= 31 Then
     SetText = s
    Else
     If C = True Then
      s = "  ***  " & s
      lngi = lngi - 1
      lngi = IIf(lngi < 1, Len(s), lngi)
      i = IIf(lngi <= Len(s), Len(s) - lngi, 1)
      s = Right(s, lngi) & Left(s, i)
     End If
     SetText = s
    End If

End Function
Public Sub PlaySlider(Value As Integer)

    With MPlay
     If Value >= CInt(.duration) Or .Playstate = 2 Then Exit Sub
     .CurrentPosition = Value
    End With

End Sub
Public Sub ReadCredits()

    On Error GoTo RError
    Dim strData As String

    With frmSkn
     .txtInfo.Text = ""
     Close #1
     Open .Files.Path & "\" & .Files.List(0) For Input As #1
      Do While Not EOF(1)
       Line Input #1, strData
       .txtInfo.Text = .txtInfo.Text & strData & vbCrLf
      Loop
     Close #1
    End With

RError:
    If Err.Number <> 0 Then Close #1: Exit Sub

End Sub
Public Sub SetSkin()

    With frmSkn
     If .Dirs.ListIndex < 0 Then .Dirs.ListIndex = 0
     Call RefreshSkins(.Dirs.List(.Dirs.ListIndex))
    End With

    Call LoadPictures
    Call frmPl.Scroller.Update
    With frmMn
     With MP3
      .FileName = MPlay.FileName
      .GetHeader
      If Not .ValidHeader Then Call NoMode
     End With
     Call SkinString(.Bit, MP3.Bitrate)
     Call SkinString(.Hrz, Left(MP3.Frequency, 2))
     Call Ster
    End With
    Call GetAllColors

End Sub
Public Sub SetSkinPath()

    On Error Resume Next
    With frmSkn
     .Dirs.ListIndex = .lstSkins.ListIndex
     .Files.Path = .Dirs.List(.Dirs.ListIndex) & "\"
    End With

End Sub
Public Sub UpdateSkinDir(Path As String, SetIn As Boolean)

    On Error GoTo UpError
    Dim i As Integer

    With frmSkn
     .Dirs.Path = Path
     .Dirs.Refresh
     .lstSkins.Clear
     For i = 0 To .Dirs.ListCount
      If .Dirs.List(i) <> "" Then .lstSkins.AddItem Right(.Dirs.List(i), Len(.Dirs.List(i)) - InStrRev(.Dirs.List(i), "\"))
     Next i
     .fraSkins.Caption = "Available skins..." & .lstSkins.ListCount
     Call SetSkinPath
     If SetIn = True Then .lstSkins.ListIndex = 0
    End With

UpError:
    If Err.Number <> 0 Then Exit Sub

End Sub
