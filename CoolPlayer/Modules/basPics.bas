Attribute VB_Name = "basPic"
Option Explicit
Public Sub DrawTitleBar()
    
    On Error Resume Next
    Call frmMn.PaintPicture(frmMn.Titlebar, 0, 0, 4125, 210, 405, 0, 4125, 210)

End Sub
Public Sub EqDown()

    On Error Resume Next
    Call frmMn.imgEq.PaintPicture(frmMn.Shufrep, 0, 0, 345, 180, 0, 1095, 345, 180)

End Sub
Public Sub EqUp()

    On Error Resume Next
    Call frmMn.imgEq.PaintPicture(frmMn.Shufrep, 0, 0, 345, 180, 0, 915, 345, 180)

End Sub

Public Sub MinUp()

    On Error Resume Next
    Call frmMn.picMin.PaintPicture(frmMn.Titlebar, 0, 0, 135, 135, 135, 0, 135, 135)

End Sub
Public Sub MinDown()

    On Error Resume Next
    Call frmMn.picMin.PaintPicture(frmMn.Titlebar, 0, 0, 135, 135, 135, 135, 135, 135)

End Sub
Public Sub ExitUp()

    On Error Resume Next
    Call frmMn.picEx.PaintPicture(frmMn.Titlebar, 0, 0, 135, 135, 270, 0, 135, 135)

End Sub
Public Sub ExitDown()

    On Error Resume Next
    Call frmMn.picEx.PaintPicture(frmMn.Titlebar, 0, 0, 135, 135, 270, 135, 135, 135)

End Sub
Public Sub MoveFormX(Frm As Form, OldX As Single, NewX As Single)

    If (Frm.Left + (NewX - OldX) + Frm.Width) / Screen.TwipsPerPixelX > Scr.Right - 20 And (Frm.Left + (NewX - OldX) + Frm.Width) / Screen.TwipsPerPixelX < Scr.Right + 20 Then
     Frm.Left = (Scr.Right * Screen.TwipsPerPixelX) - Frm.Width
    ElseIf (Frm.Left + (NewX - OldX)) / Screen.TwipsPerPixelX < Scr.Left + 20 And (Frm.Left + (NewX - OldX)) / Screen.TwipsPerPixelX > Scr.Left - 20 Then
     Frm.Left = (Scr.Left * Screen.TwipsPerPixelX)
    Else
     Frm.Left = Frm.Left + (NewX - OldX)
    End If

End Sub
Public Sub MoveFormY(Frm As Form, OldY As Single, NewY As Single)

    If (Frm.Top + (NewY - OldY) + Frm.Height) / Screen.TwipsPerPixelY > Scr.Bottom - 20 And (Frm.Top + (NewY - OldY) + Frm.Height) / Screen.TwipsPerPixelY < Scr.Bottom + 20 Then
     Frm.Top = (Scr.Bottom * Screen.TwipsPerPixelY) - Frm.Height
    ElseIf (Frm.Top + (NewY - OldY)) / Screen.TwipsPerPixelY < Scr.Top + 20 And (Frm.Top + (NewY - OldY)) / Screen.TwipsPerPixelY > Scr.Top - 20 Then
     Frm.Top = (Scr.Top * Screen.TwipsPerPixelY)
    Else
     Frm.Top = Frm.Top + (NewY - OldY)
    End If

End Sub
Public Sub NoMode()

    On Error Resume Next
    With frmMn
     If .Monoster.Width = 870 Then
      .picM.Width = 405
      .picM.Left = 3180
      Call .picM.PaintPicture(.Monoster, 0, 0, 405, 180, 435, 180, 405, 180)
     ElseIf .Monoster.Width < 870 Then
      .picM.Width = 375
      .picM.Left = 3195
      Call .picM.PaintPicture(.Monoster, 0, 0, 375, 180, 435, 180, 375, 180)
     End If
     Call .picS.PaintPicture(.Monoster, 0, 0, 420, 180, 0, 180, 420, 180)
    End With

End Sub
Public Sub IsStereo()

    On Error Resume Next
    With frmMn
     If .Monoster.Width = 870 Then
      .picM.Width = 405
      .picM.Left = 3180
      Call .picM.PaintPicture(.Monoster, 0, 0, 405, 180, 435, 180, 405, 180)
     ElseIf .Monoster.Width < 870 Then
      .picM.Width = 375
      .picM.Left = 3195
      Call .picM.PaintPicture(.Monoster, 0, 0, 375, 180, 435, 180, 375, 180)
     End If
     Call .picS.PaintPicture(.Monoster, 0, 0, 420, 180, 0, 0, 420, 180)
    End With

End Sub
Public Sub IsMono()

    On Error Resume Next
    With frmMn
     If .Monoster.Width = 870 Then
      .picM.Width = 405
      .picM.Left = 3180
      Call .picM.PaintPicture(.Monoster, 0, 0, 405, 180, 435, 0, 405, 180)
     ElseIf .Monoster.Width < 870 Then
      .picM.Width = 375
      .picM.Left = 3195
      Call .picM.PaintPicture(.Monoster, 0, 0, 375, 180, 435, 0, 375, 180)
     End If
     Call .picS.PaintPicture(.Monoster, 0, 0, 420, 180, 0, 0, 420, 180)
    End With

End Sub
Public Sub OpenForFile(Frm As Form)

    On Error GoTo FileError
    Call DisableForms(False)

    With ODialog
     .CancelError = True
     .hOwner = Frm.hwnd
     .DialogTitle = "Add file..."
     .FileName = ""
     .Filter = "MP3 Files (*.mp3)|*.mp3|All Media Files (*.mp3)(*.wav)(*.wma)(*.mid)(*.midi)|*.mp3;*.wav;*.wma;*.mid;*.midi;|All Files (*.*)|*.*"
     Call .ShowOpen(CI.bTop)
    End With
    Call DisableForms(True)
    Call AddFile(ODialog.FileName, ODialog.FileTitle, True, True)

FileError:
    If Err.Number <> 0 Then Call DisableForms(True): Exit Sub

End Sub
Public Sub OpenForFolder(Frm As Form)

    On Error GoTo AddDirError
    Dim fFolder As String

    frmMn.Files.Refresh
    Call DisableForms(False)
    fFolder = File.BrowseDir(Frm.hwnd, "Select a path for search.", CI.bTop)
    If Len(fFolder) = 0 Then
     Call DisableForms(True)
     Exit Sub
    End If
    frmMn.Files.Path = fFolder
    Call DisableForms(True)
    Call AddDirectory(fFolder)
    Call Lst.saveM3U(App.Path & Def, frmPl.l)

AddDirError:
    If Err.Number <> 0 Then Call DisableForms(True): Exit Sub

End Sub
Public Sub OpenForLoad(Frm As Form)

    On Error GoTo LoadError
    Call DisableForms(False)
    With ODialog
     .CancelError = True
     .hOwner = Frm.hwnd
     .DialogTitle = "Load playlist..."
     .FileName = ""
     .Filter = "M3U File (*.m3u)|*.m3u|PLS File (*.pls)|*.pls|All Files (*.*)|*.*"
     Call .ShowOpen(CI.bTop)
    End With
    Call DisableForms(True)
    Call LoadList(ODialog.FileName, False)

LoadError:
    If Err.Number <> 0 Then Call DisableForms(True): Exit Sub

End Sub
Public Sub OpenForSave(Frm As Form)

    On Error GoTo SaveError
    Call DisableForms(False)
    With ODialog
     .CancelError = True
     .hOwner = Frm.hwnd
     .DialogTitle = "Save playlist..."
     .FileName = ""
     .FileTitle = ""
     .Filter = "M3U File (*.m3u)|*.m3u|PLS File (*.pls)|*.pls"
     Call .ShowSave(CI.bTop)
    End With
    Call DisableForms(True)
    Call Lst.saveList(ODialog.FileName, ODialog.FilterIndex, frmPl.l)

SaveError:
    If Err.Number <> 0 Then Call DisableForms(True): Exit Sub

End Sub
Public Sub PExitDown()

    On Error Resume Next
    Call frmPl.picExp.PaintPicture(frmMn.Pledit, 0, 0, 135, 135, 780, 630, 135, 135)

End Sub
Public Sub PExitUp()

    On Error Resume Next
    Call frmPl.picExp.PaintPicture(frmMn.Pledit, 0, 0, 135, 135, 2505, 45, 135, 135)

End Sub
Public Sub RefreshSkins(Path As String)

    With frmMn
     .Picture = LoadPicture(Path & "\Main.bmp")
     .Monoster = LoadPicture(Path & "\Monoster.bmp")
     .Cbuttons = LoadPicture(Path & "\Cbuttons.bmp")
     .Shufrep = LoadPicture(Path & "\Shufrep.bmp")
     .Text = LoadPicture(Path & "\Text.bmp")
     .Pledit = LoadPicture(Path & "\Pledit.bmp")
     .Posbar = LoadPicture(Path & "\Posbar.bmp")
     .Titlebar = LoadPicture(Path & "\Titlebar.bmp")
     .Volume = LoadPicture(Path & "\Volume.bmp")

     If Dir(Path & "\Numbers.bmp") <> "" And Dir(Path & "\Nums_Ex.bmp") <> "" Then
      .Numbers = LoadPicture(Path & "\Nums_ex.bmp")
     ElseIf Dir(Path & "\Nums_Ex.bmp") = "" And Dir(Path & "\Numbers.bmp") <> "" Then
      .Numbers = LoadPicture(Path & "\Numbers.bmp")
     Else
      .Numbers = LoadPicture(Path & "\Nums_Ex.bmp")
     End If

     If Dir(Path & "\Balance.bmp") <> "" Then
      .Balance = LoadPicture(Path & "\Balance.bmp")
     Else
      .Balance = LoadPicture(Path & "\Volume.bmp")
     End If

     .picMSli.Height = .Posbar.Height
     If .Numbers.Height > 195 Then .Numbers.Height = 195
     .A.Height = .Numbers.Height
     .B.Height = .A.Height
     .C.Height = .A.Height
     .D.Height = .A.Height
    End With

End Sub
Public Sub TExitUp()

    On Error Resume Next
    Call frmPl.picCl.PaintPicture(frmMn.Pledit, 0, 0, 135, 135, 2370, 45, 135, 135)

End Sub
Public Sub TExitDown()

    On Error Resume Next
    With frmPl.picCl
     If GL.lMin = False Then
      Call .PaintPicture(frmMn.Pledit, 0, 0, 135, 135, 930, 630, 135, 135)
     ElseIf GL.lMin = True Then
      Call .PaintPicture(frmMn.Pledit, 0, 0, 135, 135, 2250, 630, 135, 135)
     End If
    End With

End Sub
Public Sub BarMove()

    On Error GoTo PError
    Dim OffSet As Integer

    With MPlay
     OffSet = 3285 * (.CurrentPosition / .duration)
     If GL.sDrag = False Then
      frmMn.picMSli.Left = OffSet + 240
      Call SliderUp
     End If
    End With

PError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub ShowTime()
    
    On Error GoTo TError
    Call TimePosition(MPlay.CurrentPosition, True)
    Call SkinString(frmPl.Cn, CStr(frmPl.l.ListItems.Count))

TError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub SliderUp()

    On Error Resume Next
    Call frmMn.picMSli.PaintPicture(frmMn.Posbar, 0, 0, 435, frmMn.Posbar.Height, 3720, 0, 435, frmMn.Posbar.Height)

End Sub
Public Sub SliderDown()

    On Error Resume Next
    Call frmMn.picMSli.PaintPicture(frmMn.Posbar, 0, 0, 435, frmMn.Posbar.Height, 4170, 0, 435, frmMn.Posbar.Height)

End Sub
Public Sub VolUp()

    On Error Resume Next
    Dim intS As Integer

    With frmMn
     If .Volume.Height > 6330 Then
      intS = .Volume.Height - 6330
      .picMVol.Height = intS
      Call .picMVol.PaintPicture(.Volume, 0, 0, 210, intS, 225, 6330, 210, intS)
     ElseIf .Volume.Height <= 6330 Then
      .picMVol.Height = 165
      Call .picMVol.PaintPicture(.Volume, 0, 0, 210, 165, 810, 6090, 210, 165)
     End If
    End With

End Sub
Public Sub VolDown()

    On Error Resume Next
    Dim intS As Integer

    With frmMn
     If .Volume.Height > 6330 Then
      intS = .Volume.Height - 6330
      .picMVol.Height = intS
      Call .picMVol.PaintPicture(.Volume, 0, 0, 210, intS, 0, 6330, 210, intS)
     ElseIf .Volume.Height <= 6330 Then
      .picMVol.Height = 165
      Call .picMVol.PaintPicture(.Volume, 0, 0, 210, 165, 810, 6090, 210, 165)
     End If
    End With

End Sub
Public Sub BalDown()

    On Error Resume Next
    Dim S As Integer

    With frmMn
     If .Balance.Height > 6330 Then
      S = .Balance.Height - 6330
      .picMBal.Height = S
      Call .picMBal.PaintPicture(.Balance, 0, 0, 210, S, 0, 6330, 210, S)
     ElseIf .Balance.Height <= 6330 Then
      .picMBal.Height = 165
      Call .picMBal.PaintPicture(.Balance, 0, 0, 210, 165, 135, 6090, 210, 165)
     End If
    End With

End Sub
Public Sub BalUp()

    On Error Resume Next
    Dim S As Integer

    With frmMn
     If .Balance.Height > 6330 Then
      S = .Balance.Height - 6330
      .picMBal.Height = S
      Call .picMBal.PaintPicture(.Balance, 0, 0, 210, S, 225, 6330, 210, S)
     ElseIf .Balance.Height <= 6330 Then
      .picMBal.Height = 165
      Call .picMBal.PaintPicture(.Balance, 0, 0, 210, 165, 135, 6090, 210, 165)
     End If
    End With

End Sub
Public Sub BackPic()

    On Error Resume Next
    With frmMn
     Call .PaintPicture(.Posbar, 240, 1080, 3720, .Posbar.Height, 0, 0, 3720, .Posbar.Height)
    End With

End Sub
Public Sub BackBalance()

    On Error Resume Next
    With frmMn
     Call .picBal.PaintPicture(.Picture, 0, 0, 570, 195, 2660, 855, 570, 195)
     Call .picBal.PaintPicture(.Balance, 0, 0, 570, 195, 130, 0, 570, 195)
    End With

End Sub
Public Sub BackVolume()

    On Error Resume Next
    With frmMn
     Call .picVol.PaintPicture(.Picture, 0, 0, 1020, 195, 1605, 855, 1020, 195)
     Call .picVol.PaintPicture(.Volume, 0, 0, 1020, 195, 0, 0, 1020, 195)
    End With

End Sub
Public Sub PlDown()

    On Error Resume Next
    Call frmMn.imgPlaylist.PaintPicture(frmMn.Shufrep, 0, 0, 345, 180, 345, 1095, 345, 180)

End Sub
Public Sub PlUp()

    On Error Resume Next
    Call frmMn.imgPlaylist.PaintPicture(frmMn.Shufrep, 0, 0, 345, 180, 345, 915, 345, 180)

End Sub
Public Sub PrevUp()

    On Error Resume Next
    Call frmMn.picRv.PaintPicture(frmMn.Cbuttons, 0, 0, 345, 270, 0, 0, 345, 270)

End Sub

Public Sub PrevDown()

    On Error Resume Next
    Call frmMn.picRv.PaintPicture(frmMn.Cbuttons, 0, 0, 345, 270, 0, 270, 345, 270)

End Sub
Public Sub PlayUp()

    On Error Resume Next
    Call frmMn.picPl.PaintPicture(frmMn.Cbuttons, 0, 0, 345, 270, 345, 0, 345, 270)

End Sub
Public Sub PlayDown()

    On Error Resume Next
    Call frmMn.picPl.PaintPicture(frmMn.Cbuttons, 0, 0, 345, 270, 345, 270, 345, 270)

End Sub
Public Sub StopUp()

    On Error Resume Next
    Call frmMn.picSt.PaintPicture(frmMn.Cbuttons, 0, 0, 345, 270, 1035, 0, 345, 270)

End Sub
Public Sub StopDown()

    On Error Resume Next
    Call frmMn.picSt.PaintPicture(frmMn.Cbuttons, 0, 0, 345, 270, 1035, 270, 345, 270)

End Sub
Public Sub PauseUp()

    On Error Resume Next
    Call frmMn.picPa.PaintPicture(frmMn.Cbuttons, 0, 0, 345, 270, 690, 0, 345, 270)

End Sub
Public Sub PauseDown()

    On Error Resume Next
    Call frmMn.picPa.PaintPicture(frmMn.Cbuttons, 0, 0, 345, 270, 690, 270, 345, 270)

End Sub
Public Sub NextUp()

    On Error Resume Next
    Call frmMn.picNe.PaintPicture(frmMn.Cbuttons, 0, 0, 345, 270, 1380, 0, 345, 270)

End Sub
Public Sub NextDown()

    On Error Resume Next
    Call frmMn.picNe.PaintPicture(frmMn.Cbuttons, 0, 0, 345, 270, 1380, 270, 345, 270)

End Sub
Public Sub OpenUp()

    On Error Resume Next
    Call frmMn.picOp.PaintPicture(frmMn.Cbuttons, 0, 0, 330, 240, 1710, 0, 330, 240)

End Sub
Public Sub OpenDown()

    On Error Resume Next
    Call frmMn.picOp.PaintPicture(frmMn.Cbuttons, 0, 0, 330, 240, 1710, 240, 330, 240)

End Sub
Public Sub PlFileUp()

    On Error Resume Next
    Call frmPl.picFile.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 210, 1200, 330, 270)
    'Call frmPl.picFile.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 0, 2235, 330, 270)

End Sub
Public Sub PlFileDown()

    On Error Resume Next
    Call frmPl.picFile.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 0, 2235, 330, 270)
    'Call frmPl.picFile.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 345, 2235, 330, 270)

End Sub
Public Sub PlOptUp()

    On Error Resume Next
    Call frmPl.picOpt.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 1515, 1200, 330, 270)
    'Call frmPl.picOpt.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 2310, 2235, 330, 270)

End Sub
Public Sub PlOptDown()

    On Error Resume Next
    Call frmPl.picOpt.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 2310, 2235, 330, 270)
    'Call frmPl.picOpt.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 2655, 2235, 330, 270)

End Sub
Public Sub PlListUp()

    On Error Resume Next
    Call frmPl.picList.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 3480, 1200, 330, 270)
    'Call frmPl.picList.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 3065, 2235, 330, 270)

End Sub
Public Sub PlRemUp()

    On Error Resume Next
    Call frmPl.picRem.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 650, 1200, 330, 270)
    'Call frmPl.picRem.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 815, 2235, 330, 270)

End Sub
Public Sub PlRemDown()

    On Error Resume Next
    Call frmPl.picRem.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 810, 2235, 330, 270)
    'Call frmPl.picRem.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 1155, 2235, 330, 270)

End Sub
Public Sub PlListDown()

    On Error Resume Next
    Call frmPl.picList.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 3060, 2235, 330, 270)
    'Call frmPl.picList.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 3405, 2235, 330, 270)

End Sub
Public Sub LoopUp()

    On Error Resume Next
    Call frmMn.picRe.PaintPicture(frmMn.Shufrep, 0, 0, 420, 225, 0, 0, 420, 225)

End Sub
Public Sub ShuffUp()

    On Error Resume Next
    Call frmMn.picSh.PaintPicture(frmMn.Shufrep, 0, 0, 675, 225, 435, 0, 675, 225)

End Sub
Public Sub ShuffDown()

    On Error Resume Next
    Call frmMn.picSh.PaintPicture(frmMn.Shufrep, 0, 0, 675, 225, 435, 450, 675, 225)

End Sub
Public Sub LoopDown()

    On Error Resume Next
    Call frmMn.picRe.PaintPicture(frmMn.Shufrep, 0, 0, 420, 225, 0, 450, 420, 225)

End Sub
Public Sub PlTrackUp()

    On Error Resume Next
    Call frmPl.picTrack.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 1080, 1200, 330, 270)
    'Call frmPl.picTrack.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 1560, 2235, 330, 270)

End Sub
Public Sub PlTrackDown()

    On Error Resume Next
    Call frmPl.picTrack.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 1560, 2235, 330, 270)
    'Call frmPl.picTrack.PaintPicture(frmMn.Pledit, 0, 0, 330, 270, 1905, 2235, 330, 270)

End Sub
Public Sub TopUp()

    On Error Resume Next
    Call frmMn.picCl.PaintPicture(frmMn.Titlebar, 0, 0, 135, 135, 0, 270, 135, 135)

End Sub
Public Sub TopDown()

    On Error Resume Next
    With frmMn
     If GL.mMin = False Then
      Call .picCl.PaintPicture(.Titlebar, 0, 0, 135, 135, 135, 270, 135, 135)
     ElseIf GL.mMin = True Then
      Call .picCl.PaintPicture(.Titlebar, 0, 0, 135, 135, 135, 405, 135, 135)
     End If
    End With

End Sub
Public Sub AboutUp()

    On Error Resume Next
    Call frmMn.picOpt.PaintPicture(frmMn.Titlebar, 0, 0, 135, 135, 0, 0, 135, 135)

End Sub
Public Sub AboutDown()

    On Error Resume Next
    Call frmMn.picOpt.PaintPicture(frmMn.Titlebar, 0, 0, 135, 135, 0, 135, 135, 135)

End Sub
