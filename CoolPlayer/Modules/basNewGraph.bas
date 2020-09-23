Attribute VB_Name = "basNG"
Option Explicit


Public Sub HideForms(Hide As Boolean)

    With frmPl
     If Hide = True Then
      With frmMn
       .WindowState = vbNormal
       .Visible = True
      End With
      If CI.BList = True Then .Visible = True
     ElseIf Hide = False Then
      .Visible = False
      With frmMn
       .WindowState = vbMinimized
       .Visible = False
      End With
     End If
    End With

End Sub
Public Sub ListLeft()

    With frmMn
     frmPl.Top = .Top
     frmPl.Left = .Left + .Width
    End With

End Sub
Public Sub LoadfrmText()

    With frmTxt
     Load frmTxt
     Call DisableForms(False)
     Call Graph.Ontop(.hwnd, CI.bTop)
     .Show
     .txtSearch.SetFocus
    End With

End Sub
Public Sub MinMain()

    With frmMn
     If GL.mMin = False Then
      .Height = 210: GL.mMin = True
      Call .picCl.PaintPicture(.Titlebar, 0, 0, 135, 135, 0, 405, 135, 135)
     ElseIf GL.mMin = True Then
      .Height = 1740: GL.mMin = False
      Call .picCl.PaintPicture(.Titlebar, 0, 0, 135, 135, 0, 270, 135, 135)
     End If
    End With

End Sub
Public Sub MinPlaylist()

    With frmPl
     If GL.lMin = False Then
      .Height = 210: GL.lMin = True
      Call .picCl.PaintPicture(frmMn.Pledit, 0, 0, 135, 135, 1930, 675, 135, 135)
     ElseIf GL.lMin = True Then
      .Height = 6960: GL.lMin = False
      Call .picCl.PaintPicture(frmMn.Pledit, 0, 0, 135, 135, 2370, 45, 135, 135)
     End If
     Call .Scroller.Update
    End With

End Sub

Public Sub RemoveItem()

    On Error GoTo EError
    With frmPl.l
     If .ListItems.Count <> 0 Then
      Call .ListItems.Remove(.SelectedItem.Index)
      Call SetMax
     End If
    End With

EError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub SelList(Sel As Boolean)

    Dim intI As Integer
    With frmSet.lstTypes
     For intI = 0 To .ListCount - 1
      .Selected(intI) = Sel
     Next intI
     If Sel = False Then .ListIndex = -1
    End With

End Sub
Public Sub SetMax()

    With frmPl
     .Scroller.Max = .l.ListItems.Count
     If .l.ListItems.Count <> 0 Then
      .Scroller.Value = .l.SelectedItem.Index
     Else
      .Scroller.Value = 0
     End If
    End With

End Sub
Public Sub SHLabels(Value As Boolean)

    With frmAb
     If Value = True Then .picPar.Top = 2880
     .tmrSc.Enabled = CBool(True - Value)
     .lblAb.FontBold = CBool(True - Value): .lblCP.FontBold = CBool(False - Value)
     .lblMAb.Visible = CBool(False - Value): .picPar.Visible = CBool(True - Value)
    End With

End Sub
Public Sub LoadSkinNumber()

    On Error Resume Next
    Dim O As Object
    Set O = frmMn.Ini

    O.Text = Ini.LoadIni("Sets", "SPath")
    Call CheckPath(O.Text, False)
    O.Text = Ini.LoadIni("Sets", "Skin")

    With frmSkn
     If O.Text = "Error" Then
      .lstSkins.ListIndex = 0
     Else
      If CInt(O.Text) > .lstSkins.ListCount Or CInt(O.Text) < 0 Then
       .lstSkins.ListIndex = 0
      ElseIf CInt(O.Text) < .lstSkins.ListCount And CInt(O.Text) >= 0 Then
       .lstSkins.ListIndex = CInt(O.Text)
      Else
       .lstSkins.ListIndex = 0
      End If
     End If
    End With
    Call SkinIt

End Sub
Public Sub LoadfrmAbout()

    Load frmAb
    Call DisableForms(False)
    Call HideForms(True)
    With frmAb
     Call Graph.Ontop(.hwnd, CI.bTop)
     .lblMAb.Caption = .lblMAb.Caption & vbCrLf & "v " & App.Major & "." & App.Minor & "." & App.Revision
     .Show
    End With

End Sub
Public Sub LoadfrmSettings()

    With frmSet
     Load frmSet
     .lstPlugs.AddItem (.lstPlugs.ListCount + 1) & ". " & File.version
     .TrayI.Picture = frmMn.Icon
     Call DisableForms(False)
     Call HideForms(True)
     Call LoadIniSettings(False, .hwnd)
     Call Graph.Ontop(.hwnd, CI.bTop)
     .Show
    End With

End Sub
Public Sub LoadfrmSkins()

    With frmSkn
     Load frmSkn
     Call DisableForms(False)
     Call LoadSkinNumber
     Call Graph.Ontop(.hwnd, CI.bTop)
     .Show
    End With

End Sub
Public Sub LoadfrmID3(DoIt As Boolean)

    With frmID3
     Load frmID3
     Call IDx.SplitGenres(frmID3.lstGen)
     Call DisableForms(False)
     Call FullName(False)
     If DoIt = True Then Call GetIt(GL.sTrack)
     Call Graph.Ontop(.hwnd, CI.bTop)
     .Show
    End With

End Sub
Public Sub SetNumbers(A As Integer, B As Integer, C As Integer, D As Integer)

    On Error GoTo ErrSet
    With frmMn
     Call .A.PaintPicture(.Numbers, 0, 0, 135, 195, 135 * A, 0, 135, 195)
     Call .B.PaintPicture(.Numbers, 0, 0, 135, 195, 135 * B, 0, 135, 195)
     Call .C.PaintPicture(.Numbers, 0, 0, 135, 195, 135 * C, 0, 135, 195)
     Call .D.PaintPicture(.Numbers, 0, 0, 135, 195, 135 * D, 0, 135, 195)
    End With
    
ErrSet:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub CheckOnEnd()

    On Error Resume Next
    With frmPl.l.ListItems
     If .Count = 0 Then Call StopPlay: Exit Sub
     If GL.lPos = .Count Then
      If CI.bRand = False And CI.bLoop = True Then
       Call SetScroller(1)
       Call GetPlay(True)
      End If
     ElseIf CI.bRand = False And GL.lPos <> .Count Then
      Call SetScroller(GL.lPos + 1)
      Call GetPlay(True)
     End If
    End With
    Call Shuffle(CI.bRand)

End Sub
Public Sub SHFile(S As Boolean)

    With frmPl
     .AddBar.Visible = S
     .picAddDir.Visible = S
     .picAddUrl.Visible = S
    End With

End Sub

Public Sub SHRem(S As Boolean)

    With frmPl
     .SelBar.Visible = S
     .picSelZero.Visible = S
     .picInv.Visible = S
    End With

End Sub
Public Sub SHMisc(S As Boolean)

    With frmPl
     .MisBar.Visible = S
     .picInfo.Visible = S
     .picSort.Visible = S
    End With

End Sub
Public Sub SHList(S As Boolean)

    With frmPl
     .ListBar.Visible = S
     .picLoad.Visible = S
     .picNew.Visible = S
    End With

End Sub
Public Sub SHSel(S As Boolean)

    With frmPl
     .RemBar.Visible = S
     .picCrop.Visible = S
     .picRemMisc.Visible = S
     .picRemAll.Visible = S
    End With

End Sub
Public Sub SkinIt()

    Call SetSkinPath
    Call ReadCredits
    Call SetSkin

End Sub
Public Function SkinString(D As PictureBox, St As String) As String

    On Error Resume Next
    Dim TextX As Long, TextY As Long
    Dim A As Long, B As Long
    Dim C As Long, i As Long
    Dim S3 As String, S2 As String
    Dim S1 As String, S As String

    D.Width = IIf(Len(St) * 75 > 2310, 2310, Len(St) * 75)
    SkinString = St
    S1 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & """" & "@"
    S2 = "0123456789"
    S3 = ".:()-'!_+\/[]^&%,=$#"

    For i = 1 To Len(St)
     S = Mid(UCase(St), i, 1)
     TextX = -1: TextY = -1

     For A = 1 To Len(S1)
      If S = Mid(S1, A, 1) Then
       TextX = (A - 1) * 75: TextY = 0
      End If
     Next A
        
     For B = 1 To Len(S2)
      If S = Mid(S2, B, 1) Then
       TextX = (B - 1) * 75: TextY = 90
      End If
     Next B
        
     For C = 1 To Len(S3)
      If Mid(St, i, 1) = Mid(S3, C, 1) Then
       TextX = ((C - 1) * 75) + 825: TextY = 90
      End If
     Next C

     If Mid(St, i, 1) = " " Or TextX = -1 Then
      TextX = 2145: TextY = 0
     End If
     
     If Mid(St, i, 1) = "*" Then
      TextX = 300: TextY = 180
     End If

     If Mid(St, i, 1) = "?" Then
      TextX = 275: TextY = 180
     End If
     If i <= Len(St) Then Call D.PaintPicture(frmMn.Text, ((i - 1) * 75), 0, 75, 90, TextX, TextY, 75, 90)
    Next i

End Function


Public Sub CheckPath(Path As String, SetIn As Boolean)

    On Error GoTo DefError
    If Path = "" Or Path = "Error" Then
     Call UpdateSkinDir(App.Path & "\Skins\", SetIn)
     CI.sPath = Ini.GetShortPath(App.Path & "\Skins\")
    ElseIf Path <> "" Then
     Call UpdateSkinDir(Ini.GetLongPath(Path), SetIn)
     CI.sPath = Ini.GetShortPath(Path)
    End If

DefError:
    If Err.Number <> 0 Then
     Call UpdateSkinDir(App.Path & "\Skins\", SetIn)
     CI.sPath = Ini.GetShortPath(App.Path & "\Skins\")
    End If

End Sub
Public Sub SHPics(ShowS As Boolean)

    With frmMn
     .A.Visible = ShowS
     .B.Visible = ShowS
     .C.Visible = ShowS
     .D.Visible = ShowS
     .Bit.Visible = ShowS
     .Hrz.Visible = ShowS
     .picMSli.Visible = ShowS
    End With

End Sub
