Attribute VB_Name = "basIR"
Option Explicit

Public Type Ini
    iIcon As Integer
    iSkin As Integer
    yLay As Byte

    BList As Boolean
    bRand As Boolean
    bSnap As Boolean
    bGraph As Boolean
    bClick As Boolean
    bSplash As Boolean
    bSort As Boolean
    bInst As Boolean
    bAss As Boolean
    bMute As Boolean
    bTop As Boolean
    bLoop As Boolean
    bDoub As Boolean
    bTray As Boolean
    bScroll As Boolean

    sPath As String
End Type

Public CI As Ini
Public Sub CheckSplash()

    On Error Resume Next
    With frmMn
     .Ini.Text = Ini.LoadIni("Sets", "Splash")
     If .Ini.Text = "true" Then
      CI.bSplash = True
     ElseIf .Ini.Text = "false" Then
      CI.bSplash = False
     End If

     .Ini.Text = Ini.LoadIni("Sets", "Inst")
     If .Ini.Text = "true" Then
      CI.bInst = True
     ElseIf .Ini.Text = "false" Then
      CI.bInst = False
     End If

     If App.PrevInstance = True Then
      If CI.bInst = False Then End
     End If
 
     .Ini.Text = Ini.LoadIni("List", "Dubs")
     If .Ini.Text = "true" Then
      CI.bDoub = True
     ElseIf .Ini.Text = "false" Then
      CI.bDoub = False
     End If
 
     .Ini.Text = Ini.LoadIni("List", "Sort")
     If .Ini.Text = "true" Then
      CI.bSort = True
     ElseIf .Ini.Text = "false" Then
      CI.bSort = False
     End If
    End With

    With frmSp
     If CI.bSplash = True Then
      .lblDate.Caption = Date
      .Visible = True
     ElseIf CI.bSplash = False Then
      .Visible = False
     End If
    End With
    Call NoMode

End Sub
Public Sub LoadTransparency(Apply As Boolean)

    On Error Resume Next
    Dim O As Object
    Set O = frmSet

    If Ini.LoadIni("Sets", "Trans") = 0 Or Ini.LoadIni("Sets", "Trans") > 255 Then
     If Apply = True Then Call SetLay(255)
     CI.yLay = 255
    End If

    CI.yLay = CByte(Ini.LoadIni("Sets", "Trans"))
    If Apply = True Then Call SetLay(CI.yLay)
    frmSet.sliT.Value = CInt(CI.yLay)

    CI.iIcon = CInt(Ini.LoadIni("Sets", "Icon"))
    frmSet.sliI.Value = CInt(CI.iIcon)
    Call ChangeIcon

    With frmMn
     .Ini.Text = Ini.LoadIni("Sets", "Tray")
     If .Ini.Text = "true" Then
      O.chkMin.Value = 1
      CI.bTray = True
      If Apply = True Then Call HideForms(False)
     ElseIf .Ini.Text = "false" Then
      O.chkMin.Value = 0
      CI.bTray = False
      If Apply = True Then Call HideForms(True)
     End If

     .Ini.Text = Ini.LoadIni("Sets", "Inst")
     If .Ini.Text = "true" Then
      O.chkInst.Value = 1
      CI.bInst = True
     ElseIf .Ini.Text = "false" Then
      O.chkInst.Value = 0
      CI.bInst = False
     End If

     .Ini.Text = Ini.LoadIni("List", "Sort")
     If .Ini.Text = "true" Then
      O.chkSort.Value = 1
      CI.bSort = True
     ElseIf .Ini.Text = "false" Then
      O.chkSort.Value = 0
      CI.bSort = False
     End If

     .Ini.Text = Ini.LoadIni("List", "Single")
     If .Ini.Text = "true" Then
      O.chkSingl.Value = 1
      CI.bClick = True
     ElseIf .Ini.Text = "false" Then
      O.chkSingl.Value = 0
      CI.bClick = False
     End If

     .Ini.Text = Ini.LoadIni("Sets", "Splash")
     If .Ini.Text = "true" Then
      O.chkSplash.Value = 1
      CI.bSplash = True
     ElseIf .Ini.Text = "false" Then
      O.chkSplash.Value = 0
      CI.bSplash = False
     End If

     .Ini.Text = Ini.LoadIni("Sets", "Graph")
     If .Ini.Text = "true" Then
      O.chkGraph.Value = 1
      CI.bGraph = True
     ElseIf .Ini.Text = "false" Then
      O.chkGraph.Value = 0
      CI.bGraph = False
     End If

     .Ini.Text = Ini.LoadIni("Sets", "Snap")
     If .Ini.Text = "true" Then
      O.chkSnap.Value = 1
      CI.bSnap = True
     ElseIf .Ini.Text = "false" Then
      O.chkSnap.Value = 0
      CI.bSnap = False
     End If

     .Ini.Text = Ini.LoadIni("Sets", "Assoc")
     If .Ini.Text = "true" Then
      O.chkAss.Value = 1
      CI.bAss = True
     ElseIf .Ini.Text = "false" Then
      O.chkAss.Value = 0
      CI.bAss = False
     End If

     .Ini.Text = Ini.LoadIni("Sets", "Scroll")
     If .Ini.Text = "true" Then
      O.chkScroll.Value = 1
      CI.bScroll = True
     ElseIf .Ini.Text = "false" Then
      O.chkScroll.Value = 0
      CI.bScroll = False
     End If
    End With

End Sub
Public Sub SaveIniSettings(Sav As Boolean)
    
    On Error GoTo SError

    With CI
     If .bTop = True Then
      Call Ini.saveini("Sets", "Top", "true")
     ElseIf .bTop = False Then
      Call Ini.saveini("Sets", "Top", "false")
     End If

     If .bInst = True Then
      Call Ini.saveini("Sets", "Inst", "true")
     ElseIf .bInst = False Then
      Call Ini.saveini("Sets", "Inst", "false")
     End If

     If .bSnap = True Then
      Call Ini.saveini("Sets", "Snap", "true")
     ElseIf .bSnap = False Then
      Call Ini.saveini("Sets", "Snap", "false")
     End If

     If .bGraph = True Then
      Call Ini.saveini("Sets", "Graph", "true")
     ElseIf .bGraph = False Then
      Call Ini.saveini("Sets", "Graph", "false")
     End If

     If .bSort = True Then
      Call Ini.saveini("List", "Sort", "true")
     ElseIf .bSort = False Then
      Call Ini.saveini("List", "Sort", "false")
     End If

     If .bClick = True Then
      Call Ini.saveini("List", "Single", "true")
     ElseIf .bClick = False Then
      Call Ini.saveini("List", "Single", "false")
     End If

     If .bSplash = True Then
      Call Ini.saveini("Sets", "Splash", "true")
     ElseIf .bSplash = False Then
      Call Ini.saveini("Sets", "Splash", "false")
     End If

     If .bAss = True Then
      Call Ini.saveini("Sets", "Assoc", "true")
     ElseIf .bAss = False Then
      Call Ini.saveini("Sets", "Assoc", "false")
     End If

     If .BList = True Then
      Call Ini.saveini("List", "Show", "true")
     ElseIf .BList = False Then
      Call Ini.saveini("List", "Show", "false")
     End If

     If .bLoop = True Then
      Call Ini.saveini("List", "Loop", "true")
     ElseIf .bLoop = False Then
      Call Ini.saveini("List", "Loop", "false")
     End If

     If .bRand = True Then
      Call Ini.saveini("List", "Rand", "true")
     ElseIf .bRand = False Then
      Call Ini.saveini("List", "Rand", "false")
     End If

     If .bMute = True Then
      Call Ini.saveini("List", "Mute", "true")
     ElseIf .bMute = False Then
      Call Ini.saveini("List", "Mute", "false")
     End If

     If .bDoub = True Then
      Call Ini.saveini("List", "Dubs", "true")
     ElseIf .bDoub = False Then
      Call Ini.saveini("List", "Dubs", "false")
     End If

     If .bTray = True Then
      Call Ini.saveini("Sets", "Tray", "true")
     ElseIf .bTray = False Then
      Call Ini.saveini("Sets", "Tray", "false")
     End If

     If .bScroll = True Then
      Call Ini.saveini("Sets", "Scroll", "true")
     ElseIf .bScroll = False Then
      Call Ini.saveini("Sets", "Scroll", "false")
     End If

     Call Ini.saveini("Sets", "Vol", frmMn.picMVol.Left)
     Call Ini.saveini("Sets", "X", frmMn.Left)
     Call Ini.saveini("Sets", "Y", frmMn.Top)
     Call Ini.saveini("Sets", "Trans", CStr(.yLay))
     Call Ini.saveini("Sets", "Icon", CStr(.iIcon))

     .iSkin = frmSkn.lstSkins.ListIndex
     If Sav = True Then Call Ini.saveini("Sets", "Skin", CStr(.iSkin))
     Call Ini.saveini("Sets", "SPath", .sPath)
    End With

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub LoadIniSettings(Apply As Boolean, Frm As Long)

    On Error Resume Next
    Dim O As Object
    Set O = frmSet

    With frmMn
     .Ini.Text = Ini.LoadIni("Sets", "Top")
     If .Ini.Text = "true" Then
      O.chkTop.Value = 1
      CI.bTop = True
      Call DialogTop
     ElseIf .Ini.Text = "false" Then
      O.chkTop.Value = 0
      CI.bTop = False
      Call DialogBottom
     End If

     .Ini.Text = Ini.LoadIni("List", "Show")
     If Apply = True Then Call ListLeft
     If .Ini.Text = "true" Then
      CI.BList = True
      frmPl.Visible = True
     ElseIf .Ini.Text = "false" Then
      CI.BList = False
      frmPl.Visible = False
     End If

     .Ini.Text = Ini.LoadIni("List", "Loop")
     If .Ini.Text = "true" Then
      O.chkLoop.Value = 1
      CI.bLoop = True
     ElseIf .Ini.Text = "false" Then
      O.chkLoop.Value = 0
      CI.bLoop = False
     End If

     .Ini.Text = Ini.LoadIni("List", "Rand")
     If .Ini.Text = "true" Then
      O.chkRand.Value = 1
      CI.bRand = True
     ElseIf .Ini.Text = "false" Then
      O.chkRand.Value = 0
      CI.bRand = False
     End If

     .Ini.Text = Ini.LoadIni("List", "Mute")
     If .Ini.Text = "true" Then
      O.chkMute.Value = 1
      CI.bMute = True
      MPlay.Mute = True
      frmMnu.mnuMuteP.Checked = True
     ElseIf .Ini.Text = "false" Then
      O.chkMute.Value = 0
      CI.bMute = False
      MPlay.Mute = False
      frmMnu.mnuMuteP.Checked = False
     End If

     .Ini.Text = Ini.LoadIni("List", "Dubs")
     If .Ini.Text = "true" Then
      O.chkDubs.Value = 1
      CI.bDoub = True
     ElseIf .Ini.Text = "false" Then
      O.chkDubs.Value = 0
      CI.bDoub = False
     End If

     If Dir(File.SpecialFolder(7, Frm) & "CoolPlayer.lnk") <> "" Then
      O.chkStart.Value = 1
     Else
      O.chkStart.Value = 0
     End If
     
     If Apply = True Then
      Call MoveVolume(Ini.LoadIni("Sets", "Vol"))
      .Left = Ini.LoadIni("Sets", "X")
      .Top = Ini.LoadIni("Sets", "Y")
     End If
     Call LoadTransparency(Apply)
    End With

End Sub
 Public Sub GetAllColors()
    
    On Error GoTo AllError
    Dim O As Object
    Set O = frmPl.l

    With frmMn.Ini
     .Text = Ini.loadcolor("Text", "Normal", frmSkn.Files.Path)
     If Len(.Text) = 6 Then .Text = "#" & .Text
     If Len(.Text) = 7 Then
      O.ForeColor = Graph.GetRGB(.Text)
     ElseIf .Text = "#Error" Or .Text = "Error" Then
      O.ForeColor = &HFF00&
     End If

     .Text = Ini.loadcolor("Text", "mbBG", frmSkn.Files.Path)
     If Len(.Text) = 6 Then .Text = "#" & .Text
     If Len(.Text) = 7 Then
      O.BackColor = Graph.GetRGB(.Text)
     ElseIf .Text = "#Error" Or .Text = "Error" Then
      O.BackColor = &H0&
     End If
    End With

AllError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub SaveOptions()

    On Error GoTo SError
    With frmSet
     CI.bLoop = IIf(.chkLoop.Value = 1, True, False)
     CI.bRand = IIf(.chkRand.Value = 1, True, False)
     CI.bGraph = IIf(.chkGraph.Value = 1, True, False)
     CI.bClick = IIf(.chkSingl.Value = 1, True, False)
     CI.bSort = IIf(.chkSort.Value = 1, True, False)
     CI.bSplash = IIf(.chkSplash.Value = 1, True, False)
     CI.bSnap = IIf(.chkSnap.Value = 1, True, False)
     CI.bInst = IIf(.chkInst.Value = 1, True, False)
     CI.bAss = IIf(.chkAss.Value = 1, True, False)
     CI.bDoub = IIf(.chkDubs.Value = 1, True, False)
     CI.bTray = IIf(.chkMin.Value = 1, True, False)

     With frmMn
      If frmSet.chkMute.Value = 1 Then
       frmMnu.mnuMuteP.Checked = True
       CI.bMute = True
       MPlay.Mute = True
      ElseIf frmSet.chkMute.Value = 0 Then
       frmMnu.mnuMuteP.Checked = False
       CI.bMute = False
       MPlay.Mute = False
      End If
     End With

     If .chkStart.Value = 0 Then
      Call File.CheckCut(frmMn.hwnd)
     ElseIf .chkStart.Value = 1 Then
      Call File.CreateShortcut(App.Path, App.EXEName, frmMn.hwnd)
     End If

     If .chkTop.Value = 0 Then
      CI.bTop = False
      Call DialogBottom
     ElseIf .chkTop.Value = 1 Then
      CI.bTop = True
      Call DialogTop
     End If

     If .chkScroll.Value = 1 Then
      CI.bScroll = True
     ElseIf .chkScroll.Value = 0 Then
      CI.bScroll = False
     End If

     CI.yLay = CByte(.sliT.Value + 5)
     Call SaveIniSettings(False)
     Call SetLay(CI.yLay)
    End With

SError:
    If Err.Number <> 0 Then Call SaveIniSettings(False): Exit Sub

End Sub
Public Sub AddCommands(Comm As String)

    With frmPl
     Select Case Comm
      Case Is = ""
       Call LoadM3U(App.Path & Def)
       Call CheckSort(CI.bSort)

      Case Else
       Comm = Ini.GetLongPath(Comm)
       If Lst.getext(Comm) = "m3u" Or Lst.getext(Comm) = "pls" Then
        Call LoadList(Comm, True)
       Else
        Call AddFile(Comm, Right(Comm, Len(Comm) - InStrRev(Comm, "\")), True, True)
       End If
     End Select
    End With

End Sub
Public Sub CreateKey()

    On Error GoTo IError
    If Len(GetSetting("CoolPlayer", "JohnnyB", "Loaded")) = 0 Then
     Call CreateKeys
    End If
    Set Ini = CreateObject("Misc_v1.clsIni")
    Set Lst = CreateObject("Misc_v1.clsList")
    Set Tray = CreateObject("Misc_v1.clsTray")
    Set MPlay = CreateObject("Misc_v1.clsMCI")
    Set Volume = CreateObject("Misc_v1.clsVol")
    Set File = CreateObject("Misc_v1.clsFiles")
    Set ODialog = CreateObject("Misc_v1.clsCDex")
    Set IDx = CreateObject("Misc_v1.clsID3")
    Set Graph = CreateObject("Misc_v1.clsGraph")
    Set MP3 = CreateObject("Misc_v1.clsMP3info")
    Set Reg = CreateObject("Misc_v1.clsReg")
    Call LoadSkinNumber

IError:
    If Err.Number <> 0 Then
     Call DeleteKeys
     MsgBox ("Failed to load:  " & vbCrLf & App.Path & "\Plugins\Misc_v1.dll"), vbCritical, "Runtime Error 1": End
    End If

End Sub
