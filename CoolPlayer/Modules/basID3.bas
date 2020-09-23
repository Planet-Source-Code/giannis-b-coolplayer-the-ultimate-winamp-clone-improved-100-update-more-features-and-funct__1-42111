Attribute VB_Name = "basID3"
Option Explicit

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long

Public Sub CreateKeys()

    On Error Resume Next
    Call SaveSetting("CoolPlayer", "JohnnyB", "Loaded", "True")
    Call Register(frmMn.hwnd, App.Path & "\Plugins\Misc_v1.dll", True)

End Sub
Public Sub DeleteKeys()

    On Error Resume Next
    Call Register(frmMn.hwnd, App.Path & "\Plugins\Misc_v1.dll", False)
    If Len(GetSetting("CoolPlayer", "JohnnyB", "Loaded")) <> 0 Then
     Call DeleteSetting("CoolPlayer", "JohnnyB", "Loaded")
    End If

End Sub
Public Function GetGenre(Value As Integer) As Integer

    Dim i As Integer
    For i = 0 To 148
     If frmID3.lstGen.ItemData(i) = Value Then Exit For
    Next i
    GetGenre = i

End Function
Public Sub Register(Frm As Long, Path As String, Reg As Boolean)

    On Error Resume Next
    Dim LB As Long, PA As Long

    LB = LoadLibrary(Path)
    PA = IIf(Reg = True, GetProcAddress(LB, "DllRegisterServer"), GetProcAddress(LB, "DllUnregisterServer"))
    Call CallWindowProc(PA, Frm, ByVal 0&, ByVal 0&, ByVal 0&)
    Call FreeLibrary(LB)

End Sub
Public Sub Ster()

    Select Case MP3.mode
     Case Is = "Stereo"
      Call IsStereo

     Case Is = "Joint Stereo"
      Call IsStereo

     Case Is = "Dual Channel"
      Call IsMono

     Case Is = "Mono"
      Call IsMono
    End Select

End Sub
Public Sub CheckSort(Check As Boolean)

    If Check = True Then Call SortList(frmPl.l)
    Call GetPlay(True)

End Sub
Public Sub CloseEditor()

    Call DisableForms(True)
    GL.sTrack = "": Call IDx.ClearName
    Unload frmID3

End Sub
Public Sub GetID3(Name As String)

    With frmID3
     .Status.Panels(1).Text = IDx.ReadTag(Name)
     .Status.Panels(2).Text = File.MakeShort(Ini.GetShortPath(IDx.FileName))
     .txtTitle.Text = IDx.Title
     .txtArt.Text = IDx.Artist
     .txtAlb.Text = IDx.Album
     .txtYear.Text = IDx.Year
     .txtCom.Text = IDx.Comments
     .lstGen.ListIndex = GetGenre(CInt(IDx.Genre))
    End With

End Sub
Public Sub GetIt(Name As String)
    
    On Error GoTo LoadError
    Call IDx.ClearData
    Call MP3Info(Name)
    Call GetID3(Name)

LoadError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub MP3Info(Name As String)

    On Error GoTo MP3Error
    Dim O As Object
    Set O = frmID3

    With MP3
     .FileName = Name
     Call .GetHeader
     If Not .ValidHeader Then
      O.lblInfo = "No valid header. Could not retrieve MPEG information."
      O.lblInfoS.Caption = "": Exit Sub
     End If
    End With

    With MP3
     O.lblInfo.Caption = .id & vbCrLf & "Frequency: " & .Frequency _
     & " Hz" & vbCrLf & "Bitrate: " & .Bitrate & " Kbps" & vbCrLf & "Channel mode: " & .mode _
     & vbCrLf & .filesize & vbCrLf & .length & vbCrLf & .frames
    End With

    With MP3
     O.lblInfoS.Caption = .Padded & vbCrLf & .PrivateBit _
     & vbCrLf & .Copyrighted & vbCrLf & .Original _
     & vbCrLf & .Emphasis & vbCrLf & .ProtectionChecksum _
     & vbCrLf & .ModeExt
    End With

MP3Error:
    If Err.Number <> 0 Then Exit Sub

End Sub
