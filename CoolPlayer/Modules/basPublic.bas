Attribute VB_Name = "basPub"
Option Explicit

Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Ini As Object
Public Tray As Object
Public MPlay As Object
Public Volume As Object
Public File As Object
Public ODialog As Object
Public Graph As Object
Public Lst As Object
Public MP3 As Object
Public Reg As Object
Public IDx As Object

Public Sub AddDirectory(Path As String)

    On Error GoTo AError
    Dim i As Long, F As Object

    Set F = frmMn.Files
    With frmPl.l.ListItems
     If F.ListCount = 0 Then
      Call DisableForms(True): Exit Sub
     Else
      For i = 0 To F.ListCount - 1
       Call .Add(.Count + 1, , Left(F.List(i), InStrRev(F.List(i), ".") - 1))
       .Item(.Count).SubItems(1) = MP3.gettime(Path & F.List(i), True)
       .Item(.Count).Tag = Path & F.List(i)
      Next i
     End If
     Call SetScroller(.Count)
    End With

AError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Function Allow(C As Control, Button As Integer, X As Single, Y As Single) As Boolean

    On Error Resume Next
    If Button = 1 And GL.bClick = True Then
     If X >= 0 And X <= C.Width And Y >= 0 And Y <= C.Height Then
      Allow = True
     Else
      Allow = False
     End If
    End If

End Function

Public Sub ChangeIcon()

    With frmSet
      frmMn.Icon = .Iml.ListImages(CI.iIcon).Picture
      Call Tray.changetrayicon(frmMn.Icon)
     .TrayI.Picture = frmMn.Icon
    End With

End Sub
Public Sub Mpop(Button As Integer, X As Single, Y As Single, Frm As Form, Mnu As Menu)

    With Frm
     If Button = 2 Then
      If X >= 0 And X <= .Width And Y >= 0 And Y <= .Height Then
       Call .PopupMenu(Mnu, True)
      End If
     End If
    End With

End Sub
Public Sub SetScroller(s As Long)

    With frmPl.l.ListItems
     If .Count = 0 Then Exit Sub
     .Item(s).EnsureVisible
     .Item(s).Selected = True: Call SetMax
    End With

End Sub

Public Sub ToTime(Dur As Long)
    
    On Error GoTo TError
    With frmTim.txtTime
     If File.ConvertMinSec(.Text) > Dur Then .Text = "00:00": Exit Sub
     If MPlay.FileName <> "" Or File.ConvertMinSec(.Text) < Dur _
     Or .Text <> "00:00" Or Left(.Text, 3) = ":" Then
      Call File.ConvertMinSec(.Text)
      MPlay.CurrentPosition = File.ConvertMinSec(.Text)
     Else
      .Text = "00:00"
     End If

TError:
     If Err <> 0 Then .Text = "00:00": Exit Sub
    End With

End Sub
Public Sub ReverseList(l As MSComctlLib.ListView)

    l.SortOrder = lvwDescending
    l.Sorted = True: l.Sorted = False
    Call RemDoubs(frmPl.l, CI.bDoub)

End Sub
Public Sub SortList(l As MSComctlLib.ListView)

    l.SortOrder = lvwAscending
    l.Sorted = True: l.Sorted = False
    Call RemDoubs(frmPl.l, CI.bDoub)

End Sub
Public Sub Play()

    On Error GoTo PError
    Dim i As Long
    GL.lPos = frmPl.l.SelectedItem.Index

    With MPlay
     If GL.sTrack <> "" And Dir(GL.sTrack) <> "" Then
      If .Playstate <> 2 Then .FileName = GL.sTrack
      Call .Play
      Call MP3PlayInfo(GL.sTrack)
      Call SHPics(True)
      Call ButtonChoice("played")
     End If
    End With

PError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub StopPlay()

    On Error GoTo SError
    With frmMn
     .picMSli.Left = 240: MPlay.StopT
     Call SHPics(False)
    End With

    Call ButtonChoice("stopped")
    Call SliderUp: GL.sTrack = ""

SError:
     If Err.Number <> 0 Then Exit Sub

End Sub
Public Function TimePosition(Time As String, Optional DoIt As Boolean) As String

    On Error GoTo TError
    Dim Min As String, Sec As String

    Min = Format(Int(Time) \ 60, "00")
    If Len(Min) = 1 Then Min = "0" & Min
    Sec = Int(Time) Mod 60
    If Sec < 0 Then Sec = 0
    TimePosition = Min & ":" & Format(Sec, "00")
    If DoIt = True Then Call SetNumbers(CInt(Left(Min, 1)), CInt(Right(Min, 1)), CInt(Left(Format(Sec, "00"), 1)), CInt(Right(Format(Sec, "00"), 1)))

TError:
    If Err.Number <> 0 Then Exit Function

End Function
Public Sub ProgramExit()

    On Error GoTo EError
    Call HideForms(True)
    Call DisableForms(False)
    DoEvents
    Call Graph.GraphExit(CI.bGraph, CI.yLay, frmMn.hwnd, frmPl.hwnd)
    Call SaveIniSettings(False)
    Call Tray.RemoveTrayIcon
    End

EError:
    If Err.Number <> 0 Then End

End Sub
