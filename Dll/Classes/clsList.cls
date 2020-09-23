VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsList"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private List As Object
Private MP As New clsMP3Info
Public Sub FindString(StrR As String, L As Object)

    On Error GoTo FError
    Dim Found As Boolean
    Static i As Long

    Set List = L
    With List.listitems
     If StrR <> "" Then
      While i <= .Count - 1 And Not Found = True
       Found = MiddleFound(StrR, .item(i + 1))
       If Found = True Then
        .item(i + 1).Selected = True: .item(i + 1).ensurevisible
       End If
       i = i + 1
      Wend
      If i >= .Count Then i = 1
     End If
    End With

FError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Function MiddleFound(This, InThis) As Boolean
    If InStr(1, LCase(InThis), LCase(This)) > 0 Then MiddleFound = True
End Function
Public Sub SaveM3U(Name As String, L As Object)

    On Error GoTo SError
    Dim i As Long

    Set List = L
    Close #1
    With List.listitems
     Open Name For Output As #1
      Print #1, "#EXTM3U"
      For i = 1 To .Count
       Print #1, "#EXTINF:" & MP.GetTime(.item(i).TAG, False) & "," & .item(i)
       Print #1, .item(i).TAG
      Next i
    End With
    Close #1

SError:
    If Err.Number <> 0 Then Close #1:  Exit Sub

End Sub
Public Sub SavePLS(Name As String, L As Object)

    On Error GoTo SError
    Dim i As Long

    Set List = L
    Close #1
    With List.listitems
     Open Name For Output As #1
      Print #1, "[playlist]"
      For i = 1 To .Count
       Print #1, "File" & i & "=" & .item(i).TAG
       Print #1, "Title" & i & "=" & .item(i)
       Print #1, "Length" & i & "=" & MP.GetTime(.item(i).TAG, False)
      Next i
      Print #1, "NumberOfEntries=" & i
      Print #1, "Version=2"
    End With
    Close #1

SError:
    If Err.Number <> 0 Then Close #1: Exit Sub

End Sub

Public Sub SaveList(Name As String, Index As Integer, L As Object)

    On Error GoTo SError
    If Name = "" Then Exit Sub

    Select Case Index
     Case Is = 1
      If GetExt(Name) = "m3u" Then
       Call SaveM3U(Name, L)
      Else
       Call SaveM3U(Name & ".m3u", L)
      End If

     Case Is = 2
      If GetExt(Name) = "pls" Then
       Call SavePLS(Name, L)
      Else
       Call SavePLS(Name & ".pls", L)
      End If
    End Select

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Function Random(Min As Long, Max As Long, sVal As Long) As Long

    Static Seed As Boolean
    If Not Seed Then Randomize: Seed = True
    If Max > 0 Then
     If Max > 1 Then
      Do
       Random = CLng((Max - Min) * Rnd + Min)
      Loop Until Random <> sVal And Random >= Min
     Else
      Random = 1
     End If
    End If

End Function
Public Sub RandomizeList(L As Object)

    Dim i As Long, T(5) As String, u As Long

    Set List = L
    With List.listitems
     For i = 1 To .Count - 1
      u = Random(1, .Count, i)
      T(0) = .item(i).Text: T(1) = .item(i).SubItems(1): T(2) = .item(i).TAG
      T(3) = .item(u).Text: T(4) = .item(u).SubItems(1): T(5) = .item(u).TAG
      Call .Remove(i): Call .Add(i, , T(3)): .item(i).SubItems(1) = T(4): .item(i).TAG = T(5)
      Call .Remove(u): Call .Add(u, , T(0)): .item(u).SubItems(1) = T(1): .item(u).TAG = T(2)
     Next i
    End With

End Sub
Private Sub Class_Initialize()

    On Error Resume Next
    Set List = CreateObject("MSComctlLib.ListItem")

End Sub
Public Function GetExt(Ext As String) As String
    GetExt = LCase(Right(Ext, Len(Ext) - InStrRev(Ext, ".")))
End Function
