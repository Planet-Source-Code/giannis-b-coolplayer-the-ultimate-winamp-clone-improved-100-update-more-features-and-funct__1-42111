VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMCI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public F As New clsIni
Public V As New clsVol

Public Enum State
    mpPlaying = 1
    mpPaused = 2
    mpStopped = 3
End Enum

Private Type MCIParam
    bMut As Boolean
    iDur As Long
    iPos As Long
    iVol As Long
    iBal As Long
    sName As String
    sState As State
End Type

Private MC As MCIParam
Private ResM As String

Private Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Property Let Filename(ByVal Name As String)

    On Error GoTo LError
    ResM = Space(128)
    Name = F.GetShortPath(Name)
    Call mciSendString("close all", ResM, 128, 0)
    Call mciSendString("open " & Chr(34) & Name & Chr(34) & " alias media", ResM, 128, 0)
    Call mciSendString("set media time format milliseconds", ResM, 128, 0)
    MC.sName = Name

LError:
    If Err.Number <> 0 Then MC.sName = "": Call StopT: Exit Property

End Property
Public Property Get Filename() As String

    On Error Resume Next
    Filename = F.GetLongPath(MC.sName)

End Property
Public Property Get Playstate() As String

    On Error Resume Next
    Playstate = MC.sState

End Property
Public Property Get EndOfStream() As Boolean

    On Error GoTo GError
    If MC.sState <> mpPlaying Then
     EndOfStream = False: Exit Property
    End If
    EndOfStream = IIf(MC.iDur - MC.iPos = 0, True, False)

GError:
    If Err.Number <> 0 Then Exit Property

End Property
Public Function Play() As Boolean

    On Error GoTo PError
    ResM = Space(128)
    Call mciSendString("play media", ResM, 128, 0)
    MC.sState = mpPlaying

PError:
    If Err.Number <> 0 Then Exit Function

End Function
Public Sub StopT()

    On Error GoTo SError
    ResM = Space(128)
    Call mciSendString("close media", ResM, 128, 0)
    Call mciSendString("close all", ResM, 128, 0)
    MC.sState = mpStopped: MC.sName = ""

SError:
    If Err.Number <> 0 Then MC.sState = mpStopped: MC.sName = "": Exit Sub

End Sub
Public Property Get Duration() As Long

    On Error GoTo GError
    If MC.sState = mpStopped Then Exit Property

    ResM = Space(128)
    Call mciSendString("status media length", ResM, 128, 0)
    MC.iDur = CInt(ResM / 1000)
    Duration = MC.iDur

GError:
    If Err.Number <> 0 Then Duration = MC.iDur: Exit Property

End Property
Public Property Get CurrentPosition() As Long

    On Error GoTo GError
    If MC.sState = mpStopped Then Exit Property

    ResM = Space(128)
    Call mciSendString("status media position", ResM, 128, 0)
    MC.iPos = CInt(ResM / 1000)
    CurrentPosition = MC.iPos

GError:
    If Err.Number <> 0 Then CurrentPosition = MC.iPos: Exit Sub

End Sub
Public Property Get Volume() As Long

    On Error Resume Next
    MC.iVol = V.GetVolume(WAVEOUT)
    Volume = MC.iVol

End Property
Public Property Get Mute() As Boolean

    On Error Resume Next
    MC.bMut = V.GetMute(mWAVEOUT)
    Mute = MC.bMut

End Property
Public Property Get Balance() As Long

    On Error Resume Next
    MC.iBal = V.GetBalance(WAVEOUT)
    Balance = MC.iBal

End Property
Public Property Let CurrentPosition(ByVal Value As Long)

    On Error GoTo LError
    If MC.sState = mpStopped Then Exit Property

    ResM = Space(128)
    If Value < 0 Then Value = 0
    Call mciSendString("seek media to " & CLng(Value * 1000), ResM, 128, 0)
    MC.iPos = CLng(Value * 1000)
    If MC.sState = mpPlaying Then Call Play

LError:
    If Err.Number <> 0 Then MC.iPos = 0: Exit Property

End Property
Public Property Let Volume(ByVal Value As Long)

    On Error GoTo LError
    If MC.bMut = False Then
     MC.iVol = CLng(Value)
     Call V.SetVolume(MC.iBal, MC.iVol, WAVEOUT)
    End If

LError:
    If Err.Number <> 0 Then Exit Property

End Property
Public Property Let Mute(ByVal Value As Boolean)

    On Error GoTo LError
    MC.bMut = Value
    Call V.SetMute(Value, mWAVEOUT)

LError:
    If Err.Number <> 0 Then Exit Property

End Property
Public Property Let Balance(ByVal Value As Long)

    On Error GoTo LError
    MC.iBal = CInt(Value)
    If MC.bMut = False Then Call V.SetBalance(MC.iBal, 100 - MC.iVol, WAVEOUT)

LError:
    If Err.Number <> 0 Then MC.iBal = 0: Exit Property

End Property
Public Sub Pause()

    On Error GoTo PError
    If MC.sState = mpStopped Then Exit Sub

    ResM = Space(128)
    Call mciSendString("pause media", ResM, 128, 0)
    MC.sState = mpPaused

PError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub ResumeT()

    On Error GoTo PError
    If MC.sState = mpStopped Then Exit Sub

    ResM = Space(128)
    Call mciSendString("resume media", ResM, 128, 0)
    MC.sState = mpPlaying

PError:
    If Err.Number <> 0 Then MC.sState = mpPaused: Exit Sub

End Sub

Private Sub Class_Terminate()
    Call StopT
End Sub
