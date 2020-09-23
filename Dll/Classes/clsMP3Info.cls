VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMP3Info"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type MPEG
     ModeExt As Integer
     Bitrate As Integer
     Frames As Integer

     Frequency As Long
     FileLength As Long
     Len As Long

     ValidHeader As Boolean
     ProtectionBitSet As Boolean
     Padded As Boolean
     PrivateBitSet As Boolean
     Copyrighted As Boolean
     Original As Boolean

     Emphasis As String
     Mode As String
     Filename As String
     ID As String
     Layer As String
     ProtectionChecksum As String
End Type

Private MP3 As MPEG
Private Fi As New clsFiles
Public Property Let Filename(sInput As String)

    Call ClearData
    MP3.Filename = sInput

End Property
Public Property Get Filename() As String
    Filename = MP3.Filename
End Property
Public Function GetTime(Name As String, C As Boolean) As String

    On Error GoTo GError
    Dim bTMP1 As Byte, bTMP2 As Byte
    Dim FileNum As Integer, S As Long

    With MP3
     FileNum = FreeFile
     .FileLength = FileLen(Name)

     Dim I As Long, SByte As Long
     Close #FileNum
     Open Name For Binary As #FileNum
      For I = 1 To LOF(FileNum) - 1
       Get #FileNum, I, bTMP1
       If bTMP1 = &HFF Then
        Get #FileNum, I + 1, bTMP2
        If bTMP2 And &HE0 = &HE0 Then
         .ValidHeader = True
         SByte = I + 1: Exit For
        End If
       End If
      Next

      If Not .ValidHeader Then
       GetTime = "0"
       Close #1: Exit Function
      End If
    
      Dim B1 As Byte, B2 As Byte
      Dim B3 As Byte, B4 As Byte, B5 As Byte
      Get #FileNum, SByte, B1
      Get #FileNum, SByte + 1, B2
      Get #FileNum, SByte + 2, B3
      Get #FileNum, SByte + 3, B4
      Get #FileNum, SByte + 4, B5
     Close #FileNum

     Dim ARG1 As Integer, ARG2 As Integer, ARG3 As Integer
     ARG1 = CInt(B1 And &H8) / 8
     ARG2 = CInt(B1 And &H6) / 2
     ARG3 = CInt(B2 And &HF0) / 16

     S = Int((Int(.FileLength * 0.8)) / (Int(GetBitrate(ARG1, ARG2, ARG3) * 100)))
     GetTime = IIf(C = True, Fi.GetTime(CStr(S)), CStr(S))
    End With

GError:
    If Err.Number <> 0 Then
     GetTime = "0": Close #FileNum
     MP3.ValidHeader = False
    End If

End Function

Public Property Get ValidHeader() As Boolean
    ValidHeader = MP3.ValidHeader
End Property
Public Property Get ID() As String
    ID = MP3.ID & " " & MP3.Layer
End Property
Public Property Get ProtectionBitSet() As Boolean
    ProtectionBitSet = MP3.ProtectionBitSet
End Property
Public Property Get ProtectionChecksum() As String
    ProtectionChecksum = "Protection: " & MP3.ProtectionChecksum
End Property
Public Property Get Bitrate() As String
    Bitrate = MP3.Bitrate
End Property
Public Property Get Length() As String
    Length = "Length: " & Int((Int(MP3.FileLength * 0.8)) / (Int(MP3.Bitrate * 100))) & " Seconds "
End Property
Public Property Get Frames() As String
    Frames = "Frames: " & MP3.Frames
End Property
Public Property Get Frequency() As String
    Frequency = MP3.Frequency
End Property
Public Property Get Padded() As String
    Padded = "Padded: " & MP3.Padded
End Property
Public Property Get PrivateBitSet() As Boolean
    PrivateBitSet = MP3.PrivateBitSet
End Property
Public Property Get PrivateBit() As String
    PrivateBit = "Private: " & MP3.PrivateBitSet
End Property
Public Property Get Mode() As String
    Mode = MP3.Mode
End Property
Public Property Get ModeExt() As String
    ModeExt = "Mode ext: " & MP3.ModeExt
End Property
Public Property Get Copyrighted() As String
    Copyrighted = "Copyrighted: " & MP3.Copyrighted
End Property
Public Property Get Original() As String
    Original = "Original: " & MP3.Original
End Property
Public Property Get Emphasis() As String
    Emphasis = MP3.Emphasis
End Property
Public Property Get Filesize() As String
    Filesize = "File size: " & MP3.FileLength & " bytes"
End Property
Public Sub GetHeader()

    On Error GoTo GError
    Dim bTMP1 As Byte, bTMP2 As Byte
    Dim FileNum As Integer

    With MP3
     FileNum = FreeFile
     .FileLength = FileLen(.Filename)

     Dim I As Long, SByte As Long
     Close #FileNum
     Open .Filename For Binary As #FileNum
      For I = 1 To LOF(FileNum) - 1
       Get #FileNum, I, bTMP1
       If bTMP1 = &HFF Then
        Get #FileNum, I + 1, bTMP2
        If bTMP2 And &HE0 = &HE0 Then
         .ValidHeader = True
         SByte = I + 1: Exit For
        End If
       End If
      Next
      If Not .ValidHeader Then .ValidHeader = False: Exit Sub

      Dim B1 As Byte, B2 As Byte
      Dim B3 As Byte, B4 As Byte, B5 As Byte
      Get #FileNum, SByte, B1
      Get #FileNum, SByte + 1, B2
      Get #FileNum, SByte + 2, B3
      Get #FileNum, SByte + 3, B4
      Get #FileNum, SByte + 4, B5
     Close #FileNum

     Select Case CInt(B1 And &H18) / 8
      Case 0: .ID = "MPEG 2.5"
      Case 1: .ID = "MPEG Not Defined"
      Case 2: .ID = "MPEG 2"
      Case 3: .ID = "MPEG 1"
     End Select

     Select Case (B1 And &H6)
      Case &H0: .Layer = "Layer Not Defined"
      Case &H2: .Layer = "Layer III"
      Case &H4: .Layer = "Layer II"
      Case &H6: .Layer = "Layer I"
     End Select

     If (B1 And &H1) = &H1 Then
      .ProtectionBitSet = True
     Else
      .ProtectionBitSet = False
      .ProtectionChecksum = Hex(B4) & " " & Hex(B5)
     End If

     Dim ARG1 As Integer, ARG2 As Integer, ARG3 As Integer
     ARG1 = CInt(B1 And &H8) / 8
     ARG2 = CInt(B1 And &H6) / 2
     ARG3 = CInt(B2 And &HF0) / 16
     .Bitrate = GetBitrate(ARG1, ARG2, ARG3)

     Dim K As Long
     Select Case CInt(B1 And &H18) / 8
      Case 0: K = 1
      Case 1: K = 0
      Case 2: K = 2
      Case 3: K = 4
     End Select
    
     Select Case CInt(B2 And &HC) / 4
      Case 0: .Frequency = K * 11025
      Case 1: .Frequency = K * 12000
      Case 2: .Frequency = K * 8000
     End Select

     .Frames = CInt(.FileLength / ((.Bitrate * 144000) / (.Frequency)))
     If (B2 And &H2) = &H2 Then .Padded = True
     If (B2 And &H1) = &H1 Then .PrivateBitSet = True

     Select Case CInt(B3 And &HC0) / CInt(&H40)
      Case 0: .Mode = "Stereo"
      Case 1: .Mode = "Joint Stereo"
      Case 2: .Mode = "Dual Channel"
      Case 3: .Mode = "Mono"
     End Select

     .ModeExt = CInt(B3 And &H30) / CInt(&H10)
     If (B3 And &H8) = &H8 Then .Copyrighted = True
     If (B3 And &H4) = &H4 Then .Original = True

     Select Case B3 And &H3
      Case 0: .Emphasis = "Empasis: None"
      Case 2: .Emphasis = "Emphasis: Not defined"
      Case 1: .Emphasis = "Emphasis: 50/15 ms"
      Case 3: .Emphasis = "Emphasis: CCITT J. 17"
     End Select
    End With

GError:
    If Err.Number <> 0 Then: MP3.ValidHeader = False: Close #FileNum

End Sub
Private Function GetBitrate(ARG1 As Integer, ARG2 As Integer, ARG3 As Integer) As Integer

    On Error GoTo BitError
    Dim I As Integer
    Dim A(1, 3, 15) As Integer
    
    For I = 1 To 14
     A(1, 3, I) = I * 32
     If I < 5 Then
      A(1, 2, I) = 8 * (I + 4)
      A(1, 1, I) = 8 * (I + 3)
     Else
      A(1, 2, I) = A(1, 2, I - 4) * 2
      A(1, 1, I) = A(1, 1, I - 4) * 2
     End If
     If I < 9 Then
      A(0, 1, I) = I * 8
     Else
      A(0, 1, I) = (I - 4) * 16
     End If
      A(0, 2, I) = A(0, 1, I)
    Next
    
    A(1, 2, 1) = 32: A(0, 3, 1) = 32: A(0, 3, 2) = 48
    A(0, 3, 3) = 56: A(0, 3, 4) = 64: A(0, 3, 5) = 80
    A(0, 3, 6) = 96: A(0, 3, 7) = 112: A(0, 3, 8) = 128
    A(0, 3, 9) = 144: A(0, 3, 10) = 160: A(0, 3, 11) = 176
    A(0, 3, 12) = 192: A(0, 3, 13) = 224: A(0, 3, 14) = 256

    GetBitrate = A(ARG1, ARG2, ARG3)
    If ARG3 = 15 Then GetBitrate = 1
    If ARG3 = 0 Then GetBitrate = 0

BitError:
    If Err.Number <> 0 Then GetBitrate = 0: Exit Function

End Function
Private Sub ClearData()

    With MP3
     .Bitrate = 0: .Copyrighted = False
     .Emphasis = "": .FileLength = 0
     .Filename = "": .Frames = 0
     .Frequency = 0: .ID = ""
     .Layer = "": .Len = 0
     .Mode = "": .ModeExt = 0
     .Original = False: .Padded = False
     .PrivateBitSet = False: .ProtectionBitSet = False
     .ProtectionChecksum = "": .ValidHeader = False
    End With

End Sub
