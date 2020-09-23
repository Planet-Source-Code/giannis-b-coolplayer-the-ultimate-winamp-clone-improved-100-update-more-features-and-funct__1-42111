VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraph"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Public Sub GraphExit(Start As Boolean, From As Byte, Frm As Long, frm1 As Long)

    On Error Resume Next
    Dim bytI As Byte, Z As Integer

    If Start = False Then Exit Sub
    bytI = From
    For Z = 1 To From Step 2
     bytI = bytI - 1
     Call SetLayered(Frm, bytI)
     Call SetLayered(frm1, bytI)
    Next

End Sub
Public Sub OnTop(Frm As Long, Top As Boolean)
    
    On Error GoTo BError
    Dim i As Long

    i = IIf(Top = True, -1, -2)
    Call SetWindowPos(Frm, i, 0, 0, 0, 0, Flags)

BError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub LoadElliptic(Frm As Object)

    On Error GoTo LError
    Dim R As Long

    R = CreateRoundRectRgn(0, 0, Frm.Width / Screen.TwipsPerPixelX, Frm.Height / Screen.TwipsPerPixelY, 25, 25)
    'R = CreateEllipticRgn(0, 0, Frm.Width / Screen.TwipsPerPixelX, Frm.Height / Screen.TwipsPerPixelY)
    Call SetWindowRgn(Frm.hWnd, R, False)

LError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Function GetRGB(Value As String) As Currency

    On Error GoTo HexError
    Dim RGBs() As String

    If Len(Value) <> 7 Then Exit Function
    RGBs() = Split(HexToRGB(Right(Value, 6)), ",")
    GetRGB = RGB(RGBs(0), RGBs(1), RGBs(2))

HexError:
    If Err.Number <> 0 Then Exit Function

End Function
Public Sub SetLayered(ByVal hWnd As Long, B As Byte)

    Call SetWindowLong(hWnd, -20, GetWindowLong(hWnd, -20) Or &H80000)
    Call SetLayeredWindowAttributes(hWnd, 0, B, &H2)

End Sub
Private Function HexToRGB(ByVal Value As String, Optional Del As String = ",") As String

    Dim RGBValue(2) As String
    RGBValue(0) = ((GetHexValue(Mid(Value, 1, 1))) _
                  * 16 + (GetHexValue(Mid(Value, 2, 1))))
        
    RGBValue(1) = ((GetHexValue(Mid(Value, 3, 1))) _
                  * 16 + (GetHexValue(Mid(Value, 4, 1))))
        
    RGBValue(2) = ((GetHexValue(Mid(Value, 5, 1))) _
                  * 16 + (GetHexValue(Mid(Value, 6, 1))))
        
    HexToRGB = (Join(RGBValue(), Del))

End Function
Private Function GetHexValue(HChar As String) As String
    
    Select Case HChar
     Case "A", "a"
      GetHexValue = 10
     Case "B", "b"
      GetHexValue = 11
     Case "C", "c"
      GetHexValue = 12
     Case "D", "d"
      GetHexValue = 13
     Case "E", "e"
      GetHexValue = 14
     Case "F", "f"
      GetHexValue = 15
     Case 0 To 9
      GetHexValue = HChar
     Case Else
      GetHexValue = 0
    End Select

End Function
