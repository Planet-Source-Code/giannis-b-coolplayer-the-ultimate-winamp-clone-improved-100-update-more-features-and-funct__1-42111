VERSION 5.00
Begin VB.UserControl ctlScroller 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   LockControls    =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   300
   Begin VB.PictureBox s 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   75
      ScaleHeight     =   270
      ScaleWidth      =   120
      TabIndex        =   0
      Top             =   0
      Width           =   120
   End
End
Attribute VB_Name = "ctlScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Scroll()

Private Type Scr
    Max As Long
    Val As Long
    Y As Single
    Down As Boolean
End Type

Private SR As Scr
Public Function MSlide(l As Long)

    Dim X As Long
    l = IIf(l <= 0, 1, l / 15)
    l = IIf(l >= 388, 388, l)
    X = (l * SR.Max) / 388
    SR.Val = IIf(X = 0, 1, X)
    s.Top = l * 15

End Function
Public Sub Update()

    On Error GoTo UError
    Dim i As Integer

    For i = 0 To 14
     Call UserControl.PaintPicture(frmMn.Pledit, 0, i * 435, 300, 435, 465, 630, 300, 435)
    Next i
    If SR.Down = False Then Call USlider

UError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub USlider()
    Call s.PaintPicture(frmMn.Pledit, 0, 0, 120, 270, 780, 795, 120, 270)
End Sub
Public Sub DSlider()
    Call s.PaintPicture(frmMn.Pledit, 0, 0, 120, 270, 915, 795, 120, 270)
End Sub
Private Sub s_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo DError
    If Button = 1 Then
     SR.Down = True: SR.Y = Y
     Call DSlider
    End If

DError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo MError
    If SR.Down = True Then
     If SR.Max <> 0 Then RaiseEvent Scroll
     Call DSlider
     Call MSlide(s.Top + Y - SR.Y)
    End If

MError:
    If Err.Number <> 0 Then Exit Sub
    
End Sub
Private Sub s_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 And SR.Down = True Then
     SR.Down = False: Call USlider
    End If

End Sub

Private Sub UserControl_Resize()
    
    On Error GoTo RError
    With UserControl
     .Width = 300: .Height = 6090
     Call Update
    End With

RError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Property Get Value() As Long
    Value = SR.Val
End Property
Public Property Let Value(l As Long)

    On Error GoTo LError
    SR.Val = l
    s.Top = (SR.Val * 15 / SR.Max) * 388
    Call USlider

LError:
    If Err.Number <> 0 Then Exit Property

End Property
Public Property Let Max(l As Long)

    On Error GoTo MError
    SR.Max = IIf(l < 0, 0, l)
    If SR.Max = 0 Then s.Top = 0
    Call Update

MError:
    If Err.Number <> 0 Then Exit Property

End Property
