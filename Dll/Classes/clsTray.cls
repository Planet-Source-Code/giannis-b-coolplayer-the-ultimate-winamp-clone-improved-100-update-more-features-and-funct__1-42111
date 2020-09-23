VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTray"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private NID As NOTIFYICONDATA

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_MOUSEMOVE = &H200

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Sub AddTrayIcon(Icon As Long, Frm As Long, Optional ToolTip As String)

   On Error GoTo AError
   With NID
    .cbSize = Len(NID)
    .hWnd = Frm
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Icon
    .szTip = ToolTip & vbNullChar
   End With
   Call Shell_NotifyIcon(NIM_ADD, NID)

AError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub RemoveTrayIcon()
   
   On Error GoTo TrayRemError
   Call Shell_NotifyIcon(NIM_DELETE, NID)

TrayRemError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Function TrayEvent(X As Single) As Long

    On Error GoTo PError
    Dim Msg As Long

    Msg = (X / Screen.TwipsPerPixelX)
    Select Case Msg
     Case WM_LBUTTONUP
      TrayEvent = WM_LBUTTONUP
     Case WM_RBUTTONUP
      TrayEvent = WM_RBUTTONUP
    End Select

PError:
    If Err.Number <> 0 Then Exit Function

End Function
Public Sub ToolTip(Mes As String)

   On Error GoTo TError
   NID.szTip = Mes & vbNullChar
   Call Shell_NotifyIcon(NIM_MODIFY, NID)

TError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub ChangeTrayIcon(Icon As Long)

   On Error GoTo CError
   NID.hIcon = Icon
   Call Shell_NotifyIcon(NIM_MODIFY, NID)

CError:
    If Err.Number <> 0 Then Exit Sub

End Sub
