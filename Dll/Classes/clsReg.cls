VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReg"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1


Public Sub PublicReg(Index As Integer, Path As String, Ex As String)

    Select Case Index
     Case Is = 1: Call AssociateFile(".mp3", "CoolPlayer media file", Path, Ex)
     Case Is = 2: Call AssociateFile(".wav", "CoolPlayer media file", Path, Ex)
     Case Is = 3: Call AssociateFile(".wma", "CoolPlayer media file", Path, Ex)
     Case Is = 4: Call AssociateFile(".mid", "CoolPlayer media file", Path, Ex)
     Case Is = 5: Call AssociateFile(".midi", "CoolPlayer media file", Path, Ex)
     Case Is = 6: Call AssociateFile(".m3u", "CoolPlayer playlist file", Path, Ex)
     Case Is = 7: Call AssociateFile(".pls", "CoolPlayer playlist file", Path, Ex)
     Case Is = 8: Call AssociateFile(".lst", "CoolPlayer playlist file", Path, Ex)
    End Select

End Sub
Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)

    On Error Resume Next
    Dim hCurKey As Long, lRes As Long

    lRes = RegCreateKey(hKey, strPath, hCurKey)
    lRes = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

    If lRes <> ERROR_SUCCESS Then Exit Sub
    Call RegCloseKey(hCurKey)

End Sub
Public Sub AssociateFile(Ext As String, Promt As String, Path As String, ExeName As String)

    On Error GoTo AssocError
    Dim EXpath As String

    EXpath = IIf(Right(Path, 1) = "\", Path + ExeName + ".exe %1", Path + "\" + ExeName + ".exe %1")
    Call SaveSettingString(HKEY_CLASSES_ROOT, Ext, "", ExeName + Ext)
    Call SaveSettingString(HKEY_CLASSES_ROOT, Ext & "\ShellNew", "", "")
    Call SaveSettingString(HKEY_CLASSES_ROOT, ".lst\ShellNew", "NullFile", "")
    Call SaveSettingString(HKEY_CLASSES_ROOT, ExeName + Ext, "", Promt)
    Call SaveSettingString(HKEY_CLASSES_ROOT, ExeName + Ext + "\Shell\Open\command", "", EXpath)
    Call SaveSettingString(HKEY_CLASSES_ROOT, ExeName + Ext + "\DefaultIcon", "", Path & "\" & ExeName & ".exe,0")
    Call SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0)

AssocError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub RunStartup(Title As String, Path As String, Run As Boolean, RunOnce As Boolean)

    Dim sRun As String
    sRun = IIf(Run, "Run", "RunOnce")
    Call SaveSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\" & sRun, Title, Path + "\" + Title + ".exe")

End Sub

