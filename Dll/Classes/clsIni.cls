VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIni"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function WritePrivateProfileString _
               Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal lpApplicationname As String, ByVal _
               lpKeyName As Any, ByVal lsString As Any, _
               ByVal lplFilename As String) As Long

Private Declare Function GetPrivateProfileString Lib _
               "kernel32" Alias "GetPrivateProfileStringA" _
               (ByVal lpApplicationname As String, ByVal _
               lpKeyName As String, ByVal lpDefault As _
               String, ByVal lpReturnedString As String, _
               ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Function GetWinDir() As String

    On Error GoTo DirError
    Dim Win As String * 255
    GetWinDir = Left(Win, GetWindowsDirectory(Win, 255))

DirError:
    If Err.Number <> 0 Then Exit Function

End Function
Public Function LoadIni(Sec As String, Key As String) As String

    On Error GoTo LError
    Dim Name As String, strResult As String
    
    If Dir(GetWinDir & "\CPlay.ini") <> "" Then
     Name = GetWinDir & "\CPlay.ini"
    Else
     LoadIni = Trim("Error"): Exit Function
    End If

    strResult = String(1024, Chr(0))
    LoadIni = Left(strResult, GetPrivateProfileString(Sec, ByVal Key, "Error", strResult, Len(strResult), Name))

LError:
    If Err.Number <> 0 Then Exit Function

End Function
Public Sub SaveIni(Sec As String, Key As String, Val As String)

    On Error GoTo SError
    Dim Name As String

    Name = GetWinDir & "\CPlay.ini"
    Call WritePrivateProfileString(Sec, Key, Val, Name)

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Function LoadColor(Sec As String, Key As String, Path As String) As String

    On Error GoTo LError
    Dim Name As String, strResult As String

    If Dir(Path & "\Pledit.txt") <> "" Then
     Name = Path & "\Pledit.txt"
    Else
     LoadColor = Trim("Error"): Exit Function
    End If
    
    strResult = String(1024, Chr(0))
    LoadColor = Left(strResult, GetPrivateProfileString(Sec, ByVal Key, "Error", strResult, Len(strResult), Name))

LError:
    If Err.Number <> 0 Then Exit Function

End Function
Public Function GetLongPath(ByVal Name As String) As String

    On Error GoTo GetError
    Dim lProp As Long

    GetLongPath = Space(1024)
    lProp = GetLongPathName(Name, GetLongPath, Len(GetLongPath))
    GetLongPath = Left(GetLongPath, lProp)

GetError:
    If Err.Number <> 0 Then GetLongPath = Name

End Function
Public Function GetShortPath(ByVal Name As String) As String

    On Error GoTo GError
    Dim RetVal As Long

    GetShortPath = Space(1024)
    RetVal = GetShortPathName(Name, GetShortPath, Len(GetShortPath))
    GetShortPath = LCase(Left(GetShortPath, RetVal))

GError:
    If Err.Number <> 0 Then GetShortPath = LCase(Name)

End Function
