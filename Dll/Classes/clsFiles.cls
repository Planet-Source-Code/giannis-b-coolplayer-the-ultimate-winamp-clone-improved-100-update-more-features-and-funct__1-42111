VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFiles"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private G As New clsGraph

Private Type SHITEMID
    CD As Long
    ABID As Byte
End Type

Private Type ITEMIDLIST
    Mkid As SHITEMID
End Type

Private Type BROWSEINFO
    hWnd As Long
    Root As Long
    DisplayName As String
    Title As String
    Flags As Long
    PFN As Long
    Param As Long
    Image As Long
End Type

Private R As Long

Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal Pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, Pidl As ITEMIDLIST) As Long

Private Const WM_USER = &H400
Private Const LPTR = (&H0 Or &H40)
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)

Private BI As BROWSEINFO
Private IDL As ITEMIDLIST

Public Sub AboutPlugin(Top As Boolean)

    On Error GoTo AError
    With frmAb
     Load frmAb
     Call G.OnTop(.hWnd, Top)
     .Caption = "About Misc_v1 v" & App.Major & _
                "." & App.Minor & "." & App.Revision
     .lblDate.Caption = "Date: " & Date
     .Show vbModal
    End With

AError:
    If Err.Number <> 0 Then Unload frmAb

End Sub
Public Sub CheckCut(Frm As Long)
    
    On Error GoTo IError
    If Dir(SpecialFolder(7, Frm) & "CoolPlayer.lnk") <> "" Then
     Call Kill(SpecialFolder(7, Frm) & "CoolPlayer.lnk")
    End If

IError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Function SpecialFolder(CSIDL As Long, Frm As Long) As String

    Dim sPath As String
    sPath = Space(260)

    R = SHGetSpecialFolderLocation(Frm, CSIDL, IDL)
    If R = 0 Then
     R = SHGetPathFromIDList(ByVal IDL.Mkid.CD, ByVal sPath)
     If R Then SpecialFolder = ValidateDir(Left(sPath, InStr(sPath, Chr(0)) - 1))
    End If

End Function
Public Sub CreateShortcut(Path As String, ExName As String, Frm As Long)

    Dim SH As Object, Link As Object

    Set SH = CreateObject("WScript.Shell")
    If IsObject(SH) Then
     Set Link = SH.CreateShortcut(SpecialFolder(7, Frm) & "\CoolPlayer.lnk")
     If IsObject(Link) Then
      With Link
       .Description = "CoolPlayer for cool people"
       .IconLocation = Path & "\" & ExName & ".exe"
       .TargetPath = Path & "\" & ExName & ".exe"
       .WindowStyle = 0
       .WorkingDirectory = Path
       .Save
      End With
     End If
    End If

End Sub
Public Function ConvertMinSec(Str As String) As Integer

    On Error GoTo MError
    Dim Num1 As Integer, St1 As Integer, St2 As Integer
    
    Num1 = InStr(1, Str, ":")
    St1 = CInt(Mid(Str, 1, Num1 - 1))
    St2 = CInt(Mid(Str, Num1 + 1))
    
    If St1 = 0 Then
     St1 = "00"
    ElseIf St2 = 0 Then
     St2 = "00"
    End If
    ConvertMinSec = (St1 * 60) + St2

MError:
    If Err.Number <> 0 Then ConvertMinSec = "00:00": Exit Function

End Function
Public Function GetTime(Time As String) As String

    On Error GoTo GetError
    Dim Tim As Integer

    Time = Format(CLng(Time) * 100, "000000")
    Tim = Val(Left(Time, 4))
    GetTime = Format(Tim \ 60, "0") & ":" & Format(Tim Mod 60, "00")

GetError:
    If Err.Number <> 0 Then Exit Function

End Function

Public Function BrowseDir(Form As Long, Msg As String, SetT As Boolean) As String

    On Error GoTo DirError
    Dim Pidl As Long, tmpPath As String
    Dim Pos As Integer

    With BI
     .hWnd = Form
     Call G.OnTop(.hWnd, SetT)
     .Root = 0&
     .Title = Msg
     .Flags = BIF_RETURNONLYFSDIRS
    End With

    Pidl = SHBrowseForFolder(BI)
    tmpPath = Space(1024)
    R = SHGetPathFromIDList(ByVal Pidl, ByVal tmpPath)

    If R Then
     Pos = InStr(tmpPath, Chr(0))
     tmpPath = Left(tmpPath, Pos - 1)
     BrowseDir = LCase(ValidateDir(tmpPath))
    Else
     BrowseDir = ""
    End If

DirError:
    If Err.Number <> 0 Then Exit Function

End Function
Public Function ValidateDir(ByVal Path As String) As String

    If Right(Path, 1) = "\" Then
     ValidateDir = Path
    Else
     If Path <> "" Then
      ValidateDir = Path & "\"
     Else
      ValidateDir = ""
     End If
    End If

End Function
Public Function MakeShort(FullName As String) As String

    On Error GoTo MError
    Dim strData As String, sDir As String, sTemp As String

    strData = FullName
    sDir = Left(strData, InStr(strData, "\"))
    sTemp = Left(Right(strData, Len(strData) - Len(sDir)), _
                 InStr(Right(strData, Len(strData) - Len(sDir)), "\"))

    Do
     strData = Right(strData, Len(strData) - InStr(strData, "\"))
    Loop Until InStr(strData, "\") = 0

    MakeShort = sDir & sTemp & "...\" & strData

MError:
    If Err.Number <> 0 Then MakeShort = FullName: Exit Function

End Function
Public Property Get Version() As String
    Version = "Misc_v1 v" & App.Major & "." & App.Minor & "." & App.Revision
End Property
