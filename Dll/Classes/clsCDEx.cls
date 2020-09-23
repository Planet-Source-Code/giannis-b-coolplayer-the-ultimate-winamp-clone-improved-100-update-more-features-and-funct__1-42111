VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDex"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private G As New clsGraph

Public Enum SFlags
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
    CDLCANCEL = 32755
End Enum

Private Type OPENFILENAME
    StructureSize As Long
    hwndOwner As Long
    Instance As Long
    Filter As String
    CustomFilter As String
    MaxCustFilter As Long
    FilterIndex As Long
    File As String
    MaxFile As Long
    FileTitle As String
    MaxFileTitle As Long
    InitialDir As String
    DialogTitle As String
    Flags As Long
    FileOffset As Integer
    FileExtension As Integer
    DefaultExt As String
    CustData As Long
    Hook As Long
    TemplateName As String
End Type

Private Type Dialog
    hOwner As Long
    FilterIndex As Long

    Filename As String
    FileTitle As String
    DialogTitle As String
    Filter As String
    InitDir As String
    SelDir As String
    DefaultExt As String
    DialogPrompt As String

    CancelError As Boolean
    Flags As SFlags
End Type

Private DLG As Dialog
Private OFN As OPENFILENAME
Private R As Long
Private ID As Integer

Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Property Let CancelError(ByVal Dat As Boolean)
    DLG.CancelError = Dat
End Property
Public Property Get CancelError() As Boolean
    CancelError = DLG.CancelError
End Property
Public Property Get hOwner() As Long
    hOwner = DLG.hOwner
End Property
Public Property Let hOwner(ByVal Own As Long)
    DLG.hOwner = Own
End Property
Public Property Get Flags() As SFlags
    Flags = DLG.Flags
End Property
Public Property Let Flags(ByVal NFlag As SFlags)
    DLG.Flags = NFlag
End Property
Public Property Get DialogTitle() As String
    DialogTitle = DLG.DialogTitle
End Property
Public Property Let DialogTitle(Tit As String)
    DLG.DialogTitle = Tit
End Property
Public Property Get DialogPrompt() As String
    DialogPrompt = DLG.DialogPrompt
End Property
Public Property Let DialogPrompt(ByVal Pro As String)
    DLG.DialogPrompt = Pro
End Property
Public Property Get DefaultExt() As String
    DefaultExt = DLG.DefaultExt
End Property
Public Property Let DefaultExt(Def As String)
    DLG.DefaultExt = Def
End Property
Public Property Get Filename() As String
    Filename = DLG.Filename
End Property
Public Property Let Filename(Name As String)
    DLG.Filename = Name
End Property
Public Property Get FileTitle() As String
    FileTitle = DLG.FileTitle
End Property
Public Property Let FileTitle(Tit As String)
    DLG.FileTitle = Tit
End Property
Public Property Get Filter() As String
    Filter = DLG.Filter
End Property
Public Property Let Filter(Fil As String)
    DLG.Filter = Fil
End Property
Public Property Get FilterIndex() As Long
    FilterIndex = DLG.FilterIndex
End Property
Public Property Let FilterIndex(Ind As Long)
    DLG.FilterIndex = Ind
End Property
Public Property Get InitDir() As String
    InitDir = DLG.InitDir
End Property
Public Property Let InitDir(sDir As String)
    DLG.InitDir = sDir
End Property
Public Property Get SelDir() As String
    SelDir = DLG.SelDir
End Property
Public Property Let SelDir(sDir As String)
    DLG.SelDir = sDir
End Property
Public Sub ShowOpen(Top As Boolean)
  
    On Error GoTo SError
    Call InitializeOFN(Top)
    R = GetOpenFileName(OFN)
    If R > 0 Then
     ID = InStr(OFN.FileTitle, vbNullChar)
     If ID Then DLG.FileTitle = Left(OFN.FileTitle, ID - 1)
     ID = InStr(OFN.File, vbNullChar)
     If ID Then DLG.Filename = Left(OFN.File, ID - 1)
     DLG.FilterIndex = OFN.FilterIndex
    Else
     If DLG.CancelError Then Err.Raise 0
    End If

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Public Sub ShowSave(Top As Boolean)
  
    On Error GoTo SError
    Call InitializeOFN(Top)
    R = GetSaveFileName(OFN)
    If R > 0 Then
     ID = InStr(OFN.FileTitle, vbNullChar)
     If ID Then DLG.FileTitle = Left(OFN.FileTitle, ID - 1)
     ID = InStr(OFN.File, vbNullChar)
     If ID Then DLG.Filename = Left(OFN.File, ID - 1)
     DLG.FilterIndex = OFN.FilterIndex
    Else
     If DLG.CancelError Then Err.Raise 0
    End If

SError:
    If Err.Number <> 0 Then Exit Sub

End Sub
Private Sub InitializeOFN(Top As Boolean)
    
    On Error GoTo IError
    Dim i As Integer

    With OFN
     .StructureSize = Len(OFN)
     .hwndOwner = DLG.hOwner
     Call G.OnTop(.hwndOwner, Top)
     .Flags = DLG.Flags And (&H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H100 Or &H200 Or &H400 Or &H800 Or &H1000 Or &H2000 Or &H4000 Or &H8000 Or &H40000 Or &H80000 Or &H100000 Or &H200000)
     .DefaultExt = DLG.DefaultExt
     .InitialDir = DLG.InitDir
     If (.Flags And OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT Then
      .Flags = .Flags Or OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_LONGNAMES
      .MaxFile = 32768
     Else
      .Flags = .Flags Or OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_LONGNAMES
      .MaxFile = 1024
      .File = DLG.Filename & String(255 - Len(DLG.Filename), 0)
      .FileTitle = String(255, 0)
     End If
     .MaxFileTitle = 255
     For i = 1 To Len(DLG.Filter)
      If Mid(DLG.Filter, i, 1) = "|" Then
       Mid(DLG.Filter, i, 1) = vbNullChar
      End If
     Next i
     .Filter = DLG.Filter & String(2, 0)
     .FilterIndex = DLG.FilterIndex
     .DialogTitle = DLG.DialogTitle
     .Instance = App.hInstance
    End With

IError:
    If Err.Number <> 0 Then Exit Sub

End Sub
