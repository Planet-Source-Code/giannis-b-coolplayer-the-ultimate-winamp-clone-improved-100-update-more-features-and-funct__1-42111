VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type ID3Info
    sTitle As String * 30
    sArtist As String * 30
    sAlbum As String * 30
    sYear  As String * 4
    sComments As String * 30
    sGenre As String * 1
End Type

Private sName As String
Private ID3 As ID3Info

Private Const mGenre As String = "|A Cappella|Acid|Acid Jazz|Acid Punk|Acoustic|Alt. Rock|Alternative|Ambient|Anime|Avantgarde|" & _
                     "Ballad|Bass|Beat|Bebob|Big Band|Black Metal|Bluegrass|Blues|Booty Bass|BritPop|Cabaret|Celtic|" & _
                     "Chamber Music|Chanson|Chorus|Christian Gangsta Rap|Christian Rap|Christian Rock|Classic Rock|" & _
                     "Classical|Club|Club-House|Comedy|Contemporary Christian|Country|Crossover|Cult|Dance|Dance Hall|" & _
                     "Darkwave|Death Metal|Disco|Dream|Drum & Bass|Drum Solo|Duet|Easy Listening|Electronic|Ethnic|" & _
                     "Eurodance|Euro-House|Euro-Techno|Fast-Fusion|Folk|Folk/Rock|Folklore|Freestyle|Funk|Fusion|Game|" & _
                     "Gangsta Rap|Goa|Gospel|Gothic|Gothic Rock|Grunge|Hard Rock|Hardcore|Heavy Metal|Hip-Hop|House|Humour|" & _
                     "Indie|Industrial|Instrumental|Instrumental Pop|Instrumental Rock|Jazz|Jazz+Funk|JPop|Jungle|Latin|Lo-Fi|" & _
                     "Meditative|Merengue|Metal|Musical|National Folk|Native American|Negerpunk|New Age|New Wave|Noise|Oldies|Opera|" & _
                     "Other|Polka|Polsk Punk|Pop|Pop/Funk|Pop-Folk|Porn Groove|Power Ballad|Pranks|Primus|Progressive Rock|Psychedelic|" & _
                     "Psychedelic Rock|Punk|Punk Rock|R&B|Rap|Rave|Reggae|Retro|Revival|Rhythmic Soul|Rock|Rock & Roll|Salsa|Samba|Satire|" & _
                     "Showtunes|Ska|Slow Jam|Slow Rock|Sonata|Soul|Sound Clip|Soundtrack|Southern Rock|Space|Speech|Swing|Syphonic Rock|" & _
                     "Symphony|Synthpop|Tango|Techno|Techno-Industrial|Terror|Thrash Metal|Top 40|Trailer|Trance|Tribal|Trip-Hop|Vocal"

Private Const mIndex As String = "255|123|34|74|73|99|40|20|26|145|90|116|41|135|85|96|138|89|0|107|132|65|88|104|102|97|136|61|141|1|32|112|" & _
                                 "128|57|140|2|139|58|3|125|50|22|4|55|127|122|120|98|52|48|124|25|54|84|115|80|81|119|5|30|36|59|126|38|49|91|" & _
                                 "6|79|129|137|7|35|100|131|19|33|46|47|146|8|29|63|86|71|45|142|9|77|82|64|133|10|66|39|11|103|12|75|134|13|53|" & _
                                 "62|109|117|23|108|92|67|93|43|121|14|15|68|16|76|87|118|17|78|143|114|110|69|21|111|95|105|42|37|24|56|44|101|" & _
                                 "83|94|106|147|113|18|51|130|60|70|31|144|72|27|28"
Public Sub ClearData()

    With ID3
     .sGenre = Chr("255")
     .sAlbum = ""
     .sArtist = ""
     .sComments = ""
     .sTitle = ""
     .sYear = ""
    End With

End Sub
Public Sub ClearName()
    sName = ""
End Sub

Public Sub SplitGenres(L As Object)

    Dim GenreArray() As String, sArray() As String
    Dim i As Integer, a As ComboBox
    Set a = L

    a.Clear
    GenreArray = Split(mGenre, "|"): sArray = Split(mIndex, "|")
    For i = LBound(GenreArray) To UBound(GenreArray)
     a.AddItem GenreArray(i)
     a.ItemData(a.NewIndex) = sArray(i)
    Next
    a.ListIndex = 0

End Sub
Public Function RemoveTag(s As String) As String

    On Error GoTo RError
    Dim a As String * 127

    Close #1
     Open s For Binary Access Write As #1
      Put #1, FileLen(s) - 127, a
    Close #1
    RemoveTag = "Info: ID3 Tag removed."
    If Err.Number = 0 Then Call ClearData

RError:
    If Err.Number <> 0 Then
     Close #1
     RemoveTag = "Info: " & Err.Description & ".": Exit Function
    End If

End Function
Public Property Get Filename() As String
    Filename = "" & RTrim(sName) & ""
End Property
Public Function ReadTag(s As String) As String
    
    On Error GoTo RError
    If Len(s) = 0 Then ReadTag = "Info: No valid filename.": Exit Function
    If FileExist(s) = False Then
     ReadTag = "Info: No valid filename.": Exit Function
    End If
    If HasTag(s) = True Then
     Close #1
      Open s For Binary Access Read As #1
       Get #1, FileLen(s) - 124, ID3
     Close #1
     ReadTag = "Info: ID3 Tag found."
    Else
     ReadTag = "Info: No ID3 Tag found."
    End If

RError:
    If Err.Number <> 0 Then
     Close #1
     ReadTag = "Info: " & Err.Description & ".": Exit Function
    End If

End Function

Public Function WriteTag(s As String) As String

    On Error GoTo WError
    If Len(s) = 0 Then WriteTag = "Info: No valid filename.": Exit Function
    If FileExist(s) = False Then
     WriteTag = "Info: No valid filename.": Exit Function
    End If
    If HasTag(s) = False Then
     Close #1
     Open s For Binary Access Write As #1
      Put #1, FileLen(s) - 127, "TAG"
      Put #1, FileLen(s) - 124, ID3
     Close #1
    Else
     Close #1
     Open s For Binary Access Write As #1
      Put #1, FileLen(s) - 124, ID3
     Close #1
    End If
    WriteTag = "Info: ID3 Tag saved."

WError:
    If Err.Number <> 0 Then
     Close #1
     WriteTag = "Info: " & Err.Description & ".": Exit Function
    End If

End Function
Public Property Get Artist() As String
    Artist = "" & RTrim(ID3.sArtist) & ""
End Property
Public Property Let Artist(s As String)
    ID3.sArtist = s
End Property
Public Property Get Comments() As String
    Comments = "" & RTrim(ID3.sComments) & ""
End Property
Public Property Let Comments(s As String)
    ID3.sComments = s
End Property
Public Property Get Title() As String
    Title = "" & RTrim(ID3.sTitle) & ""
End Property
Public Property Let Title(s As String)
    ID3.sTitle = s
End Property
Public Property Get Year() As String
    Year = "" & RTrim(ID3.sYear) & ""
End Property
Public Property Let Year(s As String)
    ID3.sYear = s
End Property
Public Property Get Genre() As Integer
    Genre = Asc(ID3.sGenre)
End Property
Private Function HasTag(s As String) As Boolean

    On Error GoTo HError
    Dim TAG As String * 3

    If FileExist(s) = False Then HasTag = False: Exit Function
    Close #1
     Open s For Binary As #1
      Get #1, FileLen(s) - 127, TAG
    Close #1
    HasTag = IIf(UCase(TAG) = "TAG", True, False)

HError:
    If Err.Number <> 0 Then
     Close #1
     HasTag = False: Exit Function
    End If

End Function
Private Function FileExist(s As String) As Boolean

    On Error Resume Next
    If FileLen(s) <> 0 Then
     FileExist = True: sName = s
    Else
     FileExist = False: sName = ""
    End If

End Function
Public Property Let Genre(Gen As Integer)
    ID3.sGenre = Chr(Gen)
End Property
Public Property Get Album() As String
    Album = "" & RTrim(ID3.sAlbum) & ""
End Property
Public Property Let Album(s As String)
    ID3.sAlbum = s
End Property
