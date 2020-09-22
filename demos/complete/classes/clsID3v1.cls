VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3v1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' cMP3ID3v1.cls
' Read and Write MP3 ID3 v1 tags.
'
' http://vbaccelerator.com/
' S McMahon
' 2003-01-26
'
' Updates
' 2003/05/07: Added ID3 v1.1 support.
' Thanks to Don Beal for providing the code!
'
' 2004/01/19: Added code to strip null characters
' from the strings returned by the tag properties.


Private Type MP3ID3V1Tag
  Tag As String * 3           '-- 03
  Title As String * 30     '-- 33
  Artist As String * 30       '-- 63
  Album As String * 30        '-- 93
  Year As String * 4          '-- 97
  Comment As String * 28      '-- 125
  ' START: Mods for ID3 v1.1 support
  Filler As Byte              '-- 126
  Track As Byte               '-- 127
  ' END: Mods for ID3 v1.1. support
  Genre As Byte               '-- 128
End Type

Private m_sMp3File As String

Private m_bHasID3v1Tag As Boolean

Private m_sTag As String
Private m_sTitle As String
Private m_sArtist As String
Private m_sAlbum As String
Private m_sYear As String
Private m_sComment As String
Private m_sGenre As Byte
' Added for ID3 v1.1 support:
Private m_sTrack As Byte


Public Property Get MP3File() As String
   MP3File = m_sMp3File
End Property

Public Property Let MP3File(ByVal Value As String)
   m_sMp3File = Value
   pLoadTag
End Property

Public Sub Update()
   pUpdateTag
End Sub

Public Property Get HasID3v1Tag() As Boolean
   HasID3v1Tag = m_bHasID3v1Tag
End Property

Public Property Get Title() As String
   Title = m_sTitle
End Property
Public Property Let Title(ByVal Value As String)
   m_sTitle = Value
End Property

Public Property Get Artist() As String
   Artist = m_sArtist
End Property
Public Property Let Artist(ByVal Value As String)
   m_sArtist = Value
End Property

Public Property Get Album() As String
   Album = m_sAlbum
End Property
Public Property Let Album(ByVal Value As String)
   m_sAlbum = Value
End Property

Public Property Get Year() As String
   Year = m_sYear
End Property
Public Property Let Year(ByVal Value As String)
   m_sYear = Value
End Property

' START: Added for ID3 v1.1 support.
Public Property Get Track() As Byte
   Track = m_sTrack
End Property
Public Property Let Track(ByVal Value As Byte)
   m_sTrack = Value
End Property
' END: Added for ID3 v1.1 support.

Public Property Get Comment() As String
   Comment = m_sComment
End Property
Public Property Let Comment(ByVal Value As String)
   m_sComment = Value
End Property

Public Property Get Genre() As Byte
   Genre = m_sGenre
End Property
Public Property Let Genre(ByVal Value As Byte)
   m_sGenre = Value
End Property

Public Property Get GenreName(ByVal Genre As Byte) As String
Dim sName As String
   Select Case Genre
   Case 34: sName = "Acid"
   Case 74: sName = "Acid Jazz"
   Case 73: sName = "Acid Punk"
   Case 99: sName = "Acoustic"
   Case 40: sName = "Alt.Rock"
   Case 20: sName = "Alternative"
   Case 26: sName = "Ambient"
   Case 145: sName = "Anime"
   Case 90: sName = "Avant Garde"
   Case 116: sName = "Ballad"
   Case 41: sName = "Bass"
   Case 135: sName = "Beat"
   Case 85: sName = "Bebob"
   Case 96: sName = "Big Band"
   Case 138: sName = "Black Metal"
   Case 89: sName = "Blue Grass"
   Case 0: sName = "Blues"
   Case 107: sName = "Booty Bass"
   Case 132: sName = "Brit Pop"
   Case 65: sName = "Cabaret"
   Case 88: sName = "Celtic"
   Case 104: sName = "Chamber Music"
   Case 102: sName = "Chanson"
   Case 97: sName = "Chorus"
   Case 136: sName = "Christian Gangsta Rap"
   Case 61: sName = "Christian Rap"
   Case 141: sName = "Christian Rock"
   Case 1: sName = "Classic Rock"
   Case 32: sName = "Classical"
   Case 112: sName = "Club"
   Case 128: sName = "Club - House"
   Case 57: sName = "Comedy"
   Case 140: sName = "Contemporary Christian"
   Case 2: sName = "Country"
   Case 139: sName = "Crossover"
   Case 58: sName = "Cult"
   Case 3: sName = "Dance"
   Case 125: sName = "Dance Hall"
   Case 50: sName = "Darkwave"
   Case 22: sName = "Death Metal"
   Case 4: sName = "Disco"
   Case 55: sName = "Dream"
   Case 127: sName = "Drum & Bass"
   Case 122: sName = "Drum Solo"
   Case 120: sName = "Duet"
   Case 98: sName = "Easy Listening"
   Case 52: sName = "Electronic"
   Case 48: sName = "Ethnic"
   Case 54: sName = "Eurodance"
   Case 124: sName = "Euro - House"
   Case 25: sName = "Euro - Techno"
   Case 84: sName = "Fast Fusion"
   Case 80: sName = "Folk"
   Case 81: sName = "Folk / Rock"
   Case 115: sName = "Folklore"
   Case 119: sName = "Freestyle"
   Case 5: sName = "Funk"
   Case 30: sName = "Fusion"
   Case 36: sName = "Game"
   Case 59: sName = "Gangsta Rap"
   Case 126: sName = "Goa"
   Case 38: sName = "Gospel"
   Case 49: sName = "Gothic"
   Case 91: sName = "Gothic Rock"
   Case 6: sName = "Grunge"
   Case 79: sName = "Hard Rock"
   Case 129: sName = "Hardcore"
   Case 137: sName = "Heavy Metal"
   Case 7: sName = "Hip Hop"
   Case 35: sName = "House"
   Case 100: sName = "Humour"
   Case 131: sName = "Indie"
   Case 19: sName = "Industrial"
   Case 33: sName = "Instrumental"
   Case 46: sName = "Instrumental Pop"
   Case 47: sName = "Instrumental Rock"
   Case 8: sName = "Jazz"
   Case 29: sName = "Jazz - Funk"
   Case 146: sName = "JPop"
   Case 63: sName = "Jungle"
   Case 86: sName = "Latin"
   Case 71: sName = "Lo - fi"

   Case 45: sName = "Meditative"
   Case 142: sName = "Merengue"
   Case 9: sName = "Metal"
   Case 77: sName = "Musical"
   Case 82: sName = "National Folk"

   Case 64: sName = "Native American"
   Case 133: sName = "Negerpunk"
   Case 10: sName = "New Age"
   Case 66: sName = "New Wave"

   Case 39: sName = "Noise"
   Case 11: sName = "Oldies"

   Case 103: sName = "Opera"
   Case 12: sName = "Other"
   Case 75: sName = "Polka"
   Case 134: sName = "Polsk Punk"

   Case 13: sName = "Pop"
   Case 62: sName = "Pop / Funk"
   Case 53: sName = "Pop / Folk"
   Case 109: sName = "Pr0n Groove"

   Case 117: sName = "Power Ballad"
   Case 23: sName = "Pranks"
   Case 108: sName = "Primus"
   Case 92: sName = "Progressive Rock"
   Case 67: sName = "Psychedelic"
   Case 93: sName = "Psychedelic Rock"

   Case 43: sName = "Punk"
   Case 121: sName = "Punk Rock"
   Case 14: sName = "R&B"
   Case 15: sName = "Rap"

   Case 68: sName = "Rave"
   Case 16: sName = "Reggae"
   Case 76: sName = "Retro"
   Case 87: sName = "Revival"
   Case 118: sName = "Rhythmic Soul"

   Case 17: sName = "Rock"
   Case 78: sName = "Rock 'n'Roll"
   Case 143: sName = "Salsa"
   Case 114: sName = "Samba"
   Case 110: sName = "Satire"
   Case 69: sName = "Showtunes"
   Case 21: sName = "Ska"
   Case 111: sName = "Slow Jam"

   Case 95: sName = "Slow Rock"
   Case 105: sName = "Sonata"
   Case 42: sName = "Soul"
   Case 37: sName = "Sound Clip"
   Case 24: sName = "Soundtrack"
   Case 56: sName = "Southern Rock"
   Case 44: sName = "Space"
   Case 101: sName = "Speech"
   Case 83: sName = "Swing"
   Case 94: sName = "Symphonic Rock"
   Case 106: sName = "Symphony"
   Case 147: sName = "Synth Pop"

   Case 113: sName = "Tango"
   Case 18: sName = "Techno"
   Case 51: sName = "Techno - Industrial"
   Case 130: sName = "Terror"
   Case 144: sName = "Thrash Metal"
   Case 60: sName = "Top 40"
   Case 70: sName = "Trailer"

   Case 31: sName = "Trance"
   Case 72: sName = "Tribal"

   Case 27: sName = "Trip Hop"
   Case 28: sName = "Vocal"
   
   End Select
   GenreName = sName
End Property

Private Sub pLoadTag()
Dim iFile As Integer
Dim lErr As Long
Dim sErr As String
   
   m_bHasID3v1Tag = False
   m_sComment = ""
   m_sArtist = ""
   m_sAlbum = ""
   m_sYear = ""
   m_sGenre = 255
   m_sTitle = ""
   ' Added for ID3 v1.1 support:
   m_sTrack = 0
   
   iFile = FreeFile
   On Error Resume Next
   Open m_sMp3File For Binary Access Read Lock Write As #iFile
   If (Err.Number <> 0) Then
      lErr = Err.Number
      sErr = Err.Description
      On Error GoTo 0
      Err.Raise lErr, App.EXEName & ".cMP3ID3v1", sErr
   Else
      On Error GoTo 0
      If LOF(iFile) > 128 Then
         Dim Tag As MP3ID3V1Tag
         Get #iFile, LOF(iFile) - 127, Tag.Tag
         If Not (StrComp(Tag.Tag, "TAG") = 0) Then
            '  no tag
         Else
            m_bHasID3v1Tag = True
            Get #iFile, , Tag.Title
            m_sTitle = psStripNulls(Tag.Title)
            Get #iFile, , Tag.Artist
            m_sArtist = psStripNulls(Tag.Artist)
            Get #iFile, , Tag.Album
            m_sAlbum = psStripNulls(Tag.Album)
            Get #iFile, , Tag.Year
            m_sYear = psStripNulls(Tag.Year)
            Get #iFile, , Tag.Comment
            m_sComment = psStripNulls(Tag.Comment)
            ' START: Added for ID3 v1.1 support:
            Get #iFile, , Tag.Filler
            Get #iFile, , Tag.Track
            m_sTrack = psStripNulls(Tag.Track)
            ' END: added for ID3 v1.1 support
            Get #iFile, , Tag.Genre
            m_sGenre = psStripNulls(Tag.Genre)
            
         End If
      End If
   End If
   On Error Resume Next
   Close #iFile
   On Error GoTo 0
   Err.Clear
   
End Sub
Private Function psStripNulls(ByVal sBuf As String) As String
Dim I As Long
Dim sTag As String
   
   sTag = ""
   For I = 1 To Len(sBuf)
      If Not (Asc(Mid(sBuf, I, 1)) = 0) Then
         sTag = sTag & Mid(sBuf, I, 1)
      End If
   Next I
   psStripNulls = Trim(sTag)
   
End Function
Private Sub pUpdateTag()
Dim iFile As Integer
Dim lErr As Long
Dim sErr As String
   iFile = FreeFile
   On Error Resume Next
   Open m_sMp3File For Binary Access Read Write Lock Write As #iFile
   If (Err.Number <> 0) Then
      lErr = Err.Number
      sErr = Err.Description
      On Error GoTo 0
      Err.Raise lErr, App.EXEName & ".cMP3ID3v1", sErr
   Else
      Dim Tag As MP3ID3V1Tag
      If LOF(iFile) > 0 Then
         If LOF(iFile) > 128 Then
            Get #iFile, LOF(iFile) - 127, Tag.Tag
            If Not (StrComp(Tag.Tag, "TAG") = 0) Then
               ' no MP3 tag already, need to extend the file
               ' to add it
               Seek #iFile, LOF(iFile)
               Tag.Tag = "TAG"
               Put #iFile, , Tag.Tag
            End If
            LSet Tag.Title = m_sTitle
            LSet Tag.Artist = m_sArtist
            LSet Tag.Album = m_sAlbum
            LSet Tag.Year = m_sYear
            ' START: added for ID3 v1.1 support
            Tag.Track = m_sTrack
            ' END: added for ID3 v1.1 support
            LSet Tag.Comment = m_sComment
            Tag.Genre = m_sGenre
            
            Put #iFile, , Tag.Title
            Put #iFile, , Tag.Artist
            Put #iFile, , Tag.Album
            Put #iFile, , Tag.Year
            Put #iFile, , Tag.Comment
            ' START: added for ID3 v1.1 support
            Put #iFile, , Tag.Filler
            Put #iFile, , Tag.Track
            ' END: added for ID3 v1.1 support
            Put #iFile, , Tag.Genre
            
         End If
      Else
         On Error Resume Next
         Close #iFile
         On Error GoTo 0
         Err.Raise vbObjectError + 4097, App.EXEName & ".cMP3ID3v1", m_sMp3File & _
        " is not a valid MP3 file."
      End If
   End If
   
   On Error Resume Next
   Close #iFile
   On Error GoTo 0
   Err.Clear
   
End Sub
