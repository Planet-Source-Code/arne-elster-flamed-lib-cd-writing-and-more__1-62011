VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_CueReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************
'*** Flamed Cue Reader
'***
'*** reads cue sheets and extracts tracks
'*** from bin files
'*******************************************

Private Type Track
    mode As String
    index(98) As Long
    firstindex As Integer
    indexcount As Integer
End Type

Private Type Tracks
    Tracks(98) As Track
    count As Integer
End Type
Private udtTracks   As Tracks

Private Enum FL_CUEEXPECTED
    EXPECTED_FILE
    EXPECTED_BINNAME
    EXPECTED_BINARY
    EXPECTED_TRACK
    EXPECTED_TRACKNUM
    EXPECTED_TRACKMODE
    EXPECTED_INDEX
    EXPECTED_INDEXNUM
    EXPECTED_INDEXMSF
    EXPECTED_TRACK_INDEX
End Enum

Public Enum FL_OPENCUE
    CUE_UNKNOWN_ERROR
    CUE_CUEMISSING
    CUE_BINMISSING
    CUE_CUEFILEEXPECTED
    CUE_BINFILEEXPECTED
    CUE_BINARYEXPECTED
    CUE_TRACKEXPECTED
    CUE_TRACKNUMEXPECTED
    CUE_INDEXEXPECTED
    CUE_INDEXNUMEXPECTED
    CUE_INDEXMSFEXPECTED
    CUE_OK
End Enum

Private strCueFile  As String
Private strBinFile  As String

Public Event ExtractProgress(ByVal Percent As Integer, Cancel As Boolean)

Public Function ExtractTrack(ByVal intTrack As Integer, _
                             ByVal strFile As String) As Boolean

    'On Error GoTo ErrorHandler

    Const ChunkSize As Long = 64& * 1024&

    '   start and end offset
    Dim lngStart    As Long, lngEnd     As Long
    '   file handles
    Dim fh          As Integer, fh2     As Integer
    '   read buffer
    Dim strBuffer   As String

    Dim blnCancel   As Boolean

    ' get LBA of intTrack's index 1
    With udtTracks.Tracks(intTrack - 1)
        If .firstindex = 0 Then
            lngStart = .index(1) * 2352&
        Else
            lngStart = .index(0) * 2352&
        End If
    End With

    ' get the LBA of the next track
    If intTrack = udtTracks.count Then
        lngEnd = FileLen(strBinFile) - lngStart
    Else
        lngEnd = udtTracks.Tracks(intTrack).index(0) * 2352&
    End If

    ' open input and output
    fh = FreeFile
    Open strBinFile For Binary As #fh
    fh2 = FreeFile
    Open strFile For Binary As #fh2

        Seek #fh, lngStart + 1

        ' extract data
        Do Until Seek(fh) >= lngEnd
            If ChunkSize + Seek(fh) > lngEnd Then
                strBuffer = Space(lngEnd - Seek(fh))
            Else
                strBuffer = Space(ChunkSize)
            End If
            Get #fh, , strBuffer
            Put #fh2, , CStr(strBuffer)

            RaiseEvent ExtractProgress((Seek(fh) - lngStart) / (lngEnd - lngStart) * 100, blnCancel)

            If blnCancel Then Exit Do

            DoEvents
        Loop

    Close #fh, #fh2

    ExtractTrack = True

ErrorHandler:

End Function

Private Function StrToTrackMode(aval As String) As FL_TrackModes
    Select Case aval
        Case "AUDIO": StrToTrackMode = FL_TrackModes.MODE_AUDIO
        Case "MODE1/2352": StrToTrackMode = FL_TrackModes.MODE_MODE1
        Case "MODE2/2352": StrToTrackMode = FL_TrackModes.MODE_MODE2
        Case "MODE2FORM1/2352": StrToTrackMode = FL_TrackModes.MODE_MODE2_FORM1
        Case "MODE2FORM2/2352": StrToTrackMode = FL_TrackModes.MODE_MODE2_FORM2
    End Select
End Function

Public Property Get TrackMode(Track As Integer) As FL_TrackModes
    TrackMode = StrToTrackMode(udtTracks.Tracks(Track - 1).mode)
End Property

Public Property Get TrackIndexLBA(Track As Integer, index As Integer) As Integer
    TrackIndexLBA = udtTracks.Tracks(Track - 1).index(index)
End Property

Public Property Get TrackIndexCount(Track As Integer) As Integer
    TrackIndexCount = udtTracks.Tracks(Track - 1).indexcount
End Property

Public Property Get TrackIndexFirst(Track As Integer) As Integer
    TrackIndexFirst = udtTracks.Tracks(Track - 1).firstindex
End Property

Public Property Get TrackCount() As Integer
    TrackCount = udtTracks.count
End Property

Public Function OpenCue(ByVal strFile As String) As FL_OPENCUE
' aka parse cue sheet

    On Error GoTo ErrorHandler

    Dim cMSF        As New FL_MSF

    '   file content
    Dim strContent  As String
    Dim strBuffer   As String
    Dim strLines()  As String
    Dim strCue()    As String
    '   expected lines
    Dim udeExpect   As FL_CUEEXPECTED
    '   counters
    Dim i           As Integer, intIndex As Integer
    '   file handle
    Dim FF          As Integer
    FF = FreeFile

    If Not FileExists(strFile) Then
        Exit Function
    End If

    Open strFile For Binary Access Read As #FF
        strContent = Space(LOF(FF))
        Get #1, , strContent
    Close #FF

    ' reset data
    ReDim strCue(0) As String
    udtTracks.count = 0
    strCueFile = vbNullString
    strBinFile = vbNullString

    strLines = Split(strContent, vbCrLf)
    For i = LBound(strLines) To UBound(strLines)

        strLines(i) = UCase$(Trim$(strLines(i)))

        If Not strLines(i) = vbNullString Then
            strCue(UBound(strCue)) = strLines(i)
            ReDim Preserve strCue(UBound(strCue) + 1) As String
        End If

    Next

    udeExpect = EXPECTED_FILE
    i = 0

    Do Until i = UBound(strCue)

        Select Case udeExpect

            ' file identifier
            ' only allowed once per cue sheet
            Case EXPECTED_FILE

                If Left$(strCue(i), Len("FILE")) <> "FILE" Then
                    OpenCue = CUE_CUEFILEEXPECTED
                    Exit Function
                End If

                udeExpect = EXPECTED_BINNAME

            ' filename of binary data
            Case EXPECTED_BINNAME

                If InStr(strCue(i), Chr(34)) < 1 Then
                    OpenCue = CUE_BINFILEEXPECTED
                    Exit Function
                End If

                strBuffer = Mid$(strCue(i), InStr(strCue(i), Chr(34)) + 1)
                strBuffer = Left$(strBuffer, InStrRev(strBuffer, Chr(34)) - 1)

                If Not FileExists(strBuffer) Then
                    If Not FileExists(PathFromPathFile(strFile) & strBuffer) Then
                        OpenCue = CUE_BINMISSING
                        Exit Function
                    Else
                        strBinFile = PathFromPathFile(strFile) & strBuffer
                    End If
                Else
                    strBinFile = strBuffer
                End If

                udeExpect = EXPECTED_BINARY

            ' type of raw data
            ' only support for binary
            Case EXPECTED_BINARY

                If Not Right$(strCue(i), Len("BINARY")) = "BINARY" Then
                    OpenCue = CUE_BINARYEXPECTED
                    Exit Function
                End If
    
                strCueFile = strFile

                i = i + 1
                udeExpect = EXPECTED_TRACK

            ' track identifier
            Case EXPECTED_TRACK

                If Not Left$(strCue(i), Len("TRACK")) = "TRACK" Then
                    OpenCue = CUE_TRACKEXPECTED
                    Exit Function
                End If

                udeExpect = EXPECTED_TRACKNUM

            ' track number
            Case EXPECTED_TRACKNUM

                strBuffer = Trim$(Mid(strCue(i), 7, 2))

                If Not IsNumeric(strBuffer) Then
                    OpenCue = CUE_TRACKNUMEXPECTED
                    Exit Function
                End If

                udeExpect = EXPECTED_TRACKMODE

            ' get track mode
            Case EXPECTED_TRACKMODE

                strBuffer = Mid$(strCue(i), 10)

                With udtTracks
                    .Tracks(.count).mode = strBuffer
                    .Tracks(.count).firstindex = -1
                    .count = .count + 1
                End With

                i = i + 1
                udeExpect = EXPECTED_INDEX

            ' index identifier
            Case EXPECTED_INDEX

                If Not Left$(strCue(i), Len("INDEX")) = "INDEX" Then
                    OpenCue = CUE_INDEXEXPECTED
                    Exit Function
                End If

                udeExpect = EXPECTED_INDEXNUM

            ' number of current index
            Case EXPECTED_INDEXNUM

                strBuffer = Trim$(Mid$(strCue(i), 7, 2))

                If Not IsNumeric(strBuffer) Then
                    OpenCue = CUE_INDEXNUMEXPECTED
                    Exit Function
                End If

                intIndex = CInt(strBuffer)

                udeExpect = EXPECTED_INDEXMSF

            ' MSF start time of index
            Case EXPECTED_INDEXMSF

                strBuffer = Right$(strCue(i), 8)

                If Not strBuffer Like "*:*:*" Then
                    OpenCue = CUE_INDEXMSFEXPECTED
                    Exit Function
                End If

                cMSF.MSF = strBuffer

                With udtTracks.Tracks(udtTracks.count - 1)
                    If .firstindex = -1 Then .firstindex = intIndex
                    .index(.indexcount) = cMSF.LBA + 150
                    .indexcount = intIndex
                End With

                i = i + 1

                udeExpect = EXPECTED_TRACK_INDEX

            ' next line should be either
            ' a track or index identifier
            Case EXPECTED_TRACK_INDEX

                Select Case Left$(strCue(i), 5)

                    Case "TRACK": udeExpect = EXPECTED_TRACK
                    Case "INDEX": udeExpect = EXPECTED_INDEX

                End Select

        End Select

    Loop

    OpenCue = CUE_OK

ErrorHandler:

End Function
