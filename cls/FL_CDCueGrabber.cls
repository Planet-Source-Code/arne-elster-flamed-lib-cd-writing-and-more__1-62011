VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_CueGrabber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************
'*** Flamed Cue Grabber
'***
'*** creates Bin/Cue images from sessions
'*******************************************

Public Event Progress(ByVal Percent As Integer, _
                      ByVal Track As Integer, _
                      ByVal startLBA As Long, _
                      ByVal endLBA As Long, _
                      Cancel As Boolean)

Private lngBINwritten   As Long

Private Function AppendTrackRAWtoFile(DriveID As String, _
                Track As Integer, _
                fh As Long, _
                SessStartLBA As Long, _
                SessEndLBA As Long) As FL_SAVETRACK

    Dim cTrkNfo     As New FL_TrackInfo
    Dim i           As Long
    Dim blnCancel   As Boolean

    ' read buffer
    Dim Buffer()    As Byte

    ' written bytes
    Dim dwWritten   As Long

    ' sectors per read
    Dim sectors     As Long
        sectors = 20

    ReDim Buffer(sectors * 2352& - 1&) As Byte

    If Not cTrkNfo.GetInfo(DriveID, Track) Then
        AppendTrackRAWtoFile = ST_INVALID_TRACKNO
        Exit Function
    End If

    i = cTrkNfo.TrackStart.LBA

    Do While Not i >= cTrkNfo.TrackEnd.LBA

        ' end of track?
        If i + sectors > cTrkNfo.TrackEnd.LBA Then
            sectors = cTrkNfo.TrackEnd.LBA - i
            ReDim Buffer(sectors * 2352& - 1&) As Byte
        End If

        ' read a bunch of raw sectors
        If Not CDRomReadCD(DriveID, i, sectors, VarPtr(Buffer(0)), UBound(Buffer), &HF8) Then
            AppendTrackRAWtoFile = ST_READ_ERR
            Exit Function
        End If

        ' write them to the file
        WriteFile fh, Buffer(0), UBound(Buffer) + 1, dwWritten, ByVal 0&
        If dwWritten <> UBound(Buffer) + 1 Then
            AppendTrackRAWtoFile = ST_WRITE_ERR
            Exit Function
        End If

        lngBINwritten = lngBINwritten + dwWritten

        i = i + sectors

        RaiseEvent Progress((i - SessStartLBA) / (SessEndLBA - SessStartLBA) * 100, Track, SessStartLBA, SessEndLBA, blnCancel)

        If blnCancel Then
            blnCancel = False
            AppendTrackRAWtoFile = ST_CANCELED
            Exit Function
        End If

        DoEvents

    Loop

    AppendTrackRAWtoFile = ST_FINISHED

End Function

Public Function SessionToBinCue(DriveID As String, _
                Session As Integer, _
                binfile As String, _
                cuefile As String) As FL_SAVETRACK

    ' session info class
    Dim cSessNfo    As New FL_SessionInfo
    ' CD info class
    Dim cCDNfo      As New FL_CDInfo

    ' track reader return value
    Dim udeStat As FL_SAVETRACK

    ' file handles
    Dim fhBIN       As Long
    Dim fhCUE       As Long

    ' counters
    Dim i           As Long

    lngBINwritten = 0

    If Not cCDNfo.GetInfo(DriveID) Then
        SessionToBinCue = ST_NOT_READY
        Exit Function
    End If

    If Session < 1 Or Session > cCDNfo.Sessions Then
        SessionToBinCue = ST_INVALID_SESSION
        Exit Function
    End If

    If Not cSessNfo.GetInfo(DriveID, Session) Then
        SessionToBinCue = ST_INVALID_SESSION
        Exit Function
    End If

    fhBIN = CreateFile(binfile, _
                       GENERIC_WRITE, _
                       0, _
                       ByVal 0&, _
                       CREATE_ALWAYS, _
                       0, 0)

    If fhBIN = -1 Then
        SessionToBinCue = ST_WRITE_ERR
        Exit Function
    End If

    fhCUE = CreateFile(cuefile, _
                       GENERIC_WRITE, _
                       0, _
                       ByVal 0&, _
                       CREATE_ALWAYS, _
                       0, 0)

    If fhCUE = -1 Then
        SessionToBinCue = ST_WRITE_ERR
        GoTo ExitFnc
    End If

    ' bin and cue in the same dir?
    If StrComp(PathFromPathFile(binfile), PathFromPathFile(cuefile), vbTextCompare) = 0 Then
        binfile = FileFromPathFile(binfile)
    End If

    ' write the cue header
    If Not WriteCueHeader(fhCUE, binfile) Then
        SessionToBinCue = ST_WRITE_ERR
        GoTo ExitFnc
    End If

    ' go thorugh all tracks of the session
    For i = cSessNfo.FirstTrack To cSessNfo.LastTrack

        ' read track i
        udeStat = AppendTrackRAWtoFile(DriveID, _
                    CInt(i), _
                    fhBIN, _
                    cSessNfo.FirstTrackLBA, _
                    cSessNfo.LeadOutLBA)

        Select Case udeStat

            Case ST_FINISHED:
                '
            Case Else:
                SessionToBinCue = udeStat
                GoTo ExitFnc

        End Select

        ' create a cue entry for the track
        If Not WriteCueEntry(fhCUE, DriveID, CInt(i), CInt(i - cSessNfo.FirstTrack + 1), , -cSessNfo.FirstTrackLBA) Then
            SessionToBinCue = ST_WRITE_ERR
            GoTo ExitFnc
        End If

    Next

    ' finished
    SessionToBinCue = ST_FINISHED

ExitFnc:
    ' close file handles
    CloseHandle fhBIN
    CloseHandle fhCUE

End Function

Private Function WriteCueHeader(fh As Long, binfile As String) As Boolean

    Dim strHeader   As String
    Dim bytHeader() As Byte
    Dim dwWritten   As Long

    strHeader = "FILE """ & binfile & """ BINARY" & vbCrLf

    bytHeader = StrConv(strHeader, vbFromUnicode)

    WriteFile fh, bytHeader(0), Len(strHeader), dwWritten, ByVal 0&
    WriteCueHeader = dwWritten = Len(strHeader)

End Function

Private Function WriteCueEntry(fh As Long, _
                DriveID As String, _
                Track As Integer, _
                reltrack As Integer, _
                Optional index As Integer = 1, _
                Optional AddLBA As Long = 0) As Boolean

    Dim cTrkNfo     As New FL_TrackInfo
    Dim strEntry    As String
    Dim bytEntry()  As Byte
    Dim dwWritten   As Long

    Const strLevel1 As String = "  "
    Const strLevel2 As String = "    "

    cTrkNfo.GetInfo DriveID, Track

    strEntry = strLevel1 & "TRACK " & format(reltrack, "00") & " " & TrackMode2Str(cTrkNfo.mode) & vbCrLf
    strEntry = strEntry & strLevel2 & "INDEX " & format(index, "00") & " " & cd.LBA2MSF(cTrkNfo.TrackStart.LBA - 150 + AddLBA) & vbCrLf

    bytEntry = StrConv(strEntry, vbFromUnicode)

    WriteFile fh, bytEntry(0), Len(strEntry), dwWritten, ByVal 0&
    WriteCueEntry = dwWritten = Len(strEntry)

End Function

Private Function TrackMode2Str(mode As FL_TrackModes) As String
    Select Case mode
        Case e_SectorModes.MODE_AUDIO:  TrackMode2Str = "AUDIO"
        Case e_SectorModes.MODE_MODE1:  TrackMode2Str = "MODE1/2352"
        Case e_SectorModes.MODE_MODE2:  TrackMode2Str = "MODE2/2352"
        Case e_SectorModes.MODE_MODE2_FORM1: TrackMode2Str = "MODE2FORM1/2352"
        Case e_SectorModes.MODE_MODE2_FORM2: TrackMode2Str = "MODE2FORM2/2352"
    End Select
End Function
