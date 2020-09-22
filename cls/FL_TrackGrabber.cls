VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_TrackGrabber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************
'*** Flamed track grabber
'***
'*** grabbs Mode-1 tracks to ISO
'*** and audio tracks to WAV/MP3
'*******************************************

Public Enum FL_SAVETRACK
    ST_UNKNOWN_ERR
    ST_NOT_READY
    ST_ENCODER_INIT
    ST_INVALID_TRACKNO
    ST_INVALID_SESSION
    ST_INVALID_TRACKMODE
    ST_READ_ERR
    ST_WRITE_ERR
    ST_CANCELED
    ST_FINISHED
End Enum

Public Event Progress(ByVal Percent As Integer, _
                      ByVal Track As Integer, _
                      ByVal startLBA As Long, _
                      ByVal endLBA As Long, _
                      Cancel As Boolean)

Public Event AudioData(samples() As Integer)

Public Function AudioTrackToMP3(DriveID As String, _
                Track As Integer, _
                Optional file As String, _
                Optional Bitrate As FL_MP3ENC_BITRATES = [192 kBits], _
                Optional RelStartLBA As Long = -1, _
                Optional RelEndLBA As Long = -1) _
                As FL_SAVETRACK

    Dim encMP3      As New FL_EncoderMP3

    Dim fh          As Long
    Dim i           As Long
    Dim dwWritten   As Long

    Dim Buffer()    As Byte
    Dim encbuf()    As Byte
    Dim samples()   As Integer
    Dim Cancel      As Boolean
    Dim udtTrkNfo   As t_TrackInfo

    Dim blnLast     As Boolean

    Dim udtWavHdr   As t_WavHdr

    Dim sectors     As Long
        sectors = 20&

    ' 20 sectors with 2352 bytes userdata
    ReDim Buffer(2352& * sectors - 1&) As Byte

    ' need to convert?
    encMP3.OutputBitrate = Bitrate
    If Not encMP3.Init(Buffer, sectors * 2352&) Then
        AudioTrackToMP3 = ST_ENCODER_INIT
        Exit Function
    End If

    If Not CDRomTestUnitReady(DriveID) Then
        AudioTrackToMP3 = ST_NOT_READY
        Exit Function
    End If

    If Track < 1 Or Track > 100 Then
        AudioTrackToMP3 = ST_INVALID_TRACKNO
        Exit Function
    End If

    If Not CDRomTrackInfo(DriveID, Track, udtTrkNfo) Then
        AudioTrackToMP3 = ST_INVALID_TRACKNO
        Exit Function
    End If

    ' is audio track?
    If udtTrkNfo.DataMode <> e_SectorModes.MODE_AUDIO Then
        AudioTrackToMP3 = ST_INVALID_TRACKMODE
        Exit Function
    End If

    'end LBA user defined?
    If Not (RelEndLBA < 1&) And Not (udtTrkNfo.startLBA + RelEndLBA > udtTrkNfo.endLBA) Then _
        udtTrkNfo.endLBA = udtTrkNfo.startLBA + RelEndLBA

    'start LBA user defined?
    If Not (RelStartLBA < 0&) And Not (RelStartLBA > (udtTrkNfo.endLBA - udtTrkNfo.startLBA)) Then _
        udtTrkNfo.startLBA = udtTrkNfo.startLBA + RelStartLBA

    If Not file = "" Then

        fh = CreateFile(file, _
                        GENERIC_WRITE, _
                        0&, _
                        ByVal 0&, _
                        CREATE_ALWAYS, _
                        0&, 0&)
    
        ' could not create file
        If fh = -1& Then
            AudioTrackToMP3 = ST_WRITE_ERR
            Exit Function
        End If

    End If

    i = udtTrkNfo.startLBA

    Do While Not i >= udtTrkNfo.endLBA

        ' end of track?
        If i + sectors > udtTrkNfo.endLBA Then
            sectors = udtTrkNfo.endLBA - i
            ReDim Buffer(sectors * 2352& - 1) As Byte
            blnLast = True
        End If

        ' read a bunch of audio sectors
        If Not CDRomReadCD(DriveID, i, sectors, VarPtr(Buffer(0)), UBound(Buffer), &H10) Then
            AudioTrackToMP3 = ST_READ_ERR
            GoTo ExitFnc
        End If

        encbuf = encMP3.Encode(UBound(Buffer) + 1, blnLast)

        If Not file = "" Then

            ' write them to the file
            WriteFile fh, encbuf(0), UBound(encbuf) + 1, dwWritten, ByVal 0&
            If dwWritten <> UBound(encbuf) + 1 Then
                AudioTrackToMP3 = ST_WRITE_ERR
                GoTo ExitFnc
            End If

        End If

        i = i + sectors

        ' show progress
        RaiseEvent Progress(100 * (i - udtTrkNfo.startLBA) / (udtTrkNfo.endLBA - udtTrkNfo.startLBA), Track, udtTrkNfo.startLBA, udtTrkNfo.endLBA, Cancel)

        ' canceled?
        If Cancel Then
            AudioTrackToMP3 = ST_CANCELED
            GoTo ExitFnc
        End If

    Loop

    AudioTrackToMP3 = ST_FINISHED

ExitFnc:
    encMP3.DeInit
    If Not file = "" Then CloseHandle fh

End Function

Public Function AudioTrackToWAV(DriveID As String, _
                Track As Integer, _
                Optional file As String, _
                Optional WriteWavHeader As Boolean = True, _
                Optional SampleRate As Long = 44100, _
                Optional RelStartLBA As Long = -1, _
                Optional RelEndLBA As Long = -1) _
                As FL_SAVETRACK

    Dim encWAV      As New FL_EncoderWAV

    Dim fh          As Long
    Dim i           As Long
    Dim dwWritten   As Long

    Dim Buffer()    As Byte
    Dim encbuf()    As Byte
    Dim samples()   As Integer
    Dim Cancel      As Boolean
    Dim udtTrkNfo   As t_TrackInfo

    Dim blnLast     As Boolean

    Dim udtWavHdr   As t_WavHdr

    Dim sectors     As Long
        sectors = 20&

    ' 20 sectors with 2352 bytes userdata
    ReDim Buffer(2352& * sectors - 1&) As Byte
    ReDim samples((UBound(Buffer) + 1&) / 2& - 1&) As Integer

    ' need to convert?
    If SampleRate <> 44100 Then
        encWAV.OutputSampleRate = SampleRate
        If Not encWAV.Init(Buffer, sectors * 2352&) Then
            AudioTrackToWAV = ST_ENCODER_INIT
            Exit Function
        End If
    End If

    If Not CDRomTestUnitReady(DriveID) Then
        AudioTrackToWAV = ST_NOT_READY
        Exit Function
    End If

    If Track < 1 Or Track > 100 Then
        AudioTrackToWAV = ST_INVALID_TRACKNO
        Exit Function
    End If

    If Not CDRomTrackInfo(DriveID, Track, udtTrkNfo) Then
        AudioTrackToWAV = ST_INVALID_TRACKNO
        Exit Function
    End If

    ' is audio track?
    If udtTrkNfo.DataMode <> e_SectorModes.MODE_AUDIO Then
        AudioTrackToWAV = ST_INVALID_TRACKMODE
        Exit Function
    End If

    'end LBA user defined?
    If Not (RelEndLBA < 1&) And Not (udtTrkNfo.startLBA + RelEndLBA > udtTrkNfo.endLBA) Then _
        udtTrkNfo.endLBA = udtTrkNfo.startLBA + RelEndLBA

    'start LBA user defined?
    If Not (RelStartLBA < 0&) And Not (RelStartLBA > (udtTrkNfo.endLBA - udtTrkNfo.startLBA)) Then _
        udtTrkNfo.startLBA = udtTrkNfo.startLBA + RelStartLBA

    If Not file = "" Then

        fh = CreateFile(file, _
                        GENERIC_WRITE, _
                        0&, _
                        ByVal 0&, _
                        CREATE_ALWAYS, _
                        0&, 0&)
    
        ' could not create file
        If fh = -1& Then
            AudioTrackToWAV = ST_WRITE_ERR
            Exit Function
        End If

    End If

    If WriteWavHeader Then

        With udtWavHdr
            .riff = "RIFF"                              ' RIFF chunk
            .len = udtTrkNfo.Length * 2352& + 44& - 8&  ' length of the file
            .WavFmt = "WAVEfmt "                        ' WAVEfmt
            .HdrLen = 16&                               ' header length
            .format = 1                                 ' format
            .NumChannels = 2                            ' number of channels (stereo)
            .SampleRate = SampleRate                    ' sample rate (44.1kHz)
            .BlockAlign = 4                             ' Block Align
            .BitsPerSample = 16                         ' Bits per sample
            .BytesPerSec = .SampleRate * .BitsPerSample * 1 / 8 * .NumChannels
            .Data = "data"                              ' data chunk
            .DataLen = udtTrkNfo.Length * 2352& - 44&   ' data length without header length
        End With

        If Not file = "" Then

            ' write wave header
            WriteFile fh, udtWavHdr, Len(udtWavHdr), dwWritten, ByVal 0&
            If dwWritten <> Len(udtWavHdr) Then
                AudioTrackToWAV = ST_WRITE_ERR
                GoTo ExitFnc
            End If

        Else

            ' return the wave header
            Dim intHdr(21) As Integer
            CopyMemory intHdr(0), udtWavHdr, Len(udtWavHdr)
            RaiseEvent AudioData(intHdr)

        End If

    End If

    i = udtTrkNfo.startLBA

    Do While Not i >= udtTrkNfo.endLBA

        ' end of track?
        If i + sectors > udtTrkNfo.endLBA Then
            sectors = udtTrkNfo.endLBA - i
            ReDim Buffer(sectors * 2352& - 1) As Byte
            ReDim samples((UBound(Buffer) + 1) / 2 - 1) As Integer
            blnLast = True
        End If

        ' read a bunch of audio sectors
        If Not CDRomReadCD(DriveID, i, sectors, VarPtr(Buffer(0)), UBound(Buffer), &H10) Then
            AudioTrackToWAV = ST_READ_ERR
            GoTo ExitFnc
        End If

        If SampleRate <> 44100 Then
            encbuf = encWAV.Encode(UBound(Buffer) + 1, blnLast)
        Else
            encbuf = Buffer
        End If

        If Not file = "" Then

            ' write them to the file
            WriteFile fh, encbuf(0), UBound(encbuf) + 1, dwWritten, ByVal 0&
            If dwWritten <> UBound(encbuf) + 1 Then
                AudioTrackToWAV = ST_WRITE_ERR
                GoTo ExitFnc
            End If

        Else

            ReDim samples((UBound(encbuf) + 1) \ 2) As Integer

            ' byte buffer to integer buffer
            CopyMemory samples(0), encbuf(0), UBound(encbuf) + 1

            ' give the samples back
            RaiseEvent AudioData(samples)

        End If

        i = i + sectors

        ' show progress
        RaiseEvent Progress(100 * (i - udtTrkNfo.startLBA) / (udtTrkNfo.endLBA - udtTrkNfo.startLBA), Track, udtTrkNfo.startLBA, udtTrkNfo.endLBA, Cancel)

        ' canceled?
        If Cancel Then
            AudioTrackToWAV = ST_CANCELED
            GoTo ExitFnc
        End If

    Loop

    AudioTrackToWAV = ST_FINISHED

ExitFnc:
    encWAV.DeInit
    If Not file = "" Then CloseHandle fh

End Function

' save a mode 1 track as an ISO image
Public Function DataTrackToISO(DriveID As String, _
                Track As Integer, _
                file As String, _
                Optional RelStartLBA As Long = -1, _
                Optional RelEndLBA As Long = -1) As FL_SAVETRACK

    Dim fh          As Long
    Dim i           As Long
    Dim dwWritten   As Long

    Dim Buffer()    As Byte
    Dim Cancel      As Boolean
    Dim udtTrkNfo   As t_TrackInfo

    Dim sectors     As Long
        sectors = 20


    ' 20 sectors with 2048 bytes userdata
    ReDim Buffer(2048& * sectors - 1) As Byte


    If Not CDRomTestUnitReady(DriveID) Then
        DataTrackToISO = ST_NOT_READY
        Exit Function
    End If

    If Track < 1 Or Track > 100 Then
        DataTrackToISO = ST_INVALID_TRACKNO
        Exit Function
    End If

    If Not CDRomTrackInfo(DriveID, Track, udtTrkNfo) Then
        DataTrackToISO = ST_INVALID_TRACKNO
        Exit Function
    End If

    If udtTrkNfo.DataMode <> e_SectorModes.MODE_MODE1 Then
        DataTrackToISO = ST_INVALID_TRACKMODE
        Exit Function
    End If

    'end LBA user defined?
    If Not (RelEndLBA < 1) And Not (udtTrkNfo.startLBA + RelEndLBA > udtTrkNfo.endLBA) Then _
        udtTrkNfo.endLBA = udtTrkNfo.startLBA + RelEndLBA

    'start LBA user defined?
    If Not (RelStartLBA < 0) And Not (RelStartLBA < udtTrkNfo.startLBA) Then _
        udtTrkNfo.startLBA = udtTrkNfo.startLBA + RelStartLBA

    fh = CreateFile(file, _
                    GENERIC_WRITE, _
                    0, _
                    ByVal 0&, _
                    CREATE_ALWAYS, _
                    0, 0)

    ' could not create file
    If fh = -1 Then
        DataTrackToISO = ST_WRITE_ERR
        Exit Function
    End If

    i = udtTrkNfo.startLBA

    Do While Not i >= udtTrkNfo.endLBA

        ' end of track?
        If i + 20 > udtTrkNfo.endLBA Then
            ReDim Buffer((udtTrkNfo.endLBA - i) * 2048& - 1) As Byte
            sectors = udtTrkNfo.endLBA - i
        End If

        ' read a bunch of Mode 1 sectors
        If Not CDRomRead10(DriveID, i, VarPtr(Buffer(0)), UBound(Buffer), sectors) Then
            DataTrackToISO = ST_READ_ERR
            GoTo ExitFnc
        End If

        ' write them to the file
        WriteFile fh, Buffer(0), UBound(Buffer) + 1, dwWritten, ByVal 0&
        If dwWritten <> UBound(Buffer) + 1 Then
            DataTrackToISO = ST_READ_ERR
            GoTo ExitFnc
        End If

        i = i + 20

        ' show progress
        RaiseEvent Progress(100 * (i - udtTrkNfo.startLBA) / (udtTrkNfo.endLBA - udtTrkNfo.startLBA), Track, udtTrkNfo.startLBA, udtTrkNfo.endLBA, Cancel)

        ' canceled?
        If Cancel Then
            DataTrackToISO = ST_CANCELED
            GoTo ExitFnc
        End If

    Loop

    DataTrackToISO = ST_FINISHED

ExitFnc:
    CloseHandle fh

End Function
