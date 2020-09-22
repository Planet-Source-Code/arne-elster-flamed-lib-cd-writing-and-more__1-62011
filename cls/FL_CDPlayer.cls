VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_CDPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************
'*** Flamed Audio CD Player class
'***
'*** plays audio track either
'*** analog or digital
'*******************************************

' MCI commands:
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/multimed/htm/_win32_multimedia_command_strings.asp

Public Enum FL_PlaybackState
    PBS_PLAYING
    PBS_PAUSING
    PBS_STOPPED
End Enum
Private udePlayState    As FL_PlaybackState

Private Const MCIAlias  As String = "FLAMEDv4"

Private Type Track
    num     As Byte
    Start   As FL_MSF
    end     As FL_MSF
    Length  As FL_MSF
End Type

Private Type Tracks
    Track(98)   As Track
    count       As Byte
End Type
' track collection
Private udtTracks           As Tracks

' digital or analog playback?
Private blnDigital          As Boolean

' Current track
Private intTrack            As Integer

' current drive ID
Private strDrvID            As String

' current position
Private lngPos              As Long

' MSF<>LBA conversion
Private MSF                 As New FL_MSF

' End Of Track Timer
Private WithEvents cTmr     As clsTimer
Attribute cTmr.VB_VarHelpID = -1

' Safe Subclasser
Private WithEvents subcls   As lvSubclasser
Attribute subcls.VB_VarHelpID = -1

' Mixer control
Private cMixer              As New clsMixer

' Digital volume
Private dblVol                  As Double

' Subclassed window
Private lngHWND                 As Long

' Wave Out Handle
Private hWaveOut                As Long

' Blocks per read
Private Const Blocks            As Long = 10
' number of buffers
Private Const BUFFERS           As Long = 5

' buffer
Private Type wavbuffer
    hdr                         As WAVEHDR
    btArr(2352 * Blocks - 1)    As Byte
End Type

' buffer collection
Private btBuffers(BUFFERS)      As wavbuffer

' Red Book audio samples
' raised in digital mode
Public Event samples(Data() As Integer)
' raised when Play, Pause or Stop is called
Public Event StateChanged(ByVal State As FL_PlaybackState)
' raised every second when playing
Public Event Timer(ByVal pos As Long)
' raised when the track gets changed
Public Event TrackChanged(ByVal Track As Integer)
' raised when reaching the end of a track
Public Event TrackEndReached(ByVal Track As Integer, stopplay As Boolean)

Private Function InitWaveOut() As Boolean

    ' create a new window
    lngHWND = CreateWindowEx( _
                0&, "STATIC", _
                "fl_WaveOut", _
                0&, 0&, 0&, _
                0&, 0&, 0&, _
                0&, App.hInstance, _
                0&)

    ' got a hWnd?
    If lngHWND <> 0 Then
        ' subclass the window
        subcls.SubclassMe lngHWND
    Else
        Exit Function
    End If

    ' open WaveOut
    If Not OpenDevice(lngHWND) Then
        DeInitWaveOut
    Else
        InitWaveOut = True
    End If

End Function

Private Function DeInitWaveOut() As Boolean
    ' close wave out
    CloseDevice

    ' unsubclass the window
    subcls.UnSubclassMe lngHWND
    ' and destroy it
    DestroyWindow lngHWND
End Function

Private Function OpenDevice(ByVal hWnd As Long) As Boolean

    Dim wfx     As WAVEFORMAT

    ' red book audio header
    wfx.cbSize = 0
    wfx.nAvgBytesPerSec = 44100 * 2 * 2
    wfx.nBlockAlign = 4
    wfx.nChannels = 2
    wfx.nSamplesPerSec = 44100
    wfx.wBitsPerSample = 16
    wfx.wFormatTag = 1

    ' open the waveout for red book audio
    OpenDevice = MMSYSERR_NOERROR = waveOutOpen(hWaveOut, 0, wfx, hWnd, 0, CALLBACK_WINDOW)

End Function

Private Sub CloseDevice()
    waveOutClose hWaveOut
End Sub

Public Property Let Volume(aval As Integer)
    If blnDigital Then
        dblVol = CDbl(aval) / 100
    Else
        cMixer.Value = aval
    End If
End Property

Public Property Get Volume() As Integer
    If blnDigital Then
        Volume = dblVol * 100
    Else
        Volume = cMixer.Value
    End If
End Property

Public Property Get State() As FL_PlaybackState
    State = udePlayState
End Property

Private Property Let State(aval As FL_PlaybackState)
    udePlayState = aval
    If aval = PBS_PAUSING Then cTmr.Enabled = False
    If aval = PBS_PLAYING Then cTmr.Enabled = True
    If aval = PBS_STOPPED Then cTmr.Enabled = False
    RaiseEvent StateChanged(udePlayState)
End Property

Public Property Let DigitalMode(aval As Boolean)
    blnDigital = aval

    If blnDigital Then
        If Not InitWaveOut Then
            MsgBox "Could not open WaveOut!", vbExclamation, "Error"
        End If
    Else
        DeInitWaveOut
    End If
End Property

Public Property Get DigitalMode() As Boolean
    DigitalMode = blnDigital
End Property

Public Function OpenDrive(DrvID As String) As Boolean

    Dim cSessInfo       As New FL_SessionInfo
    Dim cTrackInfo      As New FL_TrackInfo
    Dim i               As Integer

    StopTrack
    mciSendString "close " & MCIAlias, vbNullString, 0, 0

    ' get tracks of the first session
    If Not cSessInfo.GetInfo(DrvID, 1) Then
        Exit Function
    End If

    ' get all tracks of the first session
    For i = 1 To cSessInfo.LastTrack

        If Not cTrackInfo.GetInfo(DrvID, i) Then
            Exit Function
        End If

        ' save track data
        With udtTracks.Track(i)

            ' reached the first data track?
            If cTrackInfo.mode <> FL_TrackModes.MODE_AUDIO Then
                'ReDim Preserve udtTracks.Track(i - 1) As Track
                Exit For
            End If

            .num = i
            Set .Start = New FL_MSF
            Set .end = New FL_MSF
            Set .Length = New FL_MSF
            .Start.LBA = cTrackInfo.TrackStart.LBA
            .end.LBA = cTrackInfo.TrackEnd.LBA
            .Length.LBA = cTrackInfo.TrackLength.LBA

        End With

        udtTracks.count = i

    Next

    If blnDigital Then
        strDrvID = DrvID
    Else
        strDrvID = cd.DrvChrFromDrvID(DrvID)
        mciSendString "open " & strDrvID & ": type cdaudio alias " & MCIAlias, vbNullString, 0&, 0&
        mciSendString "set " & MCIAlias & " time format msf", vbNullString, 0, 0
    End If

    OpenDrive = True

End Function

Public Property Get TrackCount() As Integer
    TrackCount = udtTracks.count
End Property

Public Property Get TrackLength(Track As Integer) As FL_MSF
    With udtTracks.Track(Track)
        Set TrackLength = .Length
    End With
End Property

Public Property Get TrackEnd(Track As Integer) As FL_MSF
    Set TrackEnd = udtTracks.Track(Track).end
End Property

Public Property Get TrackStart(Track As Integer) As FL_MSF
    Set TrackStart = udtTracks.Track(Track).Start
End Property

Public Property Get CurrentTrack() As Integer
    CurrentTrack = intTrack
End Property

Public Property Get CurrentPos() As FL_MSF
    Dim strBuffer As String * 255

    If blnDigital Then
        MSF.LBA = lngPos - udtTracks.Track(intTrack).Start.LBA
        Set CurrentPos = MSF
    Else
        mciSendString "status " & MCIAlias & " position", strBuffer, Len(strBuffer), 0
        MSF.MSF = Left$(strBuffer, InStr(strBuffer, Chr$(0)) - 1)
        MSF.LBA = MSF.LBA - udtTracks.Track(intTrack).Start.LBA
        Set CurrentPos = MSF
    End If
End Property

Public Sub SeekTrack(pos As Long)

    If udePlayState = PBS_STOPPED Then Exit Sub

    If blnDigital Then
        lngPos = pos + udtTracks.Track(intTrack).Start.LBA
    Else
        MSF.LBA = udtTracks.Track(intTrack).Start.LBA + pos
        mciSendString "stop " & MCIAlias, vbNullString, 0, 0
        mciSendString "play " & MCIAlias & " from " & MSF.MSF & " to " & TrackEnd(intTrack).MSF, vbNullString, 0, 0
    End If

End Sub

'###########################################

Public Sub NextTrack()
    If CurrentTrack + 1 > TrackCount Then
        PlayTrack 1
    Else
        PlayTrack CurrentTrack + 1
    End If
End Sub

Public Sub PrevTrack()
    If CurrentTrack - 1 < 1 Then
        PlayTrack TrackCount
    Else
        PlayTrack CurrentTrack - 1
    End If
End Sub

Public Sub PlayTrack(Track As Integer)

    StopTrack

    intTrack = Track

    RaiseEvent TrackChanged(intTrack)
    State = PBS_PLAYING

    If blnDigital Then
        PlayTrackDigital
    Else
        PlayTrackAnalog
    End If

End Sub

Private Sub PlayTrackAnalog()

    With udtTracks.Track(intTrack)
        mciSendString "play " & MCIAlias & " from " & .Start.MSF & " to " & .end.MSF, vbNullString, 0, 0
    End With

End Sub

Private Sub PlayTrackDigital()

    Dim i   As Integer

    DeInitWaveOut
    InitWaveOut

    ' set the read speed to 4x
    ' to slow the drive down
    ' and reduce noise
    CDRomSetCDSpeed strDrvID, 4 * 177, 0, False

    ' prepare the buffers
    For i = 0 To UBound(btBuffers)

        ' prepare buffers
        btBuffers(i).hdr.lpData = VarPtr(btBuffers(i).btArr(0))
        btBuffers(i).hdr.dwBufferLength = UBound(btBuffers(i).btArr) + 1
        btBuffers(i).hdr.dwUser = i

        waveOutPrepareHeader hWaveOut, btBuffers(i).hdr, Len(btBuffers(i).hdr)

    Next

    lngPos = udtTracks.Track(intTrack).Start.LBA - 150

    ' fill the just prepared buffers
    ' to write them to WaveOut
    For i = 0 To BUFFERS
        SendMessageA lngHWND, MM_WOM_DONE, 0, ByVal VarPtr(btBuffers(i).hdr)
    Next

End Sub

'###########################################

Public Sub PauseResumeTrack()
    If blnDigital Then
        PauseResumeTrackDigital
    Else
        PauseResumeTrackAnalog
    End If
End Sub

Private Sub PauseResumeTrackDigital()
    If udePlayState = PBS_PAUSING Then
        waveOutRestart hWaveOut
        State = PBS_PLAYING
    ElseIf udePlayState = PBS_PLAYING Then
        waveOutPause hWaveOut
        State = PBS_PAUSING
    End If
End Sub

Private Sub PauseResumeTrackAnalog()
    If udePlayState = PBS_PAUSING Then
        mciSendString "resume " & MCIAlias, vbNullString, 0, 0
        State = PBS_PLAYING
    ElseIf udePlayState = PBS_PLAYING Then
        mciSendString "pause " & MCIAlias, vbNullString, 0, 0
        State = PBS_PAUSING
    End If
End Sub

'###########################################

Public Sub StopTrack()
    udePlayState = PBS_STOPPED
    If blnDigital Then
        StopTrackDigital
    Else
        StopTrackAnalog
    End If
    State = PBS_STOPPED
End Sub

Private Sub StopTrackAnalog()
    mciSendString "stop " & MCIAlias, vbNullString, 0, 0
End Sub

Private Sub StopTrackDigital()

    Dim i   As Integer
    Dim j   As Integer

    ' reset the buffers associated
    ' with our wave out handle
    waveOutReset hWaveOut

    ' del buffers
    For i = 0 To UBound(btBuffers)
        Do While 0 <> waveOutUnprepareHeader(hWaveOut, btBuffers(i).hdr, Len(btBuffers(i).hdr))
            If j = 10 Then Exit Do
            Sleep 50
            j = j + 1
        Loop
    Next

    ' reset the LBA
    lngPos = 0

End Sub

Private Sub Class_Initialize()

    ' timer event every 1000 ms
    Set cTmr = New clsTimer
    cTmr.Interval = 1000

    ' new subclasser
    Set subcls = New lvSubclasser

    ' we want to change the vol of CD in
    cMixer.Choose CompactDiscIn, Volume

    ' digital volume (100)
    dblVol = 1

End Sub

Private Sub Class_Terminate()
    StopTrack
    DeInitWaveOut
End Sub

Private Sub cTmr_Timer()
    Dim blnStop As Boolean

    ' end of track reached?
    If CurrentPos.LBA > TrackLength(intTrack).LBA Then

        ' yep, alert
        RaiseEvent TrackEndReached(intTrack, blnStop)

        ' stop track?
        If blnStop Then
            StopTrack
        Else
            NextTrack
        End If

    Else

        RaiseEvent Timer(CurrentPos.LBA)

    End If
End Sub

Private Sub FillBuffer(index As Integer)

    On Error Resume Next

    Dim intSamples()    As Integer

    If udePlayState = PBS_STOPPED Then Exit Sub

    ' refill the played buffer
    CDRomReadCD strDrvID, lngPos, Blocks, VarPtr(btBuffers(index).btArr(0)), UBound(btBuffers(index).btArr), RCD_RAW

    ' increase CD position
    lngPos = lngPos + Blocks

    ' move byte buffer to integer buffer
    ReDim intSamples(UBound(btBuffers(index).btArr) / 2 - 1) As Integer
    CopyMemory intSamples(0), btBuffers(index).btArr(0), UBound(btBuffers(index).btArr) + 1

    ' give em to the main app
    RaiseEvent samples(intSamples)

    ' set volume
    NormalizeChunk btBuffers(index).btArr, dblVol

    ' play the buffer
    waveOutWrite hWaveOut, btBuffers(index).hdr, Len(btBuffers(index).hdr)

End Sub

'for 16 bit wave samples
Private Sub NormalizeChunk(samples() As Byte, dblnorm As Double)

    On Error Resume Next

    Dim intSamples() As Integer
    Dim i            As Long

    ' need to normalise?
    If dblnorm > 0.99 Then Exit Sub

    ' copy the byte buffer to an integer buffer
    ' for easier handling of the data
    ReDim intSamples((UBound(samples) + 1) \ 2) As Integer
    CopyMemory intSamples(0), samples(0), UBound(samples) + 1

    For i = 0 To UBound(intSamples)

        '   no overflows
        If CDbl(intSamples(i)) * dblnorm > 32767 Then
            intSamples(i) = 32767

        ElseIf CDbl(intSamples(i)) * dblnorm < -32767 Then
            intSamples(i) = -32767

        Else
            'normalise the current sample
            intSamples(i) = intSamples(i) * dblnorm + 0.5
        End If

    Next

    ' copy the integer buffer to the byte buffer
    CopyMemory samples(0), intSamples(0), UBound(intSamples) * 2

End Sub

Private Sub subcls_ProcessMessage(ByVal hWnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Long, bOverriden As Boolean, lOverrideReturn As Long)

    Dim hdr         As WAVEHDR

    ' sound buffer empty?
    If wMsg = MM_WOM_DONE Then

        ' copy the just played wave header
        CopyMemory hdr, ByVal lParam, Len(hdr)

        ' finished playing the track?
        Select Case udePlayState
            Case PBS_PAUSING:
            Case PBS_STOPPED:
            Case PBS_PLAYING:
                FillBuffer CInt(hdr.dwUser)
        End Select

    End If

End Sub
