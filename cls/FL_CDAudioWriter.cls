VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_CDAudioWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************
'*** Flamed Audio CD writer
'***
'*** writes MP3 and WAV to CD-R(W)
'*******************************************

Private Type FL_AUDIOPROJECTHEADER
    ID                      As String * 7
    count                   As Integer
End Type

Private Type Buffer
    ol                      As overlapped
    Buffer()                As Byte
    read                    As Long
    full                    As Boolean
End Type

Private fh                  As Long
Private lngFileOffset       As Long

Private lngBufSize          As Long
Private lngFileSize         As Long

Private lngEstimatedSize    As Long
Private lngNWA              As Long

Private blnBURNProof        As Boolean

Private blnIsNullFile       As Boolean
Private blnEOF              As Boolean

Private strFiles()          As String
Private strTemp             As String

Private blnEjectAfterWrite  As Boolean

Public Event StartCaching()
Public Event StartWriting()
Public Event CacheProgress(ByVal Percent As Integer, ByVal Track As Integer)
Public Event WriteProgress(ByVal Percent As Integer, ByVal Track As Integer)
Public Event Finished()
Public Event ClosingTrack(ByVal Track As Integer)
Public Event ClosingSession()

Public Property Get EjectAfterWrite() As Boolean
    EjectAfterWrite = blnEjectAfterWrite
End Property

Public Property Let EjectAfterWrite(aval As Boolean)
    blnEjectAfterWrite = aval
End Property

Private Function SetWriteParams(DrvID As String) As Boolean

    Dim cWP     As New FL_WriteParams

    ' read page
    If Not cWP.GetData(DrvID) Then
        Exit Function
    End If

    ' get data from page
    cWP.ExtractData

    ' always activate BURN-Proof if available
    cWP.BURNProof = blnBURNProof
    ' Test Mode (no physical burning)
    cWP.TestMode = False

    ' "B0 pointer
    '  00b - No B0 pointer
    '  01b - Next session allowed
    '  10b - Reserved
    '  11b - Next session allowed
    '        B0 pointer = next program area"
    cWP.Multisession = 0
    ' We write data
    cWP.SessionFormat = SF_CDDA_DATA
    ' 2352 bytes/sector
    cWP.DataBlockType = DB_RAW_2352
    ' Track-At-Once writing
    cWP.WriteType = WT_TAO
    ' 2 channel audio
    cWP.TrackMode = 0

    cWP.Copy = False
    cWP.FixedPacket = False
    cWP.LinkSizeValid = False

    ' 2 seconds pause between tracks
    cWP.AudioPauseLength = 150
    cWP.ApplicationCode = 0
    cWP.LinkSize = 0

    ' send page
    SetWriteParams = cWP.SendData(DrvID)

End Function

Public Function WriteAudioToCD(ByVal DrvID As String) As FL_BURNRET

    '   CD Information Class
    Dim cCDInfo             As New FL_CDInfo
    '   Drive Information Class
    Dim cDrvInfo            As New FL_DriveInfo
    '   track information
    Dim udtTrkNfo           As t_RTI

    '   buffers
    Dim b1      As Buffer, b2   As Buffer
    '   LBA
    Dim i       As Long, j      As Long

    '   EOF?
    Dim blnDoEnd            As Boolean


    ' get information about media
    If Not cCDInfo.GetInfo(DrvID) Then
        WriteAudioToCD = BURNRET_INVALID_MEDIA
        Exit Function
    End If

    ' check for valid media (CD-R/RW needed)
    If cCDInfo.MediaType <> e_CDType.ROMTYPE_CDR And _
       cCDInfo.MediaType <> e_CDType.ROMTYPE_CDRW Then
            WriteAudioToCD = BURNRET_INVALID_MEDIA
            Exit Function
    End If

    '
    If cCDInfo.MediaStatus <> FL_Status.STAT_EMPTY Then
        WriteAudioToCD = BURNRET_NOT_EMPTY
        Exit Function
    End If

    ' drive has BURN-Proof feature?
    cDrvInfo.GetInfo DrvID
    blnBURNProof = CBool(cDrvInfo.WriteCapabilities And WC_BURNPROOF)

    lngNWA = 0

    For j = 0 To FileCount - 1

        ' set track mode
        If Not SetWriteParams(DrvID) Then
            WriteAudioToCD = BURNRET_WPMP
            Exit Function
        End If

        ' get next writable address
        If Not CDRomReadTrackInformation(DrvID, 1, j + 1, VarPtr(udtTrkNfo), Len(udtTrkNfo) - 1) Then
            WriteAudioToCD = BURNRET_NO_NEXT_WRITABLE_LBA
            Exit Function
        End If

        lngNWA = cd.LShift(udtTrkNfo.Track_Next_Writable(0), 24) Or _
                 cd.LShift(udtTrkNfo.Track_Next_Writable(1), 16) Or _
                 cd.LShift(udtTrkNfo.Track_Next_Writable(2), 8) Or _
                 udtTrkNfo.Track_Next_Writable(3)

        i = lngNWA

        RaiseEvent StartCaching

        ' decode file
        If Not CacheFileToHDD(strFiles(j), j + 1) Then
            WriteAudioToCD = BURNRET_FILE_ACCESS
            Exit Function
        End If

        RaiseEvent StartWriting

        ' open the current file
        If Not OpenFile(strTemp & FileFromPathFile(strFiles(j)) & ".tmp") Then
            WriteAudioToCD = BURNRET_FILE_ACCESS
            Exit Function
        End If

        ' prepare and fill first buffer
        b1 = PrepareBuffer
        ReadBuffer b1
        WaitForBuffer b1

        ' do while not at the end of the file
        Do

            ' first buffer is full?
            If b1.full Then
                'yep, burn it

                ' end of file reached?
                If blnDoEnd Then
                    ProcessBuffer DrvID, b1, i, j + 1
                    Exit Do
                End If

                ' fill the second buffer while writing
                b2 = PrepareBuffer
                ReadBuffer b2

                ' burn the first buffer
                If Not ProcessBuffer(DrvID, b1, i, j + 1) Then
                    WriteAudioToCD = BURNRET_WRITE
                    CDRomSyncCache DrvID
                    CDRomUnlockMedia DrvID
                    If blnEjectAfterWrite Then CDRomUnloadTray DrvID
                    CloseFile
                    Exit Function
                End If

                ' end of file reached?
                If EndOfFile Then blnDoEnd = True

                ' wait for second buffer to finish
                WaitForBuffer b2

            ' second buffer is full?
            ElseIf b2.full Then
                ' yep

                ' end of file reached?
                If blnDoEnd Then
                    ProcessBuffer DrvID, b2, i, j + 1
                    Exit Do
                End If

                ' fill the first buffer while burning
                b1 = PrepareBuffer
                ReadBuffer b1

                ' write second buffer to disk
                If Not ProcessBuffer(DrvID, b2, i, j + 1) Then
                    WriteAudioToCD = BURNRET_WRITE
                    CDRomSyncCache DrvID
                    CDRomUnlockMedia DrvID
                    If blnEjectAfterWrite Then CDRomUnloadTray DrvID
                    CloseFile
                    Exit Function
                End If

                ' wait for first buffer to finish
                WaitForBuffer b1

            End If

        Loop

        ' Reset data
        CloseFile
        KillEx strTemp & FileFromPathFile(strFiles(j)) & ".tmp"
        blnDoEnd = False
        blnEOF = False

        ' synchronize cache
        CDRomSyncCache DrvID

        RaiseEvent ClosingTrack(j + 1)

        ' close track
        If Not CDRomCloseCD(DrvID, 1, j + 1) Then
            WriteAudioToCD = BURNRET_CLOSE_TRACK
            Exit Function
        End If

    Next

    RaiseEvent ClosingSession

    ' close session/disk
    If Not CDRomCloseCD(DrvID, 2, 0) Then
        WriteAudioToCD = BURNRET_CLOSE_SESSION
        Exit Function
    End If

    RaiseEvent Finished

    If blnEjectAfterWrite Then CDRomUnloadTray DrvID

    WriteAudioToCD = BURNRET_OK

End Function

Private Function CacheFileToHDD(ByVal strFile As String, ByVal Track As Integer) As Boolean

    On Error GoTo ErrorHandler

    Dim FF  As Integer
    Dim FF2 As Integer

    ' decoders
    Dim cDecWav     As New FL_DecoderWAV
    Dim cDecMP3     As New FL_DecoderMP3
    Dim cDecoder    As Object

    Dim lngFileLen  As Long

    Dim btBuffer()  As Byte
    ReDim btBuffer(4096& - 1&) As Byte

    ' decoded data buffer
    Dim btDecoded() As Byte
    ' result buffer
    Dim btMem()     As Byte
    ' result buffer length
    Dim lngMemLen   As Long

    ListCodecs

    Select Case LCase$(Right$(strFile, 3))

        Case "wav":
            ' set PCM wave input format
            With GetWavFormat(strFile)
                cDecWav.InputBlockAlign = .nBlockAlign
                cDecWav.InputBitsPerSample = .wBitsPerSample
                cDecWav.InputChannels = .nChannels
                cDecWav.InputSampleRate = .nSamplesPerSec
            End With

            Set cDecoder = cDecWav

        Case "mp3":

            With ReadMP3(strFile)
                cDecMP3.InputChannels = .Channels
                cDecMP3.InputSampleRate = .Frequency
            End With
            Set cDecoder = cDecMP3

    End Select

    If Not cDecoder.Init(btBuffer, UBound(btBuffer) + 1) Then
        Exit Function
    End If

    FF = FreeFile
    Open strFile For Binary As #FF

    ' get DATA chunk for wave files
    Select Case LCase$(Right$(strFile, 3))

        Case "wav"
            With GetWavChunkPos(strFile, "data")
                If .Length = 0 Then
                    lngFileLen = LOF(FF)
                Else
                    lngFileLen = .Length + 45
                End If
                Seek #FF, .Start + 1
            End With

        Case "mp3"
            lngFileLen = LOF(FF)

    End Select

    FF2 = FreeFile
    Open strTemp & FileFromPathFile(strFile) & ".tmp" For Binary As #FF2

        Do While Seek(FF) < lngFileLen

            If lngFileLen - Seek(FF) > 0 Then
                Get #FF, , btBuffer
                btDecoded = cDecoder.Decode(UBound(btBuffer) + 1, False)
            Else
                ReDim btBuffer(lngFileLen - Seek(FF) - 1) As Byte
                Get #FF, , btBuffer
                btDecoded = cDecoder.Decode(UBound(btBuffer) + 1, True)
            End If

            Put #FF2, , btDecoded

            RaiseEvent CacheProgress(Seek(FF) / lngFileLen * 100, Track)

            DoEvents

        Loop

        lngEstimatedSize = Seek(FF2)

    Close #FF, #FF2

    cDecoder.DeInit

    CacheFileToHDD = True

ErrorHandler:

End Function

Private Sub KillEx(ByVal strFile As String)
    On Error Resume Next: Kill strFile
End Sub

'#############################################
'# Async Read Buffer    ######################
'#############################################

' End of file reached?
Private Property Get EndOfFile() As Boolean
    EndOfFile = blnEOF
End Property

' set/get buffer size
Private Property Get BufferSize() As Long
    BufferSize = lngBufSize
End Property

Private Property Let BufferSize(aval As Long)
    lngBufSize = aval
End Property

' open a file for asnychronous reading
Private Function OpenFile(ByVal file As String) As Boolean

    On Error GoTo ErrorHandler

    lngFileOffset = 0
    lngFileSize = FileLen(file)

    ' write at least 1 block per empty file
    If lngFileSize = 0 Then
        blnIsNullFile = True
        lngFileSize = 2352
    Else
        blnIsNullFile = False
    End If

    ' round the file size up
    If CDbl(lngFileSize) / 2352# <> lngFileSize \ 2352& Then
        lngFileSize = lngFileSize + (2352 - (lngFileSize Mod 2352))
    End If

    ' open a file for overlapped reading
    fh = CreateFile(file, _
                    GENERIC_READ, _
                    FILE_SHARE_READ, _
                    ByVal 0&, _
                    OPEN_EXISTING, _
                    FILE_FLAG_OVERLAPPED, _
                    0)

    OpenFile = fh <> -1

    Exit Function

ErrorHandler:
    ' you never know...
    CloseHandle fh

End Function

' prepare a buffer for async reading
Private Function PrepareBuffer() As Buffer

    With PrepareBuffer.ol
        .hEvent = CreateEvent(ByVal 0&, 1, 0, "BUF")
        ResetEvent .hEvent
        .offset = lngFileOffset
    End With

    With PrepareBuffer
        ReDim .Buffer(lngBufSize - 1) As Byte
        .read = 0
    End With

End Function

' read a bunch of data to a buffer
Private Function ReadBuffer(buf As Buffer) As Boolean

    If blnEOF Then Exit Function

    ReadBuffer = ReadFile(fh, buf.Buffer(0), lngBufSize, _
                          buf.read, buf.ol)

    If blnIsNullFile Then buf.read = 2352

    If lngFileOffset + lngBufSize > lngFileSize Then
        blnEOF = True
    End If

    lngFileOffset = lngFileOffset + lngBufSize

End Function

' buffer is full?
Private Function IsBufferFull(buf As Buffer) As Boolean
    'IsBufferFull = GetOverlappedResult(fh, buf.ol, buf.read, False)
    IsBufferFull = buf.full
End Function

' wait for buffer to finish reading
Private Function WaitForBuffer(buf As Buffer) As Boolean
    If Not blnIsNullFile Then
        WaitForBuffer = GetOverlappedResult(fh, buf.ol, buf.read, True)
        buf.full = WaitForBuffer
    Else
        buf.full = True
    End If
End Function

' close the file
Private Function CloseFile()
    lngFileSize = 0
    lngFileOffset = 0
    CloseHandle fh
End Function

'###############################################
'###############################################
'###############################################

Private Function RoundUp(ByVal lng As Long) As Long

    If lng / 2352# <> lng \ 2352& Then
        lng = lng + (2352& - (lng Mod 2352&))
    End If

    RoundUp = lng

End Function

Private Function ProcessBuffer(DrvID As String, buf As Buffer, LBA As Long, Track As Integer) As Boolean

    Dim startLBA    As Long
    Dim dblLastLen  As Double

    startLBA = LBA

    For LBA = startLBA To LBA + (lngBufSize \ 2352&) Step 20

        ' next write will go over the bounds?
        If (LBA * 2352&) + (20& * 2352&) > (startLBA * 2352&) + buf.read Then

            ' round last write length up
            ' invalid buffer length will cause app to freeze
            dblLastLen = (buf.read - ((LBA - startLBA) * 2352&))
            If dblLastLen / 2352# <> dblLastLen \ 2352& Then
                dblLastLen = dblLastLen + (2352& - (dblLastLen Mod 2352&))
            End If
            dblLastLen = dblLastLen - 1&

            If dblLastLen < 1 Then
                ' close the ReadFile() event
                CloseHandle buf.ol.hEvent
                ' buffer processed
                buf.full = False

                ProcessBuffer = True
                Exit Function
            End If

            ' write the last sectors
            If Not CDRomBurnCD(DrvID, LBA, startLBA - (((LBA * 2352&) - buf.read) \ 2352&), VarPtr(buf.Buffer((LBA - startLBA) * 2352&)), dblLastLen) Then
                Exit Function
            End If

            LBA = LBA + startLBA - (((LBA * 2352&) - buf.read) \ 2352&)

            Exit For

        Else

            If Not CDRomBurnCD(DrvID, LBA, 20&, VarPtr(buf.Buffer((LBA - startLBA) * 2352&)), 20& * 2352&) Then
                Exit Function
            End If

        End If

        RaiseEvent WriteProgress(((LBA - lngNWA) * 2352) / lngEstimatedSize * 100&, Track)

        DoEvents

    Next

    ' close the ReadFile() event
    CloseHandle buf.ol.hEvent
    ' buffer processed
    buf.full = False

    ProcessBuffer = True

End Function
 
Public Property Get TempDir() As String
    TempDir = strTemp
End Property

Public Property Let TempDir(aval As String)
    strTemp = aval
    If Not Right$(strTemp, 1) = "\" Then
        strTemp = strTemp & "\"
    End If
End Property

Public Function LoadProject(ByVal strFile As String) As Boolean

    On Error GoTo ErrorHandler

    Dim FF  As Integer: FF = FreeFile
    Dim hdr As FL_AUDIOPROJECTHEADER

    Open strFile For Binary As #FF

        Get #FF, , hdr
        If hdr.ID <> "FLAUDIO" Then
            Close #FF
            Exit Function
        End If

        ReDim strFiles(hdr.count) As String
        Get #FF, , strFiles

    Close #FF

    LoadProject = True

ErrorHandler:

End Function

Public Function SaveProject(ByVal strFile As String) As Boolean

    On Error GoTo ErrorHandler

    Dim FF  As Integer: FF = FreeFile
    Dim hdr As FL_AUDIOPROJECTHEADER

    hdr.ID = "FLAUDIO"
    hdr.count = FileCount

    Open strFile For Binary As #FF
        Put #FF, , hdr
        Put #FF, , strFiles
    Close #FF

    SaveProject = True

ErrorHandler:

End Function

Public Function TrackLength(ByVal index As Integer) As Long

    Select Case LCase$(Right$(strFiles(index), 3))
        Case "wav": TrackLength = GetWAVLength(strFiles(index))
        Case "mp3": TrackLength = ReadMP3(strFiles(index)).Duration
    End Select

End Function

Public Sub MoveIndexDown(ByVal index As Integer)

    Dim strBuffer   As String

    If Not index = 0 Then
        strBuffer = strFiles(index - 1)
        strFiles(index - 1) = strFiles(index)
        strFiles(index) = strBuffer
    Else
        strBuffer = strFiles(FileCount - 1)
        strFiles(FileCount - 1) = strFiles(index)
        strFiles(index) = strBuffer
    End If

End Sub

Public Sub MoveIndexUp(ByVal index As Integer)

    Dim strBuffer   As String

    If Not index = FileCount - 1 Then
        strBuffer = strFiles(index + 1)
        strFiles(index + 1) = strFiles(index)
        strFiles(index) = strBuffer
    Else
        strBuffer = strFiles(0)
        strFiles(0) = strFiles(index)
        strFiles(index) = strBuffer
    End If

End Sub

Public Property Get file(ByVal index As Integer) As String
    file = strFiles(index)
End Property

Public Property Get FileCount() As Integer
    FileCount = UBound(strFiles)
End Property

Public Sub Clear()
    ReDim strFiles(0) As String
End Sub

Public Sub RemFile(ByVal index As Integer)

    Dim i   As Integer

    For i = index To FileCount - 1
        strFiles(i) = strFiles(i + 1)
    Next

    ReDim Preserve strFiles(UBound(strFiles) - 1) As String

End Sub

Public Function AddFile(ByVal strFile As String) As Boolean

    Select Case LCase$(Right$(strFile, 3))

        Case "mp3":

            ' do some validity checks
            With ReadMP3(strFile)
                If .Duration < 1 Then Exit Function
                If .Channels <> 2 Then Exit Function
            End With

        Case "wav":

            ' do some validity checks
            With GetWavFormat(strFile)

                If .wFormatTag <> WAVE_FORMAT_PCM Then Exit Function
                If .nChannels <> 2 Then Exit Function
                If .nSamplesPerSec <> 22050 Then
                    If .nSamplesPerSec <> 44100 Then
                        Exit Function
                    End If
                End If

            End With

        Case Else: Exit Function

    End Select

    strFiles(UBound(strFiles)) = strFile
    ReDim Preserve strFiles(UBound(strFiles) + 1) As String

    AddFile = True

End Function

Private Sub Class_Initialize()
    ReDim strFiles(0) As String
    BufferSize = 2352& * 1001&
    strTemp = GetTempDir
End Sub
