VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_CDDataWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************
'*** Flamed Data CD Writer
'***
'*** writes data disks (on the fly)
'*******************************************

' ToDo: Cancel option

Private WithEvents iso  As FL_ISO9660Writer
Attribute iso.VB_VarHelpID = -1

Private Type Buffer
    ol                      As overlapped
    Buffer()                As Byte
    read                    As Long
    full                    As Boolean
End Type

Private strFile             As String

Private fh                  As Long
Private lngFileOffset       As Long

Private lngBufSize          As Long
Private lngFileSize         As Long

Private lngEstimatedSize    As Long
Private lngNWA              As Long

Private blnNextSessAllowed  As Boolean
Private blnCloseSession     As Boolean
Private blnTestMode         As Boolean
Private blnBURNProof        As Boolean
Private blnEjectAfterWrite  As Boolean
Private blnOFT              As Boolean

Private blnIsNullFile       As Boolean
Private blnEOF              As Boolean

Private strTemp             As String

Public Event StartWriting()
Public Event WriteProgress(ByVal Percent As Integer)
Public Event Finished()
Public Event ClosingSession()

Public Event CheckForFiles()
Public Event FilesMissing(strFiles() As String)

Public Event ISOProgress(ByVal lngMax As Long, _
                         ByVal lngValue As Long)

'Public Property Get CloseSession() As Boolean
'    CloseSession = blnCloseSession
'End Property
'
'Public Property Let CloseSession(aval As Boolean)
'    blnCloseSession = aval
'End Property

Public Function CloseLastSession(DrvID As String, finalize As Boolean) As Boolean

    Dim cParams     As New FL_WriteParams

    If Not cParams.GetData(DrvID) Then
        Exit Function
    End If

    cParams.Multisession = Abs(Not finalize)
    ' We write data
    cParams.SessionFormat = SF_CDDA_DATA
    ' 2048 bytes/sector
    cParams.DataBlockType = DB_MODE1_ISO
    ' Track-At-Once writing
    cParams.WriteType = WT_TAO
    ' Data uninterrupted
    cParams.TrackMode = 4

    cParams.Copy = False
    cParams.FixedPacket = False
    cParams.LinkSizeValid = False

    ' 2 seconds pause between tracks
    cParams.AudioPauseLength = 150
    cParams.ApplicationCode = 0
    cParams.LinkSize = 0

    If Not cParams.SendData(DrvID) Then
        Exit Function
    End If

    CloseLastSession = CDRomCloseCD(DrvID, 2, 0)

End Function

Public Property Let TempDir(aval As String)
    strTemp = aval
    If Not Right$(strTemp, 1) = "\" Then
        strTemp = strTemp & "\"
    End If
End Property

Public Property Get TempDir() As String
    TempDir = strTemp
End Property

Public Property Get EjectAfterWrite() As Boolean
    EjectAfterWrite = blnEjectAfterWrite
End Property

Public Property Let EjectAfterWrite(aval As Boolean)
    blnEjectAfterWrite = aval
End Property

Public Property Get NextSessionAllowed() As Boolean
    NextSessionAllowed = blnNextSessAllowed
End Property

Public Property Let NextSessionAllowed(aval As Boolean)
    blnNextSessAllowed = aval
End Property

Public Property Get TestMode() As Boolean
    TestMode = blnTestMode
End Property

Public Property Let TestMode(aval As Boolean)
    blnTestMode = aval
End Property

Private Property Get BURNProof() As Boolean
    BURNProof = blnBURNProof
End Property

Private Property Let BURNProof(aval As Boolean)
    blnBURNProof = aval
End Property

Public Property Get OnTheFly() As Boolean
    OnTheFly = blnOFT
End Property

Public Property Let OnTheFly(aval As Boolean)
    blnOFT = aval
End Property

Private Sub Class_Initialize()
    ' set buffer size to 4 Mibibyte
    BufferSize = 4& * 1024& ^ 2&

    TestMode = False
    NextSessionAllowed = False
    EjectAfterWrite = True

    strTemp = GetTempDir()

End Sub

Private Sub iso_Progress(ByVal lngMax As Long, ByVal lngValue As Long)
    RaiseEvent ISOProgress(lngMax, lngValue)
End Sub

Public Property Set ISOClass(cls As FL_ISO9660Writer)
    Set iso = cls
End Property

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
Private Function OpenFile(file As String) As Boolean

    On Error GoTo ErrorHandler

    lngFileOffset = 0
    lngFileSize = FileLen(file)

    ' write at least 1 block per empty file
    If lngFileSize = 0 Then
        blnIsNullFile = True
        lngFileSize = 2048
    Else
        blnIsNullFile = False
    End If

    ' round the file size up
    If CDbl(lngFileSize) / 2048# <> lngFileSize \ 2048& Then
        lngFileSize = lngFileSize + (2048 - (lngFileSize Mod 2048))
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

    If blnIsNullFile Then buf.read = 2048

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

    If lng / 2048 <> lng \ 2048 Then
        lng = lng + (2048& - (lng Mod 2048&))
    End If

    RoundUp = lng

End Function

Private Function ProcessBuffer(DrvID As String, buf As Buffer, LBA As Long) As Boolean

    Dim startLBA    As Long
    Dim dblLastLen  As Double

    startLBA = LBA

    For LBA = startLBA To LBA + (lngBufSize \ 2048&) Step 20

        ' next write will go over the bounds?
        If (LBA * 2048&) + (20& * 2048&) > (startLBA * 2048&) + buf.read Then

            ' round last write length up
            ' invalid buffer length will cause app to freeze
            dblLastLen = (buf.read - ((LBA - startLBA) * 2048&))
            If dblLastLen / 2048 <> dblLastLen \ 2048 Then
                dblLastLen = dblLastLen + (2048& - (dblLastLen Mod 2048&))
            End If
            dblLastLen = dblLastLen - 1&

            ' write the last sectors
            If Not CDRomBurnCD(DrvID, LBA, startLBA - (((LBA * 2048&) - buf.read) \ 2048&), VarPtr(buf.Buffer((LBA - startLBA) * 2048&)), dblLastLen) Then
                Exit Function
            End If

            LBA = LBA + startLBA - (((LBA * 2048&) - buf.read) \ 2048&)

            Exit For

        Else

            If Not CDRomBurnCD(DrvID, LBA, 20&, VarPtr(buf.Buffer((LBA - startLBA) * 2048&)), 20& * 2048&) Then
                Exit Function
            End If

        End If

        RaiseEvent WriteProgress((LBA - lngNWA) / lngEstimatedSize * 100&)

        DoEvents

    Next

    ' close the ReadFile() event
    CloseHandle buf.ol.hEvent
    ' buffer processed
    buf.full = False

    ProcessBuffer = True

End Function

' >>>> TOO BIG
' >>>> Should divide this function into
' >>>> little pieces
Public Function WriteISOtoCD(DrvID As String) As FL_BURNRET

    '   Write Parameters Mode Page Class
    Dim cWP                 As New FL_WriteParams
    '   CD Information Class
    Dim cCDInfo             As New FL_CDInfo
    '   Drive Information Class
    Dim cDrvInfo            As New FL_DriveInfo

    Dim udtTrack            As t_RTI

    '   local files list for OTF writing
    Dim strFiles()          As String
    Dim strMissing()        As String

    '   buffers
    Dim b1      As Buffer, b2   As Buffer
    '   LBA
    Dim i       As Long, j      As Long

    '   EOF?
    Dim blnDoEnd            As Boolean

    '

    ' get information about media
    If Not cCDInfo.GetInfo(DrvID) Then
        WriteISOtoCD = BURNRET_INVALID_MEDIA
        Exit Function
    End If

    ' check for valid media (CD-R/RW needed)
    If cCDInfo.MediaType <> e_CDType.ROMTYPE_CDR And _
       cCDInfo.MediaType <> e_CDType.ROMTYPE_CDRW Then
            WriteISOtoCD = BURNRET_INVALID_MEDIA
            Exit Function
    End If

    '
    If cCDInfo.MediaStatus = FL_Status.STAT_COMPLETE Then
        WriteISOtoCD = BURNRET_NOT_EMPTY
        Exit Function
    End If

    ' drive has BURN-Proof feature?
    cDrvInfo.GetInfo DrvID
    blnBURNProof = CBool(cDrvInfo.WriteCapabilities And WC_BURNPROOF)

    ' read page
    If Not cWP.GetData(DrvID) Then
        WriteISOtoCD = BURNRET_WPMP
        Exit Function
    End If

    ' get data from page
    cWP.ExtractData

    ' always activate BURN-Proof if available
    cWP.BURNProof = blnBURNProof
    ' Test Mode (no physical burning)
    cWP.TestMode = blnTestMode

    ' "B0 pointer
    '  00b - No B0 pointer
    '  01b - Next session allowed
    '  10b - Reserved
    '  11b - Next session allowed
    '        B0 pointer = next program area"
    cWP.Multisession = Abs(blnNextSessAllowed) * 3
    ' We write data
    cWP.SessionFormat = SF_CDDA_DATA
    ' 2048 bytes/sector
    cWP.DataBlockType = DB_MODE1_ISO
    ' Track-At-Once writing
    cWP.WriteType = WT_TAO
    ' Data uninterrupted
    cWP.TrackMode = 4

    cWP.Copy = False
    cWP.FixedPacket = False
    cWP.LinkSizeValid = False

    ' 2 seconds pause between tracks
    cWP.AudioPauseLength = 150
    cWP.ApplicationCode = 0
    cWP.LinkSize = 0

    ' send page
    If Not cWP.SendData(DrvID) Then
        WriteISOtoCD = BURNRET_WPMP
        Exit Function
    End If

    RaiseEvent CheckForFiles

    ' get the local pos of files in the image
    strFiles = iso.GetLocalFiles()
    ReDim strMissing(0) As String

    lngEstimatedSize = 0

    ' check if they still exist
    For i = LBound(strFiles) + 1 To UBound(strFiles)
        If Not FileExists(strFiles(i)) Then
            strMissing(UBound(strMissing)) = strFiles(i)
            ReDim Preserve strMissing(UBound(strMissing) + 1) As String
        End If
        lngEstimatedSize = lngEstimatedSize + RoundUp(FileLen(strFiles(i)))
    Next

    ' files missing?
    If strMissing(0) <> vbNullString Then
        RaiseEvent FilesMissing(strMissing)
        WriteISOtoCD = BURNRET_FILE_ACCESS
        Exit Function
    End If

    ' Either write only ISO header
    ' or full image to disk
    If Not iso.CreateISO(strTemp & "temp.iso", blnOFT) Then
        WriteISOtoCD = BURNRET_ISOCREATION
        Exit Function
    End If

    ' ISO header (blnOFT) or ISO image (Not blnOFT)
    strFiles(0) = strTemp & "temp.iso"

    ' not on-the-fly?
    If Not blnOFT Then
        ' we have only 1 file to write
        ReDim Preserve strFiles(0) As String
    End If

    lngEstimatedSize = (lngEstimatedSize + FileLen(strFiles(0)))
    lngEstimatedSize = lngEstimatedSize \ 2048

    ' get next writable address
    If Not CDRomReadTrackInformation(DrvID, 1, cCDInfo.Tracks, VarPtr(udtTrack), Len(udtTrack) - 1) Then
        WriteISOtoCD = BURNRET_NO_NEXT_WRITABLE_LBA
        Exit Function
    End If

    lngNWA = cd.LShift(udtTrack.Track_Next_Writable(0), 24) Or _
             cd.LShift(udtTrack.Track_Next_Writable(1), 16) Or _
             cd.LShift(udtTrack.Track_Next_Writable(2), 8) Or _
             udtTrack.Track_Next_Writable(3)

    RaiseEvent StartWriting

    '
    WriteISOtoCD = BURNRET_OK

    i = lngNWA

    ' Write all files
    For j = LBound(strFiles) To UBound(strFiles)

        ' open the current file
        If Not OpenFile(strFiles(j)) Then
            WriteISOtoCD = BURNRET_FILE_ACCESS
            Exit Function
        End If
    
        ' prepare and fill first buffer
        b1 = PrepareBuffer
        ReadBuffer b1
        WaitForBuffer b1

        ' end of file reached?
        If EndOfFile Then blnDoEnd = True

        ' do while not at the end of the file
        Do

            ' first buffer is full?
            If b1.full Then
                'yep, burn it

                ' end of file reached?
                If blnDoEnd Then
                    ProcessBuffer DrvID, b1, i
                    Exit Do
                End If

                ' fill the second buffer while writing
                b2 = PrepareBuffer
                ReadBuffer b2

                ' burn the first buffer
                If Not ProcessBuffer(DrvID, b1, i) Then
                    WriteISOtoCD = BURNRET_WRITE
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
                    ProcessBuffer DrvID, b2, i
                    Exit Do
                End If

                ' fill the first buffer while burning
                b1 = PrepareBuffer
                ReadBuffer b1

                ' write second buffer to disk
                If Not ProcessBuffer(DrvID, b2, i) Then
                    WriteISOtoCD = BURNRET_WRITE
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
        blnDoEnd = False
        blnEOF = False

    Next

    ' hehe :)
    RaiseEvent WriteProgress(100)

    ' write drive's cache to disk
    CDRomSyncCache DrvID

    ' long write in progress?
    Do While CheckUnitReady(DrvID) = 2
        Sleep 2000
        DoEvents
    Loop

    ' CDRomCloseCD() will fail for TestMode = True
    If Not TestMode Then

        RaiseEvent ClosingSession

        ' Close track
        If Not CDRomCloseCD(DrvID, 1, 1) Then
            WriteISOtoCD = BURNRET_CLOSE_TRACK
            Exit Function
        End If

        ' Close session
        'If blnCloseSession Then
            If Not CDRomCloseCD(DrvID, 2, 0) Then
                WriteISOtoCD = BURNRET_CLOSE_SESSION
                Exit Function
            End If
        'End If

    End If

    ' delete temporary filesystem
    If blnOFT Then DelEx strFiles(0)

    RaiseEvent Finished

    ' you never know...
    CDRomUnlockMedia DrvID

    ' eject disk?
    If blnEjectAfterWrite Then CDRomUnloadTray DrvID

End Function

' from CDR-DAO
' check if unit is really ready
Private Function CheckUnitReady(DrvID As String) As Long

    ' quick'n'dirty
    Dim cmd(5)  As Byte, cmd2(9)    As Byte
    Dim udtRDI  As t_RDI

    If cd.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0) Then

        cmd2(0) = &H51
        cmd2(8) = 4

        If Not cd.ExecCMD(DrvID, cmd2, 10, False, SRB_DIR_IN, VarPtr(udtRDI), Len(udtRDI) - 1) Then

            ' long write in progress?
            If cd.LastSK = 2 And cd.LastASC = 12 Then
                CheckUnitReady = 2
            Else
                CheckUnitReady = 0
            End If

        Else

            CheckUnitReady = 1

        End If

    Else

        CheckUnitReady = 0

    End If

End Function

Private Sub DelEx(file As String)
    On Error Resume Next
    Kill file
End Sub
