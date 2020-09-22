Attribute VB_Name = "modACM"
Option Explicit

' Audio Compression Manager API

Public Declare Function acmStreamPrepareHeader Lib "msacm32" (ByVal has As Long, _
        pash As ACMSTREAMHEADER, ByVal fdwPrepare As Long) As Long

Public Declare Function acmStreamUnprepareHeader Lib "msacm32" (ByVal has As Long, _
        pash As ACMSTREAMHEADER, ByVal fdwUnprepare As Long) As Long

Public Declare Function acmStreamOpen Lib "msacm32" (ByRef phas As Long, _
        ByVal had As Long, pwfxSrc As Any, pwfxDst As Any, ByVal pwfltr As Long, _
        ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long

Public Declare Function acmStreamSize Lib "msacm32" (ByVal has As Long, ByVal cbInput As Long, _
        pdwOutputBytes As Long, ByVal fdwSize As Long) As Long

Public Declare Function acmStreamConvert Lib "msacm32" (ByVal has As Long, _
        pash As ACMSTREAMHEADER, ByVal fdwConvert As Long) As Long

Public Declare Function acmStreamReset Lib "msacm32.dll" (ByVal has As Long, _
        ByVal fdwReset As Long) As Long

Public Declare Function acmStreamClose Lib "msacm32" (ByVal has As Long, _
        ByVal fdwClose As Long) As Long

Public Declare Function acmMetrics Lib "msacm32" (ByVal hao As Long, _
        ByVal uMetric As Integer, pMetric As Any) As Long

Public Declare Function acmDriverEnum Lib "msacm32" (ByVal fnCallback As Long, _
        dwInstance As Long, ByVal fdwEnum As Long) As Long

Public Declare Function acmDriverDetails Lib "msacm32" Alias "acmDriverDetailsA" ( _
        ByVal hadid As Long, padd As TACMDRIVERDETAILS, ByVal fdwDetails As Long) As Long

Public Declare Function acmDriverOpen Lib "msacm32" (ByRef phad As Long, _
        ByVal hadid As Long, ByVal fdwOpen As Long) As Long

Public Declare Function acmDriverClose Lib "msacm32" (ByVal had As Long, _
        ByVal fdwClose As Long) As Long

Public Declare Function acmFormatEnum Lib "msacm32" Alias "acmFormatEnumA" ( _
        ByVal had As Long, ByRef pafd As TACMFORMATDETAILS, ByVal fnCallback As Long, _
        ByRef dwInstance As Long, ByVal fdwEnum As ACM_FORMATENUMF) As Long

Public Declare Function acmFormatTagDetails Lib "msacm32" Alias "acmFormatTagDetailsA" ( _
        ByVal had As Long, paftd As TACMFORMATTAGDETAILS, ByVal fdwDetails As Long) As Long

Public Declare Function acmFormatDetails Lib "msacm32.dll" _
        Alias "acmFormatDetailsA" (ByVal had As Long, pafd As TACMFORMATDETAILS, _
        ByVal fdwDetails As Long) As Long


Public Type TACMDRIVERDETAILS
    cbStruct                    As Long
    fccType(3)                  As Byte
    fccComp(3)                  As Byte
    wMid                        As Integer
    wPid                        As Integer
    vdwACM                      As Long
    vdwDriver                   As Long
    fdwSupport                  As Long
    cFormatTags                 As Long
    cFilterTags                 As Long
    hIcon                       As Long
    szShortName(1 To 32)        As Byte
    szLongName(1 To 127)        As Byte
    szCopyright(1 To 80)        As Byte
    szLicensing(1 To 128)       As Byte
    szFeatures(1 To 512)        As Byte
End Type

Public Type TACMFORMATDETAILS
    cbStruct                    As Long
    dwFormatIndex               As Long
    dwFormatTag                 As Long
    fdwSupport                  As Long
    pwfx                        As Long
    cbwfx                       As Long
    szFormat(1 To 128)          As Byte
End Type

Public Type TACMFORMATTAGDETAILS
    cbStruct                    As Long
    dwFormatTagIndex            As Long
    dwFormatTag                 As Long
    cbFormatSize                As Long
    fdwSupport                  As Long
    cStandardFormats            As Long
    szFormatTag(1 To 128)       As Byte
End Type

Public Type ACMSTREAMHEADER
    cbStruct                    As Long
    fdwStatus                   As Long
    dwUser                      As Long
    pbSrc                       As Long
    cbSrcLength                 As Long
    cbSrcLengthUsed             As Long
    dwSrcUser                   As Long
    pbDst                       As Long
    cbDstLength                 As Long
    cbDstLengthUsed             As Long
    dwDstUser                   As Long
    dwReservedDriver(9)         As Long
End Type

Public Type WAVEFORMATEX
    wFormatTag                  As Integer
    nChannels                   As Integer
    nSamplesPerSec              As Long
    nAvgBytesPerSec             As Long
    nBlockAlign                 As Integer
    wBitsPerSample              As Integer
    cbSize                      As Integer
    xBytes(11)                  As Byte
End Type

Public Type MPEGLAYER3WAVEFORMAT
    wfx                         As WAVEFORMAT
    wID                         As Integer
    fdwFlags                    As Long
    nBlockSize                  As Integer
    nFramesPerBlock             As Integer
    nCodecDelay                 As Integer
End Type

Public Type FormatTag
    FormatTagIndex As Long
    FormatTag As Long
    szFormat As String
    wfx As WAVEFORMATEX
End Type

Public Type acmDriver
    handle As Long
    LongName As String
    ShortName As String
    FormatTagCount As Integer
    FormatTag() As FormatTag
End Type

Public Type drivers
    count As Integer
    drivers() As acmDriver
End Type
Public acmDrivers As drivers

Public Type CompatibleFormats
    haid As Long
    btWaveFormatEx() As Byte
End Type

Public cFoundFormats As Long
Public udtFoundFormats() As CompatibleFormats
Public udtTargetWaveFormat As WAVEFORMATEX

Public Enum ACM_WAVE_FORMAT
    WAVE_FORMAT_UNKNOWN = &H0
    WAVE_FORMAT_PCM = &H1
    WAVE_FORMAT_MPEGLAYER3 = &H55
End Enum

Public Const MPEGLAYER3_WFX_EXTRA_BYTES = 12

Public Const MPEGLAYER3_ID_UNKNOWN = 0
Public Const MPEGLAYER3_ID_MPEG = 1
Public Const MPEGLAYER3_ID_CONSTANTFRAMESIZE = 2

Public Const MPEGLAYER3_FLAG_PADDING_ISO = &H0
Public Const MPEGLAYER3_FLAG_PADDING_ON = &H1
Public Const MPEGLAYER3_FLAG_PADDING_OFF = &H2

Public Enum ACM_METRICS
    ACM_METRIC_COUNT_DRIVERS = 1
    ACM_METRIC_COUNT_CODECS = 2
    ACM_METRIC_COUNT_CONVERTERS = 3
    ACM_METRIC_COUNT_FILTERS = 4
    ACM_METRIC_COUNT_DISABLED = 5
    ACM_METRIC_COUNT_HARDWARE = 6
    ACM_METRIC_COUNT_LOCAL_DRIVERS = 20
    ACM_METRIC_COUNT_LOCAL_CODECS = 21
    ACM_METRIC_COUNT_LOCAL_CONVERTERS = 22
    ACM_METRIC_COUNT_LOCAL_FILTERS = 23
    ACM_METRIC_COUNT_LOCAL_DISABLED = 24
    ACM_METRIC_HARDWARE_WAVE_INPUT = 30
    ACM_METRIC_HARDWARE_WAVE_OUTPUT = 31
    ACM_METRIC_MAX_SIZE_FORMAT = 50
    ACM_METRIC_MAX_SIZE_FILTER = 51
    ACM_METRIC_DRIVER_SUPPORT = 100
    ACM_METRIC_DRIVER_PRIORITY = 101
End Enum

Public Enum ACM_STREAMCONVERTF
    ACM_STREAMCONVERTF_BLOCKALIGN = &H4
    ACM_STREAMCONVERTF_START = &H10
    ACM_STREAMCONVERTF_END = &H20
End Enum

Public Enum ACM_STREAMHEADER_STATUSF
    ACMSTREAMHEADER_STATUSF_DONE = &H10000
    ACMSTREAMHEADER_STATUSF_PREPARED = &H20000
    ACMSTREAMHEADER_STATUSF_INQUEUE = &H100000
End Enum

Public Enum ACM_STREAMOPENF
    ACM_STREAMOPENF_QUERY = &H1
    ACM_STREAMOPENF_ASYNC = &H2
    ACM_STREAMOPENF_NONREALTIME = &H4
End Enum

Public Enum ACM_STREAMSIZEF
    ACM_STREAMSIZEF_DESTINATION = &H1&
    ACM_STREAMSIZEF_SOURCE = &H0&
    ACM_STREAMSIZEF_QUERYMASK = &HF&
End Enum

Public Enum ACM_FORMATENUMF
    ACM_FORMATENUMF_WFORMATTAG = &H10000
    ACM_FORMATENUMF_NCHANNELS = &H20000
    ACM_FORMATENUMF_NSAMPLESPERSEC = &H40000
    ACM_FORMATENUMF_WBITSPERSAMPLE = &H80000
    ACM_FORMATENUMF_CONVERT = &H100000
    ACM_FORMATENUMF_SUGGEST = &H200000
    ACM_FORMATENUMF_HARDWARE = &H400000
    ACM_FORMATENUMF_INPUT = &H800000
    ACM_FORMATENUMF_OUTPUT = &H1000000
End Enum

Public Enum ACM_FORMATTAGDETAILS
    ACM_FORMATTAGDETAILSF_INDEX = &H0&
    ACM_FORMATTAGDETAILSF_FORMATTAG = &H1&
    ACM_FORMATTAGDETAILSF_LARGESTSIZE = &H2&
    ACM_FORMATTAGDETAILSF_QUERYMASK = &HF&
End Enum

Public Enum ACM_DRIVERDETAILS
    ACMDRIVERDETAILS_SUPPORTF_ASYNC = &H10&
    ACMDRIVERDETAILS_SUPPORTF_CODEC = &H1&
    ACMDRIVERDETAILS_SUPPORTF_LOCAL = &H40000000
    ACMDRIVERDETAILS_SUPPORTF_FILTER = &H4&
    ACMDRIVERDETAILS_SUPPORTF_HARDWARE = &H8&
    ACMDRIVERDETAILS_SUPPORTF_DISABLED = &H80000000
    ACMDRIVERDETAILS_SUPPORTF_CONVERTER = &H2&
End Enum

Public Enum ACM_CALLBACKS
    CALLBACK_NULL = &H0&
    CALLBACK_EVENT = &H50000
    CALLBACK_WINDOW = &H10000
    CALLBACK_THREAD = &H20000
    CALLBACK_TYPEMASK = &H70000
    CALLBACK_FUNCTION = &H30000
End Enum

Public Enum ACM_STATUS
    MM_STREAM_OPEN = &H3D4
    MM_STREAM_CLOSE = &H3D5
    MM_STREAM_DONE = &H3D6
    MM_STREAM_ERROR = &H3D7
End Enum

Public Enum ACM_ERRORS
    ACMERR_BUSY = &H201&
    ACMERR_CANCELED = &H203&
    ACMERR_UNPREPARED = &H202&
    ACMERR_NOTPOSSIBLE = &H200&
End Enum

Public Enum MMSYS_ERRORS
    MMSYSERR_NOERROR = 0
    MMSYSERR_ERROR = 1
    MMSYSERR_BADDEVICEID = 2
    MMSYSERR_NOTENABLED = 3
    MMSYSERR_ALLOCATED = 4
    MMSYSERR_INVALHANDLE = 5
    MMSYSERR_NODRIVER = 6
    MMSYSERR_NOMEM = 7
    MMSYSERR_NOTSUPPORTED = 8
    MMSYSERR_BADERRNUM = 9
    MMSYSERR_INVALFLAG = 10
    MMSYSERR_INVALPARAM = 11
    MMSYSERR_HANDLEBUSY = 12
    MMSYSERR_INVALIDALIAS = 13
    MMSYSERR_BADDB = 14
    MMSYSERR_KEYNOTFOUND = 15
    MMSYSERR_READERROR = 16
    MMSYSERR_WRITEERROR = 17
    MMSYSERR_DELETEERROR = 18
    MMSYSERR_VALNOTFOUND = 19
    MMSYSERR_NODRIVERCB = 20
End Enum

' not used atm
Public Sub acmStreamConvertCallback( _
            ByVal hStream As Long, _
            ByVal uMsg As Long, _
            ByVal dwInstance As Long, _
            ByRef lParam1 As Long, _
            ByRef lParam2 As Long)

    Select Case uMsg

        Case MM_STREAM_OPEN:
            '

        Case MM_STREAM_CLOSE:
            '

        Case MM_STREAM_DONE:
            '

    End Select

End Sub

'"  GetCompatibleCodecs SourceWaveFormat, DestWaveFormat
'   This function may be used (although it ain't necessary when using file and realtime
'   conversion functions) to determine all Drivers that support conversion between the two
'   specified Formats. The result is returned as an Array of CompatibleFormats Structures. "
Public Function GetCompatibleCodecs( _
        SourceWaveFormat As WAVEFORMATEX, _
        DestWaveFormat As WAVEFORMATEX) _
        As CompatibleFormats()

    Erase udtFoundFormats
    udtTargetWaveFormat = DestWaveFormat

    Call acmDriverEnum(AddressOf DriverEnumCallback, _
                       ByVal VarPtr(SourceWaveFormat), _
                       0)

    cFoundFormats = 0

    On Error Resume Next
    cFoundFormats = UBound(udtFoundFormats) + 1
    
    GetCompatibleCodecs = udtFoundFormats

End Function

Private Function DriverEnumCallback( _
            ByVal haid As Long, _
            WAVEFORMAT As WAVEFORMATEX, _
            ByVal fdwSupport As Long) As Long

    Dim hDriver             As Long
    Dim lMaxSize            As Long
    Dim udtFormatDetails    As TACMFORMATDETAILS
    Dim btWaveFormat()      As Byte

    ' conversion supported?
    If Not fdwSupport And ACMDRIVERDETAILS_SUPPORTF_CODEC Then
        DriverEnumCallback = 1
        Exit Function
    End If

    If acmDriverOpen(hDriver, haid, 0) = 0 Then

        ' get size of biggest possible format
        acmMetrics hDriver, _
                   ACM_METRIC_MAX_SIZE_FORMAT, _
                   udtFormatDetails.cbwfx

        ' prepare a byte array to hold the formats
        WAVEFORMAT.cbSize = udtFormatDetails.cbwfx - Len(WAVEFORMAT)
        ReDim btWaveFormat(udtFormatDetails.cbwfx + 1)

        CopyMemory btWaveFormat(0), WAVEFORMAT, Len(WAVEFORMAT)
        With udtFormatDetails
            .cbStruct = LenB(udtFormatDetails)
            .pwfx = VarPtr(btWaveFormat(0))
            .dwFormatTag = WAVEFORMAT.wFormatTag
        End With

        ' get formats for current driver
        Call acmFormatEnum(hDriver, _
                           udtFormatDetails, _
                           AddressOf FormatCallback, _
                           ByVal haid, _
                           ACM_FORMATENUMF_CONVERT)

        Call acmDriverClose(hDriver, 0)

    Else
        Err.Raise 513, , "acmDriverOpen Error"
    End If

    ' 1 for "next driver",
    ' 0 for "abort"
    DriverEnumCallback = 1

End Function

Private Function FormatCallback( _
            ByVal hACMDriverID As Long, _
            ACMFmtDet As TACMFORMATDETAILS, _
            ByVal haid As Long, _
            ByVal fdwSupport As Long) _
            As Long

    Dim udtFormat As WAVEFORMATEX

    ' get format structure
    CopyMemory udtFormat, ByVal ACMFmtDet.pwfx, Len(udtFormat)

    ' format tag is the one we search for?
    If udtFormat.wFormatTag = udtTargetWaveFormat.wFormatTag Then

        ' format fits?
        If udtFormat.nSamplesPerSec = udtTargetWaveFormat.nSamplesPerSec Then
            If udtFormat.nChannels = udtTargetWaveFormat.nChannels Then
                If udtFormat.nAvgBytesPerSec = udtTargetWaveFormat.nAvgBytesPerSec Then

                    On Error Resume Next

                    ' prepare a new index to hold the format
                    ReDim Preserve udtFoundFormats(UBound(udtFoundFormats) + 1)

                    If Err <> 0 Then
                        ReDim udtFoundFormats(0)
                    End If

                    ' store driver handle and format structure
                    With udtFoundFormats(UBound(udtFoundFormats))
                        .haid = haid
                        ReDim .btWaveFormatEx(ACMFmtDet.cbwfx - 1)
                        CopyMemory .btWaveFormatEx(0), ByVal ACMFmtDet.pwfx, ACMFmtDet.cbwfx
                    End With

                End If
            End If
        End If
    End If

    ' get next format
    FormatCallback = 1

End Function

Public Sub ListCodecs()
    Dim mmr     As Long

    mmr = acmDriverEnum(AddressOf acmDriverEnumCallback, 0, 0)
End Sub

'Rückruffunktion für die Treiberaufzählung
Public Function acmDriverEnumCallback(ByVal haid As Long, _
                                      ByVal dwInstance As Long, _
                                      ByVal fdwSupport As Long) As Long

    Dim i           As Integer
    Dim driver      As Long
    Dim mmr         As Long

    'Wenn der Treiber die Konvertierung von einem Format
    'in das andere unterstützt, ...
    If fdwSupport And ACMDRIVERDETAILS_SUPPORTF_CODEC Then

        'Treiberhandle holen
        driver = 0
        mmr = acmDriverOpen(driver, haid, 0)

        'Treiberdetails lesen
        Dim details     As TACMDRIVERDETAILS
        details.cbStruct = Len(details)
        mmr = acmDriverDetails(haid, details, 0)

        ReDim Preserve acmDrivers.drivers(acmDrivers.count)

        With acmDrivers.drivers(acmDrivers.count)
            .handle = haid
            .LongName = details.szLongName
            .ShortName = details.szShortName
        End With

        ' mögliche Format Tags abarbeiten
        For i = 1 To details.cFormatTags

            'Details zum aktuellen Format Tag lesen
            Dim fmtTagDetails  As TACMFORMATTAGDETAILS
            fmtTagDetails.cbStruct = Len(fmtTagDetails)
            fmtTagDetails.dwFormatTagIndex = i
            mmr = acmFormatTagDetails(driver, fmtTagDetails, 0)

            Dim fmtDetails      As TACMFORMATDETAILS
            Dim wavformat       As WAVEFORMATEX

            '... dann eine Enumeration der unterstützten
            'Formate vorbereiten
            wavformat.cbSize = Len(wavformat)
            wavformat.wFormatTag = WAVE_FORMAT_MPEGLAYER3
    
            'größte Größe von WaveFormatEx bestimmen
            Dim MaxSize         As Long
            acmMetrics driver, ACM_METRIC_MAX_SIZE_FORMAT, MaxSize
            If MaxSize < Len(wavformat) Then MaxSize = Len(wavformat)
    
            wavformat.cbSize = (MaxSize And &HFFFF&) - Len(wavformat)
            wavformat.wFormatTag = WAVE_FORMAT_UNKNOWN
    
            fmtDetails.cbStruct = Len(fmtDetails)
            fmtDetails.pwfx = VarPtr(wavformat)
            fmtDetails.cbwfx = MaxSize
            fmtDetails.dwFormatTag = WAVE_FORMAT_UNKNOWN
    
            'Formate aufzählen
            mmr = acmFormatEnum(driver, fmtDetails, _
                                AddressOf acmFormatCallback, _
                                0, 0)
        Next

        'Treiber schließen
        mmr = acmDriverClose(driver, 0)

        acmDrivers.count = acmDrivers.count + 1
    End If

    '1 (true) zurückgeben für nächsten Treiber,
    '0, um acmDriverEnum abzubrechen
    acmDriverEnumCallback = 1
End Function

'Rückruffunktion für die Formataufzählung
Public Function acmFormatCallback(ByVal hACMDriverID As Long, _
                                  ACMFmtDet As TACMFORMATDETAILS, _
                                  ByVal dwInstance As Long, _
                                  ByVal fdwSupport As Long) As Long

    Dim format As WAVEFORMATEX

    '1 (true) zurückgeben für nächstes Format,
    '0, um acmFormatCallback abzubrechen
    acmFormatCallback = 1

    'Daten von WaveFormatEx Pointer in wavformat kopieren
    CopyMemory format, ByVal ACMFmtDet.pwfx, Len(format)

    'Format speichern
    With acmDrivers.drivers(acmDrivers.count)
        ReDim Preserve .FormatTag(.FormatTagCount) As FormatTag
    End With

    With acmDrivers.drivers(acmDrivers.count).FormatTag(acmDrivers.drivers(acmDrivers.count).FormatTagCount)
        .wfx = format
        .szFormat = TrimNull(StrConv(ACMFmtDet.szFormat, vbUnicode))
        .FormatTag = ACMFmtDet.dwFormatTag
        .FormatTagIndex = ACMFmtDet.dwFormatIndex
    End With

    With acmDrivers.drivers(acmDrivers.count)
        .FormatTagCount = .FormatTagCount + 1
    End With
End Function

Public Function TrimNull(strVal As String) As String
    TrimNull = Left$(strVal, InStr(strVal, Chr$(0)) - 1)
End Function
