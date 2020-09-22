Attribute VB_Name = "modMMC"
Option Explicit

Public cd As New clsCDROM                             ' CD-ROM class

Public Declare Function mciSendString Lib "winmm.dll" _
Alias "mciSendStringA" ( _
    ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long _
) As Long

Private Declare Function GetShortPathName Lib "kernel32" _
Alias "GetShortPathNameA" ( _
    ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long _
) As Long

Public Type t_MSFLBA
    M                           As Byte               ' minutes
    s                           As Byte               ' seconds
    F                           As Byte               ' frames
    LBA                         As Long               ' logical block address
End Type

Public Type t_CDInfo
    Capacity                    As Long               ' capacity (only CD-R[W])
    LeadIn                      As t_MSFLBA           ' Lead-In Start
    LeadOut                     As t_MSFLBA           ' Lead-Out Start
    DiscStatus                  As e_Status           ' CD Status
    LastSessionStatus           As e_Status           ' last session's status
    CDType                      As e_CDType           ' CD Type
    CDSubType                   As e_CD_SubType       ' Sub Type
    Erasable                    As Boolean            ' erasable?
    Tracks                      As Byte               ' number of tracks
    Sessions                    As Byte               ' number of sessions
    Size                        As Double               ' size of the disc
    Vendor                      As String             ' CD-R(W) vendor
End Type

Public Type t_ATIP
    Length(1)                   As Byte               ' data len
    rsvd1(1)                    As Byte               ' reserved
    ITWP                        As Byte               ' ?
    Uru                         As Byte               ' ?
    DiscType                    As Byte               ' CD Type (CD-R/CD-RW)
    rsvd2                       As Byte               ' reserved
    LeadIn_Min                  As Byte               ' Lead-In Start (minutes)
    LeadIn_Sec                  As Byte               ' Lead-In Start (seconds)
    LeadIn_Frm                  As Byte               ' Lead-In Start (frames)
    rsvd3                       As Byte               ' reserved
    LeadOut_Min                 As Byte               ' Lead-Out Start (minutes)
    LeadOut_Sec                 As Byte               ' Lead-Out Start (seconds)
    LeadOut_Frm                 As Byte               ' Lead-Out Start (frames)
    rsvd4                       As Byte               ' reserved
    Rest(12)                    As Byte               ' rest
End Type

Public Type t_RDI
    PageLen(1)                  As Byte               ' Page len
    states                      As Byte               ' misc. data (erasable, ...)
    FirstTrack                  As Byte               ' first track on the disc (experience: 1...)
    NumSessionsLSB              As Byte               ' number of sessions
    FirstTrackLastSessionLSB    As Byte               ' first track in last session
    LastTrackLastSessionLSB     As Byte               ' last track in last session
    misc                        As Byte               ' misc.
    DiscType                    As Byte               ' CD sub type (CD-ROM/CD-I/XA)
    NumSessionsMSB              As Byte               ' number of sessions
    FirstTrackLastSessionMSB    As Byte               ' first track in last session
    LastTrackLastSessionMSB     As Byte               ' last track in last session
    DiscIdentification(3)       As Byte               ' CD ID
    LastSessionLeadInStart(3)   As Byte               ' Lead-In start time (h:m:s:f)
    LastPossibleLeadOutStart(3) As Byte               ' last possible Lead-Out Start (h:m:s:f)
    DBC(6)                      As Byte               ' Disc Bar Code
End Type

Public Type t_ReadCap
    Blocks(3)                   As Byte               ' written sectors
    BlockLen(3)                 As Byte               ' sectorsize
End Type

Public Type t_Track
    TrkNum                      As Byte               ' track number
    AudioTrack                  As Boolean            ' Audio Track?
    Start                       As t_MSFLBA           ' Track Start
    end                         As t_MSFLBA           ' Track Ende
End Type

Public Type t_Tracks
    Tracks                      As Byte               ' number of tracks
    Track(98)                   As t_Track            ' Track Collection (max. 99)
End Type

Public Type t_TOC_TRACK
    rsvd1                       As Byte               ' reserved
    ADR                         As Byte               ' ADR
    Track                       As Byte               ' Track
    rsvd2                       As Byte               ' reserved
    addr(3)                     As Byte               ' address (h:m:s:f or LBA)
End Type

Public Type t_TOC_STRUCT
    TocLen(1)                   As Byte               ' data len
    FirstTrack                  As Byte               ' first track
    LastTrack                   As Byte               ' letzter track
    TocTrack(99)                As t_TOC_TRACK        ' tracks
End Type

Public Type t_TrackInfo
    Track                       As Byte               ' track number
    Session                     As Byte               ' session number
    DataMode                    As e_SectorModes      ' data mode
    startLBA                    As Long               ' Start LBA
    endLBA                      As Long               ' End LBA
    Length                      As Long               ' length in sectors
    LastTrackInSession          As Boolean            ' last track in session?
End Type

Private Type t_RTOC_ENTRY
    sessionNr                   As Byte               ' current session
    ADRCTL                      As Byte               ' ADR/CTRL
    TNO                         As Byte               ' track number
    point                       As Byte               ' packet type
    min                         As Byte               ' minutes
    sec                         As Byte               ' seconds
    frm                         As Byte               ' frames
    zero                        As Byte               ' 0
    pmin                        As Byte               ' Point minutes
    psec                        As Byte               ' Point seconds
    pframe                      As Byte               ' Point frames
End Type

Public Type t_RTOC_STRUCT
    dummy(3)                    As Byte               ' Header
    packet(255)                 As t_RTOC_ENTRY       ' packets
End Type

Public Type t_CD_TEXT_PACKET
    idType                      As Byte               ' packet type
    idTrk                       As Byte               ' track number
    idSeq                       As Byte               ' sequence
    idFlg                       As Byte               ' flags
    txt(11)                     As Byte               ' Text data (ASCII)
    CRC(1)                      As Byte               ' CRC (Cyclic Redundancy Check)
End Type

Public Type t_CD_TEXT
    dummy(3)                    As Byte               ' Header
    CDText(255)                 As t_CD_TEXT_PACKET   ' CD-Text packets
End Type

Public Type t_DVD_CPYINFO
    Length(1)                   As Byte
    rsvd(1)                     As Byte
    cpyprotectsystype           As Byte
    regioninfo                  As Byte
End Type

Public Type t_DVD_Phys
    Length(1)                   As Byte               ' data len
    rsvd(1)                     As Byte               ' reserved
    BookType                    As Byte               ' Booktype (DVD-ROM, DVD-RAM, DVD-R, DVD+R, ...)
    discsize                    As Byte               ' DVD size/Max. Rate
    LayerType                   As Byte               ' number of layers/Layer Type
    TrackDens                   As Byte               ' Track density
    zero                        As Byte               ' 0
    StartSector(2)              As Byte               ' physical start sector
    zero2                       As Byte               ' 0
    EndSector(2)                As Byte               ' physical end sector
    zero3                       As Byte               ' 0
    EndSectorLayer0(2)          As Byte               ' physical end sector in layer 0
    bca                         As Byte               ' bca
    mspec(2030)                 As Byte               ' depends on disc
End Type

Public Type t_MMC
    PageCode                    As Byte               ' Page Code
    PageLen                     As Byte               ' Page len
    rsvd2(7)                    As Byte               ' reserved
    ReadSupported               As Byte               ' readable formats
    WriteSupported              As Byte               ' writable formats
    misc(3)                     As Byte               ' misc.
    MaxReadSpeed(1)             As Byte               ' max. read speed
    NumVolLevels(1)             As Byte               ' num. volume levels
    BufferSize(1)               As Byte               ' buffer size
    CurrReadSpeed(1)            As Byte               ' curr. read speed
    rsvd                        As Byte               ' reserved
    misc2                       As Byte               ' misc.
    MaxWriteSpeed(1)            As Byte               ' max. write speed
    CurrWriteSpeed(1)           As Byte               ' curr write speed
    RotationControl             As Byte
    CurrWriteSpeedMMC3(1)       As Byte
End Type

Public Type t_Speed
    MaxRSpeed                   As Integer            ' max. read speed
    MaxWSpeed                   As Integer            ' max. write speed
    CurrRSpeed                  As Integer            ' curr. read speed
    CurrWSpeed                  As Integer            ' curr write speed
End Type

Public Type t_ReadFeatures                            ' can read:
    CDR                         As Boolean            '      CD-R
    CDRW                        As Boolean            '      CD-RW
    DVDR                        As Boolean            '      DVD-R
    DVDROM                      As Boolean            '      DVD-ROM
    DVDRAM                      As Boolean            '      DVD-RAM
    subchannels                 As Boolean            '      Sub-Channels
    SubChannelsCorrected        As Boolean            '      Sub-Channels corrected
    SubChannelsFormLeadIn       As Boolean            '      Sub-Channels from Lead-In
    C2ErrorPointers             As Boolean            '      C2 Error Pointers
    ISRC                        As Boolean            '      International Standard Recording Code
    UPC                         As Boolean            '      ?
    BC                          As Boolean            '      Bar Code
    Mode2Form1                  As Boolean            '      Mode 2 Form 1 sectors
    Mode2Form2                  As Boolean            '      Mode 2 Form 2
    Multisession                As Boolean            '      Multi-Session CDs
    CDDARawRead                 As Boolean            '      Audio sectors
End Type

Public Type t_WriteModes
    Raw96                       As Boolean            ' Raw + 96
    Raw16                       As Boolean            ' Raw + 16
    SAO                         As Boolean            ' Session At Once
    TAO                         As Boolean            ' Track At Once
    Raw96Test                   As Boolean            ' Raw + 96 + Test-Mode
    Raw16Test                   As Boolean            ' Raw + 16 + Test-Mode
    SAOTest                     As Boolean            ' Session At Once + Test-Mode
    TAOTest                     As Boolean            ' Track At Once + Test-Mode
End Type

Public Type t_WriteFeatures                           ' can write:
    CDR                         As Boolean            '      CD-R
    CDRW                        As Boolean            '      CD-RW
    DVDR                        As Boolean            '      DVD-R
    DVDRAM                      As Boolean            '      DVD-RAM
    TestMode                    As Boolean            '      Test-Mode
    BURNProof                   As Boolean            '      BURN-Proof
    WriteModes                  As t_WriteModes       ' supported write modes
End Type

Public Type t_DrvInfo
    ReadFeatures                As t_ReadFeatures     ' read features
    WriteFeatures               As t_WriteFeatures    ' write features
    speeds                      As t_Speed            ' speeds
    AnalogAudio                 As Boolean            ' analog audio playback?
    JitterCorrection            As Boolean            ' jitter effect correction?
    BufferSize                  As Long               ' buffer size
    LockMedia                   As Boolean            ' can lock media?
    LoadingMechanism            As e_LoadingMechanism ' loading mechanism
    Interface                   As e_DrvInterfaces    ' drive interface
End Type

Public Type t_RTI
    DataLen(1)                  As Byte               ' data length
    TrackNumLSB                 As Byte               ' track number
    SessionNumLSB               As Byte               ' track session
    rsvd                        As Byte               ' reserved
    TrackMode                   As Byte               ' Track Mode
    DataMode                    As Byte               ' Data Mode
    misc                        As Byte               ' misc.
    Track_Start(3)              As Byte               ' track start (h:m:s:f)
    Track_Next_Writable(3)      As Byte               ' next writable address (h:m:s:f)
    Track_Free_Blocks(3)        As Byte               ' free sectors (h:m:s:f)
    Track_Packet_Size(3)        As Byte               ' Fixed Packet size (?)
    Track_Size(3)               As Byte               ' Track size (LBA)
    Track_Last_Recorded(3)      As Byte               ' last written address
    TrackNumMSB                 As Byte               ' Track Number
    SessionNumMSB               As Byte               ' Track Session
    rsvd2(1)                    As Byte               ' reserved
End Type

Private Type t_InqDat
    PDT                         As Byte               ' drive type
    PDQ                         As Byte               ' removable drive
    VER                         As Byte               ' MMC Version (zero for ATAPI)
    RDF                         As Byte               ' interface depending field
    DLEN                        As Byte               ' additional len
    rsv1(1)                     As Byte               ' reserved
    Feat                        As Byte               ' ?
    VID(7)                      As Byte               ' vendor
    PID(15)                     As Byte               ' Product
    PVER(3)                     As Byte               ' revision (= Firmware Version)
    FWVER(20)                   As Byte               ' ?
End Type

Public Type t_BufferCapacity
    DataLen(1)                  As Byte               ' data len
    rsvd(1)                     As Byte               ' reserved
    BufferLen(3)                As Byte               ' buffer size
    BufferBlank(3)              As Byte               ' empty part of the buffer
End Type

Public Type t_WavHdr
    riff                        As String * 4         ' RIFF Header
    len                         As Long               ' file length
    WavFmt                      As String * 8         ' Wave Format
    HdrLen                      As Long               ' header length
    format                      As Integer            ' format
    NumChannels                 As Integer            ' channels
    SampleRate                  As Long               ' frequency
    BytesPerSec                 As Long               ' Bytes/second
    BlockAlign                  As Integer            ' Block Align
    BitsPerSample               As Integer            ' Bits/Sample
    Data                        As String * 4         ' Data Chunk
    DataLen                     As Long               ' datalength w/o header
End Type

'Structure of the Q Sub-Channel data header returned by ReadSubChannel
Public Type t_SubQHeader
    rsvd As Byte            ' reserved
    audio_stat As Byte      ' audio status (Play, Stop, ...)
    Data_Len(1) As Byte     ' data length
End Type

'Structure of the Q Sub-Channel information (Current Position) returned by ReadSubChannel
Public Type t_CDCurrPos
    header As t_SubQHeader  ' header (top)
    dataformat As Byte      ' data format of the requested information
    ADRCTRL As Byte         ' ADR/CTRL
    TrkNum As Byte          ' track number
    IndexNum As Byte        ' index number of the current track (0 = Pre-Gap, 1 = data)
    AbsCDAddr(3) As Byte    ' absolute address (position)
    TrkRelAddr(3) As Byte   ' track relative address (position within a track)
End Type

'Structure of the Q Sub-Channel information (ISRC) returned by ReadSubChannel
Public Type t_Q_ISRC
    header As t_SubQHeader  ' header (top)
    dataformat As Byte      ' data format of the requested information
    ADRCTRL As Byte         ' ADR/CTRL
    TrkNum As Byte          ' track number
    rsvd As Byte            ' reserved
    ISRC(15) As Byte        ' ISRC (International Standard Recording Code)
End Type

'Structure of the Q Sub-Channel information (MCN) returned by ReadSubChannel
Public Type t_Q_MCN
    header As t_SubQHeader  ' header (top)
    dataformat As Byte      ' data format of the requested information
    rsvd(2) As Byte         ' reserved
    MCN(15) As Byte         ' MCN (Media Catalog Number)
End Type

Public Type t_Feat_Hdr
    DataLen(3) As Byte      ' data length
    rsvd(1) As Byte         ' reserved
    curr_profile(1) As Byte ' current profile
End Type

Public Type t_Feat_CD_READ
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version (= 1)
    additional_len As Byte  ' additional length (= 4)
    CDText As Byte          ' can read CD-Text and C2?
    rsvd(2) As Byte         ' reserved
End Type

Public Type t_Feat_DVD_P_RW
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version (= 0)
    additional_len As Byte  ' additional length (= 4)
    write As Byte           ' can write DVD+RW
    close_only As Byte      ' background format
    rsvd(1) As Byte         ' reserved
End Type

'Feature 2Fh: DVD-R/RW
Public Type t_Feat_DVD_R_RW
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version (= 1)
    additional_len As Byte  ' additional length (= 4)
    writeDVDRW As Byte      ' can write DVD-RW
    rsvd(2) As Byte         ' reserved
End Type

'Feature 2Bh: DVD+R
Public Type t_Feat_DVD_P_R
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version
    additional_len As Byte  ' additional length
    write As Byte           ' can write DVD+R
    rsvd(2) As Byte         ' reserved
End Type

' Feature 3Bh: DVD+R Double Layer
Public Type t_Feat_DVD_P_R_DL
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version
    additional_len As Byte  ' additional length
    write As Byte           ' can write DVD+R DL
    rsvd(2) As Byte         ' reserved
End Type

' Feature 28h: Mount Rainer
Public Type t_Feat_MRW
    hdr As t_Feat_Hdr       ' header
    code(1) As Byte         ' Feature code
    VER As Byte             ' version
    additional_len As Byte  ' additional length
    write As Byte           ' can write MRW
    rsvd(2) As Byte         ' reserved
End Type

' Multimedia Capabilities page write speed descriptor
Public Type t_MMCP_WriteSpeed
    rsvd        As Byte     ' reserved
    rotation    As Byte     ' rotation control
    speed(1)    As Byte     ' speed in kb/s
End Type

' generic Q Sub-Channel format
Public Type t_Qch
    CTLADR      As Byte     ' control/ADR
    DATAQ(8)    As Byte     ' Q data
    CRC(1)      As Byte     ' CRC
End Type

' Mode 1 Q channel (0001b)
Public Type t_Qch_Mode1
    CTLADR      As Byte     ' control/ADR
    TNO         As Byte     ' track number
    index       As Byte     ' track index
    min         As Byte     ' rel. minutes
    sec         As Byte     ' rel. seconds
    frm         As Byte     ' rel. frames
    zero        As Byte     ' zero
    amin        As Byte     ' abs. minutes
    asec        As Byte     ' abs. seconds
    afrm        As Byte     ' abs. frames
    CRC(1)      As Byte     ' CRC
End Type

' Mode 2 Q channel (0010b)
Public Type t_Qch_Mode2
    CTLADR      As Byte     ' control/ADR
    MCN(7)      As Byte     ' media catalog number
    afrm        As Byte     ' abs. frames
    CRC(1)      As Byte     ' CRC
End Type

' Mode 3 Q channel (0011b)
Public Type t_Qch_Mode3
    CTLADR      As Byte     ' control/ADR
    ISRC(7)     As Byte     ' 6 bit cells
    afrm        As Byte     ' abs. frames
    CRC(1)      As Byte     ' CRC
End Type

Public Enum e_DrvInterfaces
    IF_SCSI                                           ' SCSI
    IF_ATAPI                                          ' ATAPI
    IF_IEEE                                           ' IEEE
    IF_USB                                            ' USB
    IF_UNKNWN                                         ' unknown
End Enum

Public Enum e_READCD_FLAGS
    RCD_SYNC = &H80                                   ' sync pattern
    RCD_HDR_4BT = &H20                                ' 4-Bytes Header
    RCD_HDR_8BT = &H40                                ' 8-Bytes Header
    RCD_USRDATA = &H10                                ' userdata
    RCD_EDC_ECC = &H8                                 ' EDC+ECC correction
    RCD_RAW = &HF8                                    ' full sector
End Enum

Public Enum e_READCD_SUBCH_FLAGS
    RCD_SUBCH_PW_RAW = &H1                            ' Raw Channels P-W
    RCD_SUBCH_Q = &H2                                 ' only Q channel
    RCD_SUBCH_PW_CORRECTED = &H4                      ' Channels P-W corrected
End Enum

Public Enum e_BLANKTYPE
    BLANK_FULL                                        ' full blank
    BLANK_QUICK                                       ' quick blank
End Enum

Public Enum e_SectorModes
    MODE_AUDIO                                        ' Audio (or Mode-0)
    MODE_MODE1                                        ' Mode-1
    MODE_MODE2                                        ' Plain Mode-2
    MODE_MODE2_FORM1                                  ' Mode-2 Form-1
    MODE_MODE2_FORM2                                  ' Mode-2 Form-2
End Enum

Public Enum e_Genres
    GENRE_ADULT_CONTEMP = 4                           ' adult contemporary
    GENRE_ALT_ROCK = 5                                ' alternative rock
    GENRE_CHILDREN = 6                                ' children music
    GENRE_CLASSIC = 7                                 ' classic
    GENRE_CHRIST_CONTEMP = 6                          ' christian contemporary
    GENRE_COUNTRY = 7                                 ' country
    GENRE_DANCE = 8                                   ' dance
    GENRE_EASY_LISTENING = 9                          ' easy listening
    GENRE_EROTIC = 10                                 ' erotic
    GENRE_FOLK = 11                                   ' folk
    GENRE_GOSPEL = 12                                 ' gospel
    GENRE_HIPHOP = 13                                 ' hip hop
    GENRE_JAZZ = 14                                   ' jazz
    GENRE_LATIN = 15                                  ' latin
    GENRE_MUSICAL = 16                                ' musical
    GENRE_NEWAGE = 17                                 ' new age
    GENRE_OPERA = 18                                  ' opera
    GENRE_OPERETTA = 19                               ' operetta
    GENRE_POP = 20                                    ' pop
    GENRE_RAP = 21                                    ' rap
    GENRE_REGGAE = 22                                 ' reggae
    GENRE_ROCK = 23                                   ' rock
    GENRE_RYTHMANDBLUES = 24                          ' R'n'B
    GENRE_SOUNDEFFECTS = 25                           ' sound effects
    GENRE_SPOKEN_WORD = 26                            ' spoken words
    GENRE_WORLD_MUSIC = 27                            ' world music
End Enum

Public Enum e_CDType
    ROMTYPE_CDROM                                     ' CD-ROM
    ROMTYPE_CDR                                       ' CD-R
    ROMTYPE_CDRW                                      ' CD-RW
    ROMTYPE_CDROM_R_RW                                ' CD-ROM, CD-R oder CD-RW
    ROMTYPE_DVD_ROM                                   ' DVD-ROM
    ROMTYPE_DVD_R                                     ' DVD-R
    ROMTYPE_DVD_RW                                    ' DVD-RW
    ROMTYPE_DVD_RAM                                   ' DVD-RAM
    ROMTYPE_DVD_P_R                                   ' DVD+R
    ROMTYPE_DVD_P_RW                                  ' DVD+RW
End Enum

Public Enum e_LoadingMechanism
    LOAD_CADDY                                        ' Caddy
    LOAD_TRAY                                         ' Tray
    LOAD_POPUP                                        ' Popup
    LOAD_CHANGER                                      ' Changer
    LOAD_UNKNWN                                       ' Unknown
End Enum

Public Enum e_Status
    STAT_EMPTY                                        ' empty
    STAT_INCOMPLETE                                   ' uncomplete
    STAT_COMPLETE                                     ' complete
    STAT_UNKNWN                                       ' unknown
End Enum

Public Enum e_CD_SubType
    STYPE_CDROMDA                                     ' CD-ROM or CDDA
    STYPE_CDI                                         ' CD-I
    STYPE_XA                                          ' CD-XA
    STYPE_UNKNWN                                      ' unknown
End Enum

'(from CDROM-TOOL [GPL])
Public Enum e_SpinDown
    SD_VS                                             ' vendor specific
    SD_125MS                                          ' 125 ms
    SD_250MS                                          ' 250     "
    SD_500MS                                          ' 500     "
    SD_1SEC                                           '   1 sec
    SD_2SEC                                           '   2     "
    SD_4SEC                                           '   4     "
    SD_8SEC                                           '   8     "
    SD_16SEC                                          '  16     "
    SD_32SEC                                          '  32     "
    SD_1MIN                                           '   1 min
    SD_2MIN                                           '   2     "
    SD_4MIN                                           '   4     "
    SD_8MIN                                           '   8     "
    SD_16MIN                                          '  16     "
    SD_32MIN                                          '  32     "
End Enum

Public Enum e_CDTextPacketTypes
    CD_TEXT_PACK_ALBUM_NAME = &H80
    CD_TEXT_PACK_PERFORMER = &H81
    CD_TEXT_PACK_SONGWRITER = &H82
    CD_TEXT_PACK_COMPOSER = &H83
    CD_TEXT_PACK_ARRANGER = &H84
    CD_TEXT_PACK_MESSAGES = &H85
    CD_TEXT_PACK_DISC_ID = &H86
    CD_TEXT_PACK_GENRE = &H87
    CD_TEXT_PACK_TOC_INFO = &H88
    CD_TEXT_PACK_TOC_INFO2 = &H89
    CD_TEXT_PACK_UPC_EAN = &H8E
    CD_TEXT_PACK_SIZE_INFO = &H8F
End Enum

'datasectorsynchronisationpattern :)
Public Const SYNCPATTERN As String = "00FFFFFFFFFFFFFFFFFFFF00"

Public Function CDRomTestUnitReady(ByVal DrvID As String) As Boolean

    Dim cmd(5) As Byte
    Dim i      As Integer

    If cd.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0) Then
        CDRomTestUnitReady = True
        Exit Function
    End If

    ' no disc present
    If cd.LastASC = &H3A Then
        '
        Exit Function

    ' unit is becoming ready,
    ' or not ready to ready change
    ' because medium may have changed,
    ' wait for it
    ElseIf (cd.LastASC = &H4 And cd.LastASCQ = &H1) _
    Or (cd.LastSK = 6 And cd.LastASC = 40) Then

        ' try 5 times (~5 seconds)
        For i = 1 To 5
            If cd.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0) Then Exit For
            Sleep 1000
        Next i

        CDRomTestUnitReady = cd.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0)

    End If

End Function

'disc present?
Public Function CDRomIsDiscPresent(ByVal DrvID As String) As Boolean

    Dim media_event_req(8) As Byte

    'get Tray Status
    If Not CDRomGetEventStatusNotification(DrvID, &H10, _
                                           VarPtr(media_event_req(0)), _
                                           UBound(media_event_req)) Then

        CDRomIsDiscPresent = CDRomTestUnitReady(DrvID)
        Exit Function

    End If

    'valid data?
    If media_event_req(0) = 0 And media_event_req(1) = 0 Then
        CDRomIsDiscPresent = CDRomTestUnitReady(DrvID)
        Exit Function
    Else
        'disc present?
        CDRomIsDiscPresent = IsBitSet(media_event_req(5), 1)
    End If

End Function

Public Function CDRomGetWriteSpeeds(ByVal DrvID As String, _
                                    speeds() As Integer) As Boolean

    Dim buf(512)        As Byte
    Dim mpage()         As Byte
    Dim udtDescriptor   As t_MMCP_WriteSpeed
    Dim intSize         As Integer
    Dim intDescriptors  As Integer
    Dim i               As Integer

    ' get MMCP
    If Not CDRomModeSense10(DrvID, &H2A, VarPtr(buf(0)), 512, True, True) Then
        Exit Function
    End If

    ' get size of the page
    intSize = cd.LShift(buf(0), 8) Or buf(1)

    ' set new buffer to grab full page
    ReDim mpage(intSize + 1) As Byte

    ' get the whole MMCP
    If Not CDRomModeSense10(DrvID, &H2A, VarPtr(mpage(0)), intSize + 2, True, True) Then
        Exit Function
    End If

    If intSize > 38 Then

        ' get the number of write speed descriptors
        intDescriptors = (cd.LShift(mpage(30 + 8), 8) Or _
                          mpage(31 + 8)) \ 4

    End If

    ' write speed descriptors supplied?
    If intDescriptors > 0 Then

        ReDim speeds(intDescriptors) As Integer

        ' save CLV descriptors
        For i = 1 To intDescriptors

            ' get descriptor
            CopyMemory udtDescriptor, mpage(28 + 8 + (i * 4)), 4

            ' save speed (in kbytes/s)
            speeds(i - 1) = cd.LShift(udtDescriptor.speed(0), 8) Or udtDescriptor.speed(1)

            ' mark CAV descriptors
            If CBool(udtDescriptor.rotation And &H7) Then
                speeds(i - 1) = speeds(i - 1) Or &H8000
            End If

        Next

    Else

        ' No write speed descriptors
        ' supplied with MMCP.
        ' Simply add write speeds in 4x steps:

        intDescriptors = CDRomGetSpeed(DrvID).MaxWSpeed \ 176
        If intDescriptors > 0 Then
            ReDim speeds(intDescriptors / 4 - 1) As Integer

            For i = 4 To (intDescriptors - 4) Step 4
                speeds((i / 4) - 1) = i * 177
            Next

        Else
            ReDim speeds(0) As Integer
        End If

    End If

    If intDescriptors > 0 Then
        ' add max write speed
        speeds(UBound(speeds)) = CDRomGetSpeed(DrvID).MaxWSpeed
    End If

    ' finished
    CDRomGetWriteSpeeds = True

End Function

'Media locked?
Public Function CDRomIsTrayLocked(ByVal DrvID As String) As Boolean
    Dim mmc As t_MMC

    'read MM Capabilities Page
    If Not CDRomModeSense10(DrvID, &H2A, VarPtr(mmc), Len(mmc) - 1, True, True) Then _
        CDRomIsTrayLocked = -1: Exit Function

    'check "Lock State" Bit
    CDRomIsTrayLocked = Abs(IsBitSet(mmc.misc(2), 1))
End Function

'Tray open or closed?
Public Function CDRomIsTrayOpen(ByVal DrvID As String) As Long

    Dim media_event_req(8) As Byte

    If Not CDRomTestUnitReady(DrvID) Then

        If cd.LastASC = &H3A And cd.LastASCQ = &H1 Then
            CDRomIsTrayOpen = False
        ElseIf cd.LastASC = &H3A And cd.LastASCQ = &H2 Then
            CDRomIsTrayOpen = True

        ' drive doesn't report door status with TUR
        ElseIf cd.LastASC = &H3A And cd.LastASCQ = 0 Then

            'get Tray-Status
            If Not CDRomGetEventStatusNotification(DrvID, &H10, _
                                                   VarPtr(media_event_req(0)), _
                                                   UBound(media_event_req)) Then

                CDRomIsTrayOpen = -1: Exit Function

            End If

            'valid data?
            If media_event_req(0) = 0 And media_event_req(1) = 0 Then
                CDRomIsTrayOpen = -1: Exit Function
            Else
                'Tray open or closed?
                CDRomIsTrayOpen = Abs(IsBitSet(media_event_req(5), 0))
            End If

        End If

    Else

        CDRomIsTrayOpen = False

    End If

End Function

'read a Mode Page
Public Function CDRomModeSense10(ByVal DrvID As String, ByVal MP As Byte, _
                                 ByVal PtrBuffer As Long, _
                                 ByVal BufferLen As Long, _
                                 Optional ByVal DBD As Boolean, _
                                 Optional ByVal CV As Boolean) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H5A                       ' MODE SENSE 10 OpCode
    cmd(1) = Abs(DBD) * &H8             ' Disable Block Descriptors
    cmd(2) = MP Or (Abs(Not CV) * &H80) ' Mode Page (default values)
    cmd(7) = cd.RShift(BufferLen, 8)    ' allocation length
    cmd(8) = BufferLen And &HFF         ' allocation length

    CDRomModeSense10 = cd.ExecCMD(DrvID, cmd, 10, False, _
                                  SRB_DIR_IN, PtrBuffer, BufferLen + 1)

End Function

'load media
Public Function CDRomLoadTray(ByVal DrvID As String) As Boolean

    Dim cmd(6) As Byte

    cmd(0) = &H1B           ' LOUNLOAD OpCode
    cmd(4) = &H3            ' Load Flag

    CDRomLoadTray = cd.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0)

End Function

'eject media
Public Function CDRomUnloadTray(ByVal DrvID As String) As Boolean

    Dim cmd(6) As Byte

    cmd(0) = &H1B           ' LOUNLOAD OpCode
    cmd(4) = &H2            ' Unload Flag

    CDRomUnloadTray = cd.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0)

End Function

'lock media
Public Function CDRomLockMedia(ByVal DrvID As String) As Boolean

    Dim cmd(5) As Byte

    cmd(0) = &H1E           ' LOCK/UNLOCK OpCode
    cmd(4) = 1              ' Lock Flag

    CDRomLockMedia = cd.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0)

End Function

'unlock media
Public Function CDRomUnlockMedia(ByVal DrvID As String) As Boolean

    Dim cmd(5) As Byte

    cmd(0) = &H1E           ' LOCK/UNLOCK OpCode
    cmd(4) = 0              ' remove flags

    CDRomUnlockMedia = cd.ExecCMD(DrvID, cmd, 6, False, SRB_DIR_IN, 0, 0)

End Function

'get Event/Status
Public Function CDRomGetEventStatusNotification(ByVal DrvID As String, _
                                                ByVal Request As Byte, _
                                                ByVal PtrBuffer As Long, _
                                                ByVal BufferLen As Long) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H4A                   ' GET EVENT/STATUS NOTIFICATION OpCode
    cmd(1) = 1                      '
    cmd(4) = Request                ' Request Type
    cmd(8) = BufferLen              ' allocation length

    CDRomGetEventStatusNotification = cd.ExecCMD( _
                                       DrvID, cmd, 10, False, SRB_DIR_IN, _
                                        PtrBuffer, BufferLen _
                                      )

End Function

'get track mode by its sectorcontents
'from CDR-DAO (GPL)
Public Function CDRomGetSectorMode(ByVal DrvID As String, _
                                   ByVal SectorLBA As Long) As e_SectorModes

    Dim Buffer(2351) As Byte
    Dim sync As String, cnt As Integer

    'read sector raw
    If Not CDRomReadCD(DrvID, SectorLBA, 1, VarPtr(Buffer(0)), _
                       UBound(Buffer), &HF8) Then

        CDRomGetSectorMode = -1
        Exit Function

    End If

    'get the first 12 bytes and convert them to hex
    For cnt = 0 To 11
        sync = sync & format(Hex(Buffer(cnt)), "00")
    Next

    'is a sync pattern?
    If sync = SYNCPATTERN Then

        'has to be a data track

        'Byte 16 is the mode
        Select Case Buffer(15)

            Case 1
                'Mode-1 Track (2048 Bytes data)
                CDRomGetSectorMode = MODE_MODE1

            Case 2
                'compare pattern
                If (Buffer(16) = Buffer(20)) And (Buffer(17) = Buffer(21)) _
                    And (Buffer(18) = Buffer(22)) And (Buffer(19) = Buffer(23)) Then

                    If Not Buffer(18) And &H20 = 0 Then
                        'Mode 2 Form 2 Track
                        CDRomGetSectorMode = MODE_MODE2_FORM2
                    Else
                        'Mode 2 Form 1 Track
                        CDRomGetSectorMode = MODE_MODE2_FORM1
                    End If

                Else
                    'Mode-2 Track (2336 Bytes data)
                    CDRomGetSectorMode = MODE_MODE2
                End If

        End Select

    Else
        'Audio or Mode-0 (2352 Bytes data/2352 Bytes completely empty)
        CDRomGetSectorMode = MODE_AUDIO
    End If
End Function

'
Public Function SectorMode2Str(ByVal sMode As e_SectorModes) As String
    Select Case sMode
        Case MODE_AUDIO
            SectorMode2Str = "AUDIO/MODE 0"
        Case MODE_MODE1
            SectorMode2Str = "MODE 1"
        Case MODE_MODE2
            SectorMode2Str = "MODE 2"
        Case MODE_MODE2_FORM1
            SectorMode2Str = "MODE 2 FORM 1"
        Case MODE_MODE2_FORM2
            SectorMode2Str = "MODE 2 FORM 2"
    End Select
End Function

'read sectors
Public Function CDRomReadCD(ByVal DrvID As String, ByVal LBA As Long, _
                            ByVal NumSectors As Long, ByVal PtrBuf As Long, _
                            ByVal buflen As Long, ByVal ReadFlags As e_READCD_FLAGS, _
                            Optional ByVal SubchBits As e_READCD_SUBCH_FLAGS _
                           ) As Boolean

    Dim cmd(11) As Byte

    cmd(0) = &HBE                               ' READ CD OpCode
    cmd(2) = cd.RShift(LBA, 24) And &HFF        ' LBA
    cmd(3) = cd.RShift(LBA, 16) And &HFF
    cmd(4) = cd.RShift(LBA, 8) And &HFF
    cmd(5) = LBA And &HFF                       ' LBA
    cmd(6) = cd.RShift(NumSectors, 16) And &HFF ' num. sectors
    cmd(7) = cd.RShift(NumSectors, 8) And &HFF
    cmd(8) = NumSectors And &HFF                ' num. sectors
    cmd(9) = ReadFlags                          ' read flags
    cmd(10) = SubchBits                         ' Sub-Channel Flags

    CDRomReadCD = cd.ExecCMD(DrvID, cmd, 12, False, _
                             SRB_DIR_IN, PtrBuf, buflen + 1, 40)

End Function

'read sectors
Public Function CDRomReadCDMSF(ByVal DrvID As String, ByVal startM As Byte, _
                            ByVal startS As Byte, ByVal startF As Byte, _
                            ByVal endM As Byte, ByVal endS As Byte, ByVal endF As Byte, _
                            ByVal PtrBuf As Long, ByVal buflen As Long, ByVal ReadFlags As e_READCD_FLAGS, _
                            Optional ByVal SubchBits As e_READCD_SUBCH_FLAGS _
                           ) As Boolean

    Dim cmd(11) As Byte

    cmd(0) = &HB9                               ' READ CD OpCode
    cmd(3) = startM                             ' minutes
    cmd(4) = startS                             ' seconds
    cmd(5) = startF                             ' frames
    cmd(6) = endM                               ' minutes
    cmd(7) = endS                               ' seconds
    cmd(8) = endF                               ' frames
    cmd(9) = ReadFlags                          ' read flags
    cmd(10) = SubchBits                         ' Sub-Channel Flags

    CDRomReadCDMSF = cd.ExecCMD(DrvID, cmd, 12, False, _
                             SRB_DIR_IN, PtrBuf, buflen + 1, 10)

End Function

'read Mode 1 sectors
Public Function CDRomRead10(ByVal DrvID As String, ByVal LBA As Long, _
                            ByVal PtrBuf As Long, ByVal buflen As Long, _
                            ByVal Blocks As Long) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H28                           ' READ10 Op-Code
    cmd(2) = cd.RShift(LBA, 24) And &HFF    ' LBA
    cmd(3) = cd.RShift(LBA, 16) And &HFF
    cmd(4) = cd.RShift(LBA, 8) And &HFF
    cmd(5) = LBA And &HFF                   ' LBA
    cmd(7) = Blocks \ &HFF                  ' num. sectors
    cmd(8) = Blocks And &HFF                ' num. sectors

    CDRomRead10 = cd.ExecCMD(DrvID, cmd, 10, False, SRB_DIR_IN, _
                             PtrBuf, buflen)

End Function

'return the number of tracks present on a disc
Public Function CDRomTrackCount(ByVal DriveID As String, _
                                ByRef intTrks As Integer) As Boolean

    Dim toc As t_TOC_STRUCT

    'read simple TOC
    If Not CDRomReadTOC(DriveID, 0, True, 0, VarPtr(toc), Len(toc) - 1) Then _
        Exit Function

    'return the number of tracks
    intTrks = toc.LastTrack

    'success?
    CDRomTrackCount = intTrks > 0

End Function

'read track info
Public Function CDRomTrackInfo(ByVal DriveID As String, ByVal Track As Integer, _
                               ByRef nfo As t_TrackInfo) As Boolean

    Dim sectors As Long, LeadOut As Long, LastTrk As Long
    Dim toc As t_RTOC_STRUCT
    Dim i As Integer

    'read raw TOC
    If Not CDRomReadTOC(DriveID, 2, True, 1, _
                        VarPtr(toc), Len(toc) - 1) Then Exit Function

    'go through all packets
    For i = 0 To ((cd.LShift(toc.dummy(0), 8) Or toc.dummy(1)) \ 11) - 1

        'determine packet type
        Select Case toc.packet(i).point

            'found the chosen track?
            Case Is = Track

                'found the track, save its number
                nfo.Track = Track
                'Sessionnumber
                nfo.Session = toc.packet(i).sessionNr

                'Startaddress
                nfo.startLBA = cd.MSF2LBA(toc.packet(i).pmin, _
                                          toc.packet(i).psec, _
                                          toc.packet(i).pframe)

                'last track in the current session?
                If toc.packet(i).point = LastTrk Then

                    'yes, last LBA is Lead-Out startaddress
                    nfo.endLBA = LeadOut - 150
                    nfo.LastTrackInSession = True

                Else

                    'else normal end address
                    nfo.endLBA = cd.MSF2LBA(toc.packet(i + 1).pmin, _
                                            toc.packet(i + 1).psec, _
                                            toc.packet(i + 1).pframe)

                End If

                'tracklength in sectors
                nfo.Length = nfo.endLBA - nfo.startLBA

                'determine data mode
                nfo.DataMode = CDRomGetSectorMode(DriveID, nfo.startLBA)

                Exit For

            'found packet: "last track in current session"
            Case &HA1
                LastTrk = toc.packet(i).pmin

            'found packet: "Lead-Out of the current session"
            Case &HA2
                LeadOut = cd.MSF2LBA(toc.packet(i).pmin, _
                                     toc.packet(i).psec, _
                                     toc.packet(i).pframe)

        End Select

    Next

    CDRomTrackInfo = True
End Function

'read TOC/PMA/ATIP/CD-TEXT
Public Function CDRomReadTOC(ByVal DrvID As String, ByVal TOC_Format As Integer, _
                             ByVal MSF As Boolean, ByVal Track_Session As Integer, _
                             ByVal PtrBuffer As Long, ByVal BufferLen As Long _
                            ) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H43                           ' READ TOC OpCode
    cmd(1) = IIf(MSF, &H2, 0)               ' MSF or LBA?
    cmd(2) = TOC_Format                     ' Format (TOC, PMA, ATIP, CD-Text)
    cmd(6) = Track_Session                  ' Track/Session
    cmd(7) = BufferLen \ &HFF               ' allocation length
    cmd(8) = BufferLen And &HFF             ' allocation length

    CDRomReadTOC = cd.ExecCMD(DrvID, cmd, 10, False, _
                              SRB_DIR_IN, PtrBuffer, BufferLen)

End Function

'deactivate error correction
Public Function CDRomDeactivateECC(ByVal DrvID As String) As Boolean
    Dim page(16)

    page(8) = &H1
    page(9) = &H6
    page(10) = &H11

    CDRomDeactivateECC = CDRomModeSelect10(DrvID, VarPtr(page(0)), UBound(page))
End Function

'activate error correction
Public Function CDRomActivateECC(ByVal DrvID As String) As Boolean
    Dim page(16)

    page(8) = &H1
    page(9) = &H6
    page(10) = &H1

    CDRomActivateECC = CDRomModeSelect10(DrvID, VarPtr(page(0)), UBound(page))
End Function

'send a Mode Page
Public Function CDRomModeSelect10(ByVal DrvID As String, ByVal PtrBuffer As Long, _
                                  ByVal BufferLen As Long) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H55                   ' MODE SELECT10 OpCode
    cmd(1) = &H10                   ' PF = 1 (Page Format)
    cmd(7) = BufferLen \ &HFF       ' allocation length
    cmd(8) = BufferLen Mod &HFF     ' allocation length

    CDRomModeSelect10 = cd.ExecCMD(DrvID, cmd, 10, True, _
                                   SRB_DIR_OUT, PtrBuffer, BufferLen)

End Function

'collection drive information
Public Function CDRomGetLWInfo(ByVal strDrvID As String) As t_DrvInfo

    Dim mmc As t_MMC

    'read Multimedia Capabilities Mode Page
    CDRomModeSense10 strDrvID, &H2A, VarPtr(mmc), Len(mmc) - 1, True, True

    'read read features
    With CDRomGetLWInfo.ReadFeatures
        .CDR = IsBitSet(mmc.ReadSupported, 0)
        .CDRW = IsBitSet(mmc.ReadSupported, 1)
        .DVDROM = IsBitSet(mmc.ReadSupported, 3)
        .DVDR = IsBitSet(mmc.ReadSupported, 4)
        .DVDRAM = IsBitSet(mmc.ReadSupported, 5)
        .CDDARawRead = IsBitSet(mmc.misc(1), 0)
        .Mode2Form1 = IsBitSet(mmc.misc(0), 4)
        .Mode2Form2 = IsBitSet(mmc.misc(0), 5)
        .Multisession = IsBitSet(mmc.misc(0), 6)
        .ISRC = IsBitSet(mmc.misc(1), 5)
        .UPC = IsBitSet(mmc.misc(1), 6)
        .BC = IsBitSet(mmc.misc(1), 7)
        .subchannels = IsBitSet(mmc.misc(1), 2)
        .SubChannelsCorrected = IsBitSet(mmc.misc(1), 3)
        .SubChannelsFormLeadIn = IsBitSet(mmc.misc(3), 5)
        .C2ErrorPointers = IsBitSet(mmc.misc(1), 4)
    End With

    'read write features
    With CDRomGetLWInfo.WriteFeatures
        .CDR = IsBitSet(mmc.WriteSupported, 0)
        .CDRW = IsBitSet(mmc.WriteSupported, 1)
        .TestMode = IsBitSet(mmc.WriteSupported, 2)
        .DVDR = IsBitSet(mmc.WriteSupported, 4)
        .DVDRAM = IsBitSet(mmc.WriteSupported, 5)
        .BURNProof = IsBitSet(mmc.misc(0), 7)

        If CDRomWriteParams(strDrvID, False, False, 150, 1, 0, 0, False) Then _
            .WriteModes.TAO = True

        If CDRomWriteParams(strDrvID, True, False, 150, 1, 0, 0, False) Then _
            .WriteModes.TAOTest = True

        If CDRomWriteParams(strDrvID, False, False, 150, 2, 0, 8, False) Then _
            .WriteModes.SAO = True

        If CDRomWriteParams(strDrvID, True, False, 150, 2, 0, 8, False) Then _
            .WriteModes.SAOTest = True

        If CDRomWriteParams(strDrvID, False, False, 150, 3, 0, 1, False) Then _
            .WriteModes.Raw16 = True

        If CDRomWriteParams(strDrvID, True, False, 150, 3, 0, 1, False) Then _
            .WriteModes.Raw16Test = True

        If CDRomWriteParams(strDrvID, False, False, 150, 3, 0, 3, False) Then _
            .WriteModes.Raw96 = True

        If CDRomWriteParams(strDrvID, True, False, 150, 3, 0, 3, False) Then _
            .WriteModes.Raw96Test = True
    End With

    'generic information
    With CDRomGetLWInfo
        .speeds = CDRomGetSpeed(strDrvID)
        .Interface = CDRomGetInterface(strDrvID)
        .LockMedia = IsBitSet(mmc.misc(2), 0)
        .AnalogAudio = IsBitSet(mmc.misc(0), 0)
        .JitterCorrection = IsBitSet(mmc.misc(1), 1)
        .BufferSize = cd.LShift(mmc.BufferSize(0), 8) Or mmc.BufferSize(1)

        If IsBitSet(mmc.misc(2), 5) = False And _
           IsBitSet(mmc.misc(2), 6) = False And _
           IsBitSet(mmc.misc(2), 7) = False Then

            .LoadingMechanism = LOAD_CADDY

        ElseIf IsBitSet(mmc.misc(2), 5) And _
               IsBitSet(mmc.misc(2), 6) = False And _
               IsBitSet(mmc.misc(2), 7) = False Then

            .LoadingMechanism = LOAD_TRAY

        ElseIf IsBitSet(mmc.misc(2), 5) = False And _
              IsBitSet(mmc.misc(2), 6) And _
              IsBitSet(mmc.misc(2), 7) = False Then

            .LoadingMechanism = LOAD_POPUP

        ElseIf IsBitSet(mmc.misc(2), 5) = False And _
               IsBitSet(mmc.misc(2), 6) = False And _
               IsBitSet(mmc.misc(2), 7) Then

            .LoadingMechanism = LOAD_CHANGER

        ElseIf IsBitSet(mmc.misc(2), 5) And _
               IsBitSet(mmc.misc(2), 6) = False And _
               IsBitSet(mmc.misc(2), 7) Then

            .LoadingMechanism = LOAD_CHANGER

        Else

            .LoadingMechanism = LOAD_UNKNWN

        End If
    End With

End Function

'read read- and writespeeds
Public Function CDRomGetSpeed(ByVal strDrv As String) As t_Speed
    Dim cmd(9) As Byte
    Dim mmc As t_MMC

    CDRomModeSense10 strDrv, &H2A, VarPtr(mmc), Len(mmc) - 1, True, True

    With CDRomGetSpeed
        .MaxRSpeed = cd.LShift(mmc.MaxReadSpeed(0), 8) Or mmc.MaxReadSpeed(1)
        .CurrRSpeed = cd.LShift(mmc.CurrReadSpeed(0), 8) Or mmc.CurrReadSpeed(1)
        .MaxWSpeed = cd.LShift(mmc.MaxWriteSpeed(0), 8) Or mmc.MaxWriteSpeed(1)

        ' MMC 3/4 write speed?
        .CurrWSpeed = cd.LShift(mmc.CurrWriteSpeedMMC3(0), 8) Or mmc.CurrWriteSpeedMMC3(1)
        If .CurrWSpeed = 0 Then
            ' no, take the MMC 1/2 one
            .CurrWSpeed = cd.LShift(mmc.CurrWriteSpeed(0), 8) Or mmc.CurrWriteSpeed(1)
        End If
    End With
End Function

'read track information
Public Function CDRomReadTrackInformation(ByVal DrvID As String, _
                                          ByVal AdrType As Byte, _
                                          ByVal Track As Byte, _
                                          ByVal PtrBuffer As Long, _
                                          ByVal BufferLen As Long _
                                         ) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H52                             ' READ TRACK INFORMATION OpCode
    cmd(1) = AdrType                          ' address type
    cmd(2) = cd.RShift(Track, 24) And &HFF    ' Track/LBA/Session
    cmd(3) = cd.RShift(Track, 16) And &HFF    ' Track/LBA/Session
    cmd(4) = cd.RShift(Track, 8) And &HFF     ' Track/LBA/Session
    cmd(5) = Track And &HFF                   ' Track/LBA/Session
    cmd(7) = BufferLen \ &HFF                 ' allocation length
    cmd(8) = BufferLen And &HFF               ' allocation length

    CDRomReadTrackInformation = cd.ExecCMD(DrvID, cmd, 10, False, _
                                           SRB_DIR_IN, PtrBuffer, _
                                           BufferLen)

End Function

' send cue sheet for SAO
Public Function CDRomSendCueSheet(ByVal DrvID As String, _
                ByVal PtrBuf As Long, _
                ByVal buflen As Long) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H5D
    cmd(6) = cd.RShift(buflen, 16) And &HFF
    cmd(7) = cd.RShift(buflen, 8) And &HFF
    cmd(8) = buflen And &HFF

    CDRomSendCueSheet = cd.ExecCMD(DrvID, cmd, 10, _
                                 True, SRB_DIR_OUT, _
                                  PtrBuf, buflen)

End Function

'write buffer completely to disc
Public Function CDRomSyncCache(ByVal DrvID As String) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H35                       ' SYNCHRONIZE CACHE OpCode

    CDRomSyncCache = cd.ExecCMD(DrvID, cmd, 10, True, SRB_DIR_OUT, 0, 0)

End Function

'close Track/Session/CD
Public Function CDRomCloseCD(ByVal DrvID As String, _
                             ByVal CloseFunction As Byte, _
                             ByVal TrackNum As Integer) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H5B                       ' CLOSE TRACK/SESSION OpCode
    cmd(2) = CloseFunction              ' close function
    cmd(5) = TrackNum                   ' Track/Session Number

    CDRomCloseCD = cd.ExecCMD(DrvID, cmd, 10, True, SRB_DIR_OUT, 0, 0)

End Function

'write data to disc
Public Function CDRomWrite10(ByVal DrvID As String, _
                             ByVal LBA As Long, _
                             ByVal WrtSektors As Long, _
                             ByVal PtrBuffer As Long, _
                             ByVal BufferLen As Long) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H2A                           ' WRITE10 OpCode
    cmd(2) = cd.RShift(LBA, 24) And &HFF    ' LBA
    cmd(3) = cd.RShift(LBA, 16) And &HFF    ' LBA
    cmd(4) = cd.RShift(LBA, 8) And &HFF     ' LBA
    cmd(5) = LBA And &HFF                   ' LBA
    cmd(7) = WrtSektors \ &HFF              ' allocation length
    cmd(8) = WrtSektors And &HFF            ' allocation length

    CDRomWrite10 = cd.ExecCMD(DrvID, cmd, 10, True, SRB_DIR_OUT, _
                              PtrBuffer, BufferLen + 1)

End Function

'read buffer size
Public Function CDRomBufferCapacity(ByVal strDrvID As String, _
                                    ByVal BufPtr As Long, _
                                    ByVal buflen As Long) As Long

    Dim cmd(9)  As Byte

    cmd(0) = &H5C       ' READ BUFFER CAPACITY OpCode
    cmd(8) = buflen     ' allocation length

    CDRomBufferCapacity = cd.ExecCMD(strDrvID, cmd, 10, False, SRB_DIR_IN, _
                                     BufPtr, buflen)
End Function

'forever try to write to the drive buffer,
'to avoid buffer overflows.
'also look for buffer underruns.
Public Function CDRomBurnCD(ByVal DriveID As String, _
                      ByVal LBA As Long, ByVal sectors As Long, _
                      ByVal PtrBuf As Long, ByVal buflen As Long _
                     ) As Boolean

    Const retries   As Long = 20

    '   buffer structure
    Dim bufcap      As t_BufferCapacity
    '   buffer length        empty part of the buffer
    Dim BufferLen   As Long, BufferBlank    As Long

    Dim lngRetries  As Long

    'as long as writing doesn't work
    Do While Not CDRomWrite10(DriveID, LBA, sectors, PtrBuf, buflen)

        ' write error :(
        If cd.LastSK = 3 And cd.LastASC = 12 Then
            Exit Function
        End If

        If lngRetries = retries Then
            ' failed...
            Debug.Print "Write10 failed", _
                        "LBA: " & LBA, _
                        "Sectors: " & sectors, _
                        "Writelen: " & buflen

            Exit Function
        End If

        'read buffer stats
        If Not CDRomBufferCapacity(DriveID, VarPtr(bufcap), _
                                   Len(bufcap) - 1) Then

            Exit Function

        Else

            'buffer len
            BufferLen = cd.LShift(bufcap.BufferLen(0), 24) Or _
                        cd.LShift(bufcap.BufferLen(1), 16) Or _
                        cd.LShift(bufcap.BufferLen(2), 8) Or _
                                  bufcap.BufferLen(3)

            'empty part of the buffer
            BufferBlank = cd.LShift(bufcap.BufferBlank(0), 24) Or _
                          cd.LShift(bufcap.BufferBlank(1), 16) Or _
                          cd.LShift(bufcap.BufferBlank(2), 8) Or _
                                    bufcap.BufferBlank(3)

            'check for invalid data or buffer underruns
            If BufferBlank = BufferLen Or _
               BufferLen = 0 Or _
               BufferBlank > BufferLen Then

                'oh oh
                Exit Function

            End If

        End If

        Sleep 50
        lngRetries = lngRetries + 1

        DoEvents

    Loop

    'success
    CDRomBurnCD = True

End Function

'send new write parameters page
Public Function CDRomWriteParams(ByVal strDrv As String, ByVal TestMode As Boolean, _
                                 ByVal BURNProof As Boolean, ByVal TrackPause As Long, _
                                 ByVal WriteType As Byte, ByVal TrackMode As Byte, _
                                 ByVal DataBlockType As Byte, ByVal Multisession As Boolean _
                                ) As Boolean

    Dim bufData(60) As Byte
    Dim PS As Byte
    Dim i As Integer

    'read the page
    'CDRomModeSense10 strDrv, &H5, VarPtr(bufData(0)), UBound(bufData)

    bufData(1) = 58                           ' length

    bufData(8) = &H5                          ' Page Code
    bufData(9) = &H32                         ' Page length

    'WriteType, Test-Mode and Burn Proof
    bufData(10) = WriteType Or _
                  Abs(TestMode) * &H10 Or _
                  Abs(BURNProof) * &H40

    'Track-Mode, Multi-Session
    bufData(11) = TrackMode Or Abs(Multisession) * &HC0

    'Data-Mode (Mode 1: 2048 Bytes User-Data)
    bufData(12) = DataBlockType

    'Track Pause: 150 sectors (frames) = 2 seconds
    bufData(22) = cd.RShift(TrackPause, 8) And &HFF
    bufData(23) = TrackPause And &HFF

    'send new WPP
    CDRomWriteParams = CDRomModeSelect10(strDrv, VarPtr(bufData(0)), 60)
End Function

'gets the interface of a drive
Public Function CDRomGetInterface(ByVal strDrv As String) As e_DrvInterfaces
    Dim Buffer(15) As Byte, cmd(9) As Byte
    Dim inquiry As t_InqDat

    'try to read the drive's core feature
    CDRomGetConfiguration strDrv, 1, 2, VarPtr(Buffer(0)), UBound(Buffer)

    'determine the interface from it
    Select Case cd.LShift(Buffer(12), 24) Or cd.LShift(Buffer(13), 16) Or _
                cd.LShift(Buffer(14), 8) Or Buffer(15)

        Case 1: CDRomGetInterface = IF_SCSI
        Case 2, 7: CDRomGetInterface = IF_ATAPI
        Case 3, 4, 6: CDRomGetInterface = IF_IEEE
        Case 8: CDRomGetInterface = IF_USB
        Case Else: CDRomGetInterface = IF_UNKNWN

    End Select

    'if it didn't work, try INQUIRY
    If CDRomGetInterface = IF_UNKNWN Then

        cmd(0) = &H12               ' Inquiry OpCode
        cmd(4) = Len(inquiry) - 1   ' allocation length

        If cd.ExecCMD(strDrv, cmd, 10, False, SRB_DIR_IN, _
                      VarPtr(inquiry), Len(inquiry), 10) Then

            If inquiry.rsv1(0) = 0 Then CDRomGetInterface = IF_ATAPI

        End If

    End If
End Function

'read drive features
Public Function CDRomGetConfiguration(ByVal strDrv As String, ByVal StartFeature As Byte, _
                                      ByVal RT As Byte, ByVal PtrBuffer As Long, _
                                      ByVal buflen As Long) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H46                   ' GET CONFIGURATION Op-Code
    cmd(1) = RT                     ' RT Byte
    cmd(2) = StartFeature \ &HFF    ' startfeature
    cmd(3) = StartFeature And &HFF  ' startfeature
    cmd(7) = buflen \ &HFF          ' allocation length
    cmd(8) = buflen And &HFF        ' allocation length

    CDRomGetConfiguration = cd.ExecCMD(strDrv, cmd, 10, False, _
                                       SRB_DIR_IN, PtrBuffer, buflen + 1, 10)
End Function

'CD-Text available?
Public Function CDRomHasCDText(ByVal DrvID As String) As Boolean

    On Error GoTo ErrHandle

    Dim str  As String

    str = CDRomReadCDText(DrvID)(0)
    CDRomHasCDText = True

ErrHandle:

End Function

'read CD-Text
Public Function CDRomReadCDText(ByVal DrvID As String) As String()
    On Error Resume Next

    Dim albnameready As Boolean, IsEmpty As Boolean     '
    Dim trknames(99) As String, RetArr() As String      ' name buffer
    Dim artistname As String, albname As String         ' artist & album
    Dim trkindex As Integer, trkname As String          ' curr track idx & name
    Dim albyear As String, albgenre As String           ' year and genre
    Dim j As Integer, i As Integer                      ' counter
    Dim tCDT As t_CD_TEXT                               ' CD-Text Container

    'try to read CD-Text
    If Not CDRomReadTOC(DrvID, &H5, True, 0, VarPtr(tCDT), Len(tCDT) - 1) Then _
        Exit Function

    'go through every packet
    For j = 0 To 255
        If tCDT.CDText(j).idSeq <> j Then Exit For

            ' we don't want unicode
            If ((tCDT.CDText(j).idFlg And &H30) = 0) Then

                'packet type
                Select Case tCDT.CDText(j).idType

                    'album or track name
                    Case CD_TEXT_PACK_ALBUM_NAME
                        If (tCDT.CDText(j).idTrk = 0) Then
                            For i = 0 To 11
                                If (tCDT.CDText(j).txt(i) <> 0) Then
                                  If albnameready = False Then
                                    'album
                                    albname = albname & Chr$(tCDT.CDText(j).txt(i))
                                  Else
                                    'track
                                    trkname = trkname & Chr$(tCDT.CDText(j).txt(i))
                                  End If
                                Else
                                  Trim$ albname
                                  albnameready = True
                                End If
                            Next i

                        Else
                            'trackname
                            For i = 0 To 11
                                If (tCDT.CDText(j).txt(i) <> 0) Then
                                  trkname = trkname & Chr$(tCDT.CDText(j).txt(i))
                                Else
                                  Trim$ (trkname)
                                  trknames(trkindex) = trkname
                                  trkname = vbNullString
                                  trkindex = trkindex + 1
                                  If (tCDT.CDText(j).txt(i + 1) = 0) Then Exit For
                                End If
                            Next
                        End If

                    'artistname
                    Case CD_TEXT_PACK_PERFORMER
                        If (tCDT.CDText(j).idTrk = 0) Then
                            For i = 0 To 11
                                If (tCDT.CDText(j).txt(i) <> 0) Then
                                  artistname = artistname & Chr$(tCDT.CDText(j).txt(i))
                                Else
                                  Trim$ artistname
                                  Exit For
                                End If
                            Next
                        End If

                    'Albumgenre
'                    Case CD_TEXT_PACK_GENRE
'                        If (tCDT.CDText(J).idTrk = 0) Then
'                            For i = 0 To 11
'                                albgenre = albgenre & Chr$(tCDT.CDText(J).txt(i))
'                            Next
'                        End If

                    ' never had CD-Text with these types
                    Case CD_TEXT_PACK_ARRANGER
                        Debug.Print "Arranger"
                    Case CD_TEXT_PACK_COMPOSER
                        Debug.Print "Composer"
                    Case CD_TEXT_PACK_MESSAGES
                        Debug.Print "Messages"
                    Case CD_TEXT_PACK_SIZE_INFO
                        Debug.Print "Size info"
                    Case CD_TEXT_PACK_SONGWRITER
                        Debug.Print "Songwriter"
                    Case CD_TEXT_PACK_UPC_EAN
                        Debug.Print "UPC/EAN"

                End Select
            End If
    Next j

    RemNulls albgenre
    If Len(albgenre) > 0 Then albgenre = Asc(Left(albgenre, 1))

    'collected CD-Text data empty?
    IsEmpty = True
    If artistname = vbNullString And albname = vbNullString Then

        'check every track
        For i = LBound(trknames) To UBound(trknames)

            'if not empty, data exists
            If Not trknames(i) = vbNullString Then
                IsEmpty = False
                Exit For
            End If

        Next i

    Else
        'if album and artist are there, data exists
        IsEmpty = False
    End If

    If IsEmpty Then Exit Function

    'prepare array to return data
    ReDim RetArr(0 To UBound(trknames) + 2)

    RetArr(0) = albname
    RetArr(1) = artistname
    'RetArr(2) = albgenre

    For i = LBound(trknames) To UBound(trknames)
        RetArr(i + 2) = Trim$(trknames(i))
    Next

    'return the array
    CDRomReadCDText = RetArr
End Function

'calculate CDDB ID
Public Function CDRomCDDB_ID(ByVal DrvID As String) As String

    Dim tTOC    As t_TOC_STRUCT
    Dim i       As Long, t As Long, n As Long

    'read simple TOC
    CDRomReadTOC DrvID, 0, True, 0, VarPtr(tTOC), Len(tTOC) - 1

    'sum tracks
    For i = 0 To tTOC.LastTrack - 1
        n = UnsignedAdd( _
                n, CDDB_Sum( _
                    (tTOC.TocTrack(i).addr(1) * 60) _
                     + tTOC.TocTrack(i).addr(2) _
                   ) _
            )
    Next i

    'size of the disc
    t = ((tTOC.TocTrack(tTOC.LastTrack).addr(1) * 60) _
         + tTOC.TocTrack(tTOC.LastTrack).addr(2)) _
         - ((tTOC.TocTrack(0).addr(1) * 60) _
         + tTOC.TocTrack(0).addr(2))

    'create a CDDB ID from it
    CDRomCDDB_ID = lPAD(Hex$(cd.LShift((n Mod &HFF), 24) Or _
                             cd.LShift(t, 8) Or _
                             tTOC.LastTrack), _
                        8, _
                        0)
End Function

'CDDB sum
Private Function CDDB_Sum(ByVal n As Integer) As Integer
    Dim ret As Long
    ret = 0

    While n > 0
        ret = UnsignedAdd(ret, n Mod 10)
        n = n \ 10
    Wend

    CDDB_Sum = ret
End Function

'blanks a CD-RW
Public Function CDRomEraseCDRW(ByVal strDrvID As String, _
                ByVal BlankType As e_BLANKTYPE, _
                Optional sync As Boolean = True) As Boolean

    Dim cmd(11) As Byte

    cmd(0) = &HA1                               ' BLANK Op-Code
    cmd(1) = BlankType Or Abs(Not sync) * &H10  ' Löschfunktion

    CDRomEraseCDRW = cd.ExecCMD(strDrvID, cmd, 12, False, SRB_DIR_OUT, 0, 0, 5000)

End Function

'read the simple TOC of a disc
Public Function CDRomGetTOC(ByVal strDrv As String) As t_Tracks
    Dim toc As t_TOC_STRUCT
    Dim i As Byte

    If Not CDRomTestUnitReady(strDrv) Then Exit Function

    'read simple TOC
    CDRomReadTOC strDrv, 0, True, 0, VarPtr(toc), Len(toc) - 1

    With CDRomGetTOC

        .Tracks = toc.LastTrack

        If .Tracks > 0 Then

        For i = 0 To toc.LastTrack - 1
            'Track Number
            .Track(i).TrkNum = i + 1
            'Audio- or Datatrack?
            .Track(i).AudioTrack = (IsBitSet(toc.TocTrack(i).ADR, 2) = False)
            'Track Start address
            .Track(i).Start.M = toc.TocTrack(i).addr(1)
            .Track(i).Start.s = toc.TocTrack(i).addr(2)
            .Track(i).Start.F = toc.TocTrack(i).addr(3)
            .Track(i).Start.LBA = cd.MSF2LBA(.Track(i).Start.M, .Track(i).Start.s, .Track(i).Start.F)
            'Track End address
            .Track(i).end.M = toc.TocTrack(i + 1).addr(1)
            .Track(i).end.s = toc.TocTrack(i + 1).addr(2)
            .Track(i).end.F = toc.TocTrack(i + 1).addr(3)
            .Track(i).end.LBA = cd.MSF2LBA(.Track(i).end.M, .Track(i).end.s, .Track(i).end.F)
        Next

        End If

    End With

End Function

'set a new read an write speed
'&HFFFF& = max. speed
Public Function CDRomSetCDSpeed(ByVal strDrv As String, _
                                ByVal NewReadSpeed As Long, _
                                ByVal NewWriteSpeed As Long, _
                                ByVal CAV As Boolean) As Boolean

    Dim cmd(11) As Byte

    If NewReadSpeed > &HFFFF& Then NewReadSpeed = &HFFFF&
    If NewWriteSpeed > &HFFFF& Then NewWriteSpeed = &HFFFF&

    cmd(0) = &HBB                           ' SET CD SPEED Op-Code
    cmd(1) = Abs(CAV)                       ' CAV rotation?

    If NewReadSpeed < &HFFFF& Then
        cmd(2) = NewReadSpeed \ &HFF        ' NewReadSpeed MSB
        cmd(3) = NewReadSpeed Mod &HFF      ' NewReadSpeed LSB
    Else
        cmd(2) = &HFF                       ' max read speed MSB
        cmd(3) = &HFF                       ' max read speed LSB
    End If

    If NewWriteSpeed < &HFFFF& Then
        cmd(4) = NewWriteSpeed \ &HFF       ' NewWriteSpeed MSB
        cmd(5) = NewWriteSpeed Mod &HFF     ' NewWriteSpeed LSB
    Else
        cmd(4) = &HFF                       ' max write speed MSB
        cmd(5) = &HFF                       ' max write speed LSB
    End If

    CDRomSetCDSpeed = cd.ExecCMD(strDrv, cmd, 12, False, SRB_DIR_OUT, 0, 0)
End Function

'collects some information about the inserted disc
Public Function CDRomGetCDInfo(ByVal strDrv As String) As t_CDInfo

    Dim BufAtip         As t_ATIP
    Dim BufRDI          As t_RDI

    Dim sBuf            As String
    Dim conf_hdr(512)   As Byte


    'get some information
    CDRomReadDiscInformation strDrv, VarPtr(BufRDI), Len(BufRDI) - 1
    CDRomReadTOC strDrv, 4, True, 0, VarPtr(BufAtip), Len(BufAtip) - 1

    With CDRomGetCDInfo

        'Lead-In start time
        .LeadIn.M = BufAtip.LeadIn_Min
        .LeadIn.s = BufAtip.LeadIn_Sec
        .LeadIn.F = BufAtip.LeadIn_Frm
        .LeadIn.LBA = cd.MSF2LBA(.LeadIn.M, .LeadIn.s, .LeadIn.F)

        'Lead-Out start time
        .LeadOut.M = BufAtip.LeadOut_Min
        .LeadOut.s = BufAtip.LeadOut_Sec
        .LeadOut.F = BufAtip.LeadOut_Frm
        .LeadOut.LBA = cd.MSF2LBA(.LeadOut.M, .LeadOut.s, .LeadOut.F)

        'CD Status
        If IsBitSet(BufRDI.states, 1) = False And _
           IsBitSet(BufRDI.states, 0) = False Then

            .DiscStatus = STAT_EMPTY

        ElseIf IsBitSet(BufRDI.states, 1) = False And _
               IsBitSet(BufRDI.states, 0) Then

            .DiscStatus = STAT_INCOMPLETE

        ElseIf IsBitSet(BufRDI.states, 1) And _
               IsBitSet(BufRDI.states, 0) = False Then

            .DiscStatus = STAT_COMPLETE

        Else

            .DiscStatus = STAT_UNKNWN

        End If

        'last session status
        If Not IsBitSet(BufRDI.states, 3) And _
           Not IsBitSet(BufRDI.states, 2) Then

            .LastSessionStatus = STAT_EMPTY

        ElseIf Not IsBitSet(BufRDI.states, 3) And _
                   IsBitSet(BufRDI.states, 2) Then

            .LastSessionStatus = STAT_INCOMPLETE

        ElseIf IsBitSet(BufRDI.states, 3) And _
               IsBitSet(BufRDI.states, 2) Then

            .LastSessionStatus = STAT_COMPLETE

        Else

            .DiscStatus = STAT_UNKNWN

        End If

        'CD Type
        .CDType = CDRomGetCDType(strDrv)

        'CD Sub-Type
        If BufRDI.DiscType = &H0 Then
            .CDSubType = STYPE_CDROMDA
        ElseIf BufRDI.DiscType = &H10 Then
            .CDSubType = STYPE_CDI
        ElseIf BufRDI.DiscType = &H20 Then
            .CDSubType = STYPE_XA
        Else
            .CDSubType = STYPE_UNKNWN
        End If

        'Erasable?
        .Erasable = IsBitSet(BufRDI.states, 4)

        'CD-R(W) Vendor
        .Vendor = CDRomGetCDRWVendor(strDrv)

        'Sessions
        .Sessions = cd.LShift(BufRDI.NumSessionsMSB, 8) Or _
                              BufRDI.NumSessionsLSB

        'Tracks
        .Tracks = cd.LShift(BufRDI.LastTrackLastSessionMSB, 8) Or _
                            BufRDI.LastTrackLastSessionLSB

        'capacity in bytes (Mode 1)
        .Capacity = .LeadOut.LBA * 2048&
        If .Capacity < 0 Then .Capacity = 0

        'used part of the disc in bytes
        .Size = CDRomGetUsedBytes(strDrv)
    End With
End Function

'Warning: the written sectors will be multiplied with 2048.
'         Mode 2 or DA not supported.
Private Function CDRomGetUsedBytes(ByVal strDrv As String) As Double

    Dim cap     As t_ReadCap
    Dim cmd(9)  As Byte

    cmd(0) = &H25                     ' READ CAPACITY Op-Code

    If Not cd.ExecCMD(strDrv, cmd, 10, False, SRB_DIR_IN, _
                      VarPtr(cap), Len(cap), 10) Then

        'failed
        CDRomGetUsedBytes = -1

    Else

        'return written sectors
        CDRomGetUsedBytes = CDbl(cd.LShift(cap.Blocks(0), 24) Or _
                                 cd.LShift(cap.Blocks(1), 16) Or _
                                 cd.LShift(cap.Blocks(2), 8) Or _
                                 cap.Blocks(3)) _
                                * 2048#

    End If
End Function

'read spin down speed
'from CDROM TOOL (GPL)
Public Function CDRomGetSpinDown(ByVal strDrvID As String) As Integer
    Dim mpage(255) As Byte

    CDRomModeSense10 strDrvID, &HD, VarPtr(mpage(0)), UBound(mpage), True, True

    CDRomGetSpinDown = mpage(11)
End Function

'set new Spin Down speed
'from CDROM TOOL (GPL)
Public Function CDRomSetSpinDown(ByVal strDrvID As String, ByVal spindown As e_SpinDown) As Boolean
    Dim mpage(15) As Byte

    mpage(8) = &HD
    mpage(9) = &H6
    mpage(11) = spindown
    mpage(13) = &H3C
    mpage(15) = &H4B

    CDRomSetSpinDown = CDRomModeSelect10(strDrvID, VarPtr(mpage(0)), UBound(mpage) + 1)
End Function

Public Function CDRomGetIdleTimer(ByVal strDrvID As String) As Long

    Dim mpage(19) As Byte

    If CDRomModeSense10(strDrvID, &H1A, VarPtr(mpage(0)), UBound(mpage), True, True) Then

        CDRomGetIdleTimer = cd.LShift(mpage(12), 24) Or _
                            cd.LShift(mpage(13), 16) Or _
                            cd.LShift(mpage(14), 8) Or _
                            mpage(15)

    End If

End Function

Public Function CDRomGetStandbyTimer(ByVal strDrvID As String) As Long

    Dim mpage(19) As Byte

    If CDRomModeSense10(strDrvID, &H1A, VarPtr(mpage(0)), UBound(mpage), True, True) Then

        CDRomGetStandbyTimer = cd.LShift(mpage(16), 24) Or _
                            cd.LShift(mpage(17), 16) Or _
                            cd.LShift(mpage(18), 8) Or _
                            mpage(19)

    End If

End Function

'detect read speeds
'from CDROM TOOL (GPL)
Public Function CDRomDetectSpeeds(ByVal strDrvID As String, _
                                  ByVal tolerance As Integer) As Integer()

    On Error Resume Next

    Dim intRet()    As Integer, RetCnt      As Integer
    Dim xSpeed      As Integer, kbSpeed     As Integer
    Dim maxSpeed    As Integer, i           As Integer
    Dim j           As Integer

    maxSpeed = CDRomGetSpeed(strDrvID).MaxRSpeed

    For xSpeed = maxSpeed \ 176 To 1 Step -1

        For i = tolerance To -tolerance Step -1

            CDRomSetCDSpeed strDrvID, xSpeed * 177 + i, 0, False

            If (xSpeed * 177 + i) \ 176 = CDRomGetSpeed(strDrvID).CurrRSpeed \ 176 Then

                'speed already in the list?
                If RetCnt > 0 Then
                    For j = 0 To RetCnt - 1
                        If intRet(j) \ 176 = xSpeed Then
                            GoTo ExitLoops
                        End If
                    Next
                End If

                'save new speed
                ReDim Preserve intRet(RetCnt) As Integer
                intRet(RetCnt) = xSpeed * 177 + i
                RetCnt = RetCnt + 1
                Exit For

            End If

        Next

ExitLoops:
    Next

    CDRomDetectSpeeds = intRet
End Function

'Optimal Power Calibration
Public Function CDRomSendOPCInformation(ByVal DrvID As String, _
                                        ByVal OPC As Boolean, _
                                        ByVal PtrBuffer As Long, _
                                        ByVal BufferLen As Long) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H54                   ' SEND OPC INFORMATION OpCode
    cmd(1) = Abs(OPC)               ' OPC?
    cmd(8) = BufferLen              ' allocation length

    CDRomSendOPCInformation = cd.ExecCMD(DrvID, cmd, 10, True, SRB_DIR_OUT, _
                                         PtrBuffer, BufferLen)
End Function

'read disc information
Public Function CDRomReadDiscInformation(ByVal strDrv As String, _
                                         ByVal PtrBuffer As Long, _
                                         ByVal BufferLen As Long) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H51                   ' READ DISC INFORMATION Op-Code
    cmd(7) = BufferLen \ &HFF       ' allocation length
    cmd(8) = BufferLen Mod &HFF     ' allocation length

    CDRomReadDiscInformation = cd.ExecCMD(strDrv, cmd, 10, False, _
                                          SRB_DIR_IN, PtrBuffer, BufferLen, 10)
End Function

'trys to get the type of the inserted CD/DVD
Public Function CDRomGetCDType(ByVal strDrv As String) As e_CDType

    Dim BufRDI          As t_RDI        ' buffer for ReadDiscInformation
    Dim BufAtip         As t_ATIP       ' ATIP for CD-R/CD-RW
    Dim conf_hdr(512)   As Byte         ' configuration header (active profile)

    'read disc information
    CDRomReadDiscInformation strDrv, VarPtr(BufRDI), Len(BufRDI) - 1

    'first try to read the ATIP to exclude CD-R/W
    If Not CDRomReadTOC(strDrv, 4, True, 0, VarPtr(BufAtip), Len(BufAtip) - 1) Then

        'if the Lead-Out is 255:255.255 MSF, it should be a CD-ROM
        If cd.MSF2LBA(BufRDI.LastPossibleLeadOutStart(1), _
                      BufRDI.LastPossibleLeadOutStart(2), _
                      BufRDI.LastPossibleLeadOutStart(3)) _
            = cd.MSF2LBA(255, 255, 255) Then

            'normal CD-ROM
            CDRomGetCDType = ROMTYPE_CDROM

        Else

            'could be a CD-ROM/R/RW
            CDRomGetCDType = ROMTYPE_CDROM_R_RW

        End If

    Else

        'ATIP could be read, either CD-R oder CD-RW.
        'but we could get fooled, so check the ATIP data :)

        'valid Lead-In start time?
        If BufAtip.LeadIn_Min > 0 Or _
           BufAtip.LeadIn_Sec > 0 Or _
           BufAtip.LeadIn_Frm > 0 Then

            'valide Lead-Out start time?
            If BufAtip.LeadIn_Min < 255 Or _
               BufAtip.LeadIn_Sec < 255 Or _
               BufAtip.LeadIn_Frm < 255 Then

                If IsBitSet(BufAtip.DiscType, 6) Then
                    'CD-RW
                    CDRomGetCDType = ROMTYPE_CDRW
                Else
                    'CD-R
                    CDRomGetCDType = ROMTYPE_CDR
                End If

            End If
        End If

    End If

    'is DVD in drive?
    If IsDVD(strDrv) Then

        'seems to be a DVD, determine its type
        CDRomGetCDType = GetDVDBookType(strDrv)

        'didn't work?
        If CDRomGetCDType = -1 Then

            'read the configuration header, we want the active profile
            If CDRomGetConfiguration(strDrv, 0, 2, VarPtr(conf_hdr(0)), _
                                     UBound(conf_hdr)) Then
    
                'CD type by the active drive profile
                Select Case (cd.LShift(conf_hdr(6), 8) Or conf_hdr(7))
                    Case &H8: CDRomGetCDType = ROMTYPE_CDROM       ' CD-ROM
                    Case &H9: CDRomGetCDType = ROMTYPE_CDR         ' CD-R
                    Case &HA: CDRomGetCDType = ROMTYPE_CDRW        ' CD-RW
                    Case &H10: CDRomGetCDType = ROMTYPE_DVD_ROM    ' DVD-ROM
                    Case &H11: CDRomGetCDType = ROMTYPE_DVD_R      ' DVD-R
                    Case &H12: CDRomGetCDType = ROMTYPE_DVD_RAM    ' DVD-RAM
                    Case &H13: CDRomGetCDType = ROMTYPE_DVD_RW     ' DVD-RW
                    Case &H14: CDRomGetCDType = ROMTYPE_DVD_RW     ' DVD-RW
                    Case &H1A: CDRomGetCDType = ROMTYPE_DVD_P_RW   ' DVD+RW
                    Case &H1B: CDRomGetCDType = ROMTYPE_DVD_P_R    ' DVD+R
                    Case Else: CDRomGetCDType = ROMTYPE_CDROM_R_RW ' das ging also mal garnicht...
                End Select
    
            End If

        End If

    End If
End Function

Public Function CDRomReadSubChannel( _
                ByVal DrvID As String, _
                ByVal MSF As Boolean, _
                ByVal SUBQ As Boolean, _
                ParamList As Byte, _
                TrackNum As Byte, _
                PtrBuffer As Long, _
                BufferLen As Long) As Boolean

    Dim cmd(9) As Byte

    cmd(0) = &H42                       ' READ SUB-CHANNEL [42h]
    cmd(1) = IIf(MSF, 3, 0)             ' MSF or LBA?
    cmd(2) = IIf(SUBQ, cd.LShift(1, 6), 0) ' Q Sub-Channel?
    cmd(3) = ParamList                  ' Parameter list
    cmd(6) = TrackNum                   ' specific track number
    cmd(7) = BufferLen \ &HFF           ' allocation length
    cmd(8) = BufferLen Mod &HFF         ' allocation length

    'execute
    CDRomReadSubChannel = cd.ExecCMD( _
                        DrvID, cmd, _
                        10, False, _
                        SRB_DIR_IN, _
                        PtrBuffer, _
                        BufferLen)

End Function

'search for CD-R(W) vendor
Private Function CDRomGetCDRWVendor(ByVal strDrv As String) As String
    Dim MSF_CD As String, MSF_V1 As String, MSF_V2 As String
    Dim cdrv() As String
    Dim atip As t_ATIP
    Dim i As Integer

    ReDim cdrv(69) As String

    'vendor list (from CDR-DAO [GPL])
     cdrv(0) = "97,28,30||97,46,50||Auvistar Industry Co.,Ltd."
     cdrv(1) = "97,26,60||97,46,60||CMC Magnetics Corporation"
     cdrv(2) = "97,23,10||00,00,00||Doremi Media Co., Ltd."
     cdrv(3) = "97,26,00||97,45,00||FORNET International PTE Ltd."
     cdrv(4) = "97,46,40||97,46,40||FUJI Photo Film Co., Ltd."
     cdrv(5) = "97,26,40||00,00,00||FUJI Photo Film Co., Ltd."
     cdrv(6) = "97,28,10||97,49,10||Gigastore Corporation"
     cdrv(7) = "97,25,20||97,47,10||Hitachi Maxwell, Ltd."
     cdrv(8) = "97,27,40||97,48,10||Kodak Japan Limited"
     cdrv(9) = "97,26,50||97,48,60||Lead Data Inc."
    cdrv(10) = "97,27,50||97,48,50||Mitsui Chemicals, Inc."
    cdrv(11) = "97,34,20||97,50,20||Mitsubishi Chemical Corporation"
    cdrv(12) = "97,28,20||97,46,20||Multi Media Masters & Machinary SA"
    cdrv(13) = "97,21,40||00,00,00||Optical Disc Manufacturing Equipment"
    cdrv(14) = "97,27,30||97,48,30||Pioneer Video Corporation"
    cdrv(15) = "97,27,10||97,48,20||Plasmon Data systems Ltd."
    cdrv(16) = "97,26,10||97,47,40||POSTECH Corporation"
    cdrv(17) = "97,27,20||97,47,20||Princo Corporation"
    cdrv(18) = "97,32,10||00,00,00||Prodisc Technology Inc."
    cdrv(19) = "97,27,60||97,48,00||Ricoh Company Limited"
    cdrv(20) = "97,31,00||97,47,50||Ritek Co."
    cdrv(21) = "97,26,20||00,00,00||SKC Co., Ltd."
    cdrv(22) = "97,24,10||00,00,00||SONY Corporation"
    cdrv(23) = "97,24,00||97,46,00||Taiyo Yuden Company Limited"
    cdrv(24) = "97,32,00||97,49,00||TDK Corporation"
    cdrv(25) = "97,25,60||97,45,60||Xcitek Inc."
    cdrv(26) = "97,22,60||97,45,20||Acer Media Technology, Inc"
    cdrv(27) = "97,25,50||00,00,00||AMS Technology Inc."
    cdrv(28) = "97,23,30||00,00,00||Audio Distributors Co., Ltd."
    cdrv(29) = "97,21,30||00,00,00||Bestdisc Technology Corporation"
    cdrv(30) = "97,30,10||97,50,30||CDA Datentraeger Albrechts GmbH"
    cdrv(31) = "97,22,40||97,45,40||CIS Technology Inc."
    cdrv(32) = "97,24,20||97,46,30||Computer Support Italy s.r.l."
    cdrv(33) = "97,23,60||00,00,00||Customer Pressing Oosterhout"
    cdrv(34) = "97,28,50||00,00,00||Delphi Technology Inc."
    cdrv(35) = "97,27,00||97,48,40||Digital Storage Technology Co., Ltd."
    cdrv(36) = "97,22,30||00,00,00||EXIMPO"
    cdrv(37) = "97,28,60||00,00,00||Friendly CD-Tek Co."
    cdrv(38) = "97,31,30||97,51,10||Grand Advance Technology Ltd."
    cdrv(39) = "97,29,50||00,00,00||General Magnetics Ld"
    cdrv(40) = "97,24,50||97,45,50||Guann Yinn Co.,Ltd."
    cdrv(41) = "97,29,00||00,00,00||Harmonic Hall Optical Disc Ltd."
    cdrv(42) = "97,29,30||97,51,50||Hile Optical Disc Technology Corp."
    cdrv(43) = "97,46,10||97,22,50||Hong Kong Digital Technology Co., Ltd."
    cdrv(44) = "97,25,30||97,51,20||INFODISC Technology Co., Ltd."
    cdrv(45) = "97,24,40||00,00,00||kdg mediatech AG"
    cdrv(46) = "97,28,40||97,49,20||King Pro Mediatek Inc."
    cdrv(47) = "97,23,00||97,49,60||Matsushita Electric Industrial Co., Ltd."
    cdrv(48) = "97,15,20||00,00,00||Mitsubishi Chemical Corporation"
    cdrv(49) = "97,25,00||00,00,00||Hi-SPACE, MPO"
    cdrv(50) = "97,23,20||00,00,00||Nacar Media sr"
    cdrv(51) = "97,26,30||00,00,00||Optical Disc Corporation"
    cdrv(52) = "97,28,00||97,49,30||Opti.Me.S. S.p.A."
    cdrv(53) = "97,23,50||00,00,00||OPTROM.INC."
    cdrv(54) = "97,47,60||00,00,00||Prodisc Technology Inc."
    cdrv(55) = "97,15,10||00,00,00||Ritek Co."
    cdrv(56) = "97,22,10||00,00,00||Seantram Technology Inc."
    cdrv(57) = "97,21,50||00,00,00||Sound Sound Multi-Media Development Limited"
    cdrv(58) = "97,29,00||00,00,00||Taeil Media Co.,Ltd."
    cdrv(59) = "97,18,60||00,00,00||Taroko International Co.,Ltd."
    cdrv(60) = "97,15,00||00,00,00||TDK Corporation."
    cdrv(61) = "97,29,20||00,00,00||Unidisc Technology Co.,Ltd."
    cdrv(62) = "97,24,30||97,45,10||Unitech Japan Inc."
    cdrv(63) = "97,29,10||97,50,10||Vanguard Disc Inc."
    cdrv(64) = "97,49,40||97,23,40||Victor Company of Japan, Ltd."
    cdrv(65) = "97,29,40||00,00,00||Viva Magnetics, Ltd."
    cdrv(66) = "97,25,40||00,00,00||Vivastar AG"
    cdrv(67) = "97,18,10||00,00,00||Wealth Fair Investment, Ltd."
    cdrv(68) = "97,22,00||00,00,00||Woongjin Media corp."
    cdrv(69) = "97,17,00||00,00,00||Moser Baer India Limited"

    'read ATIP
    CDRomReadTOC strDrv, 4, True, 0, VarPtr(atip), Len(atip) - 1

    'round up Lead-In
    atip.LeadIn_Frm = Val(Left(atip.LeadIn_Frm, 1) & "0")

    'convert to a useable format
    MSF_CD = atip.LeadIn_Min & "," & atip.LeadIn_Sec & "," & atip.LeadIn_Frm

    For i = 0 To UBound(cdrv)

        'the first possible Lead-In start time
        MSF_V1 = Left(cdrv(i), InStr(1, cdrv(i), "||") - 1)

        'the second possible Lead-In start time
        MSF_V2 = Mid(cdrv(i), InStr(1, cdrv(i), "||") + 2, _
                     InStrRev(cdrv(i), "||") - 3 - Len(MSF_V1))

        'if one of the times is right...
        If (MSF_CD = MSF_V1) Or (MSF_CD = MSF_V2) Then
            '... we have a winner
            CDRomGetCDRWVendor = Mid(cdrv(i), Len(MSF_V1) + Len(MSF_V2) + 5)
            Exit Function
        End If
    Next

    'nothing found
    CDRomGetCDRWVendor = "Unbekannt"
End Function

'simple DVD detection
Public Function IsDVD(ByVal strDrv As String) As Boolean
    Dim dummy(512) As Byte

    'DVD?
    If CDRomReadDVDStructure(strDrv, 0, 0, 0, VarPtr(dummy(0)), UBound(dummy)) Then
        If dummy(0) > 0 Or dummy(1) > 0 Then IsDVD = True
    End If
End Function

'read DVD structure
Public Function CDRomReadDVDStructure(ByVal strDrv As String, ByVal LBA As Long, _
                                      ByVal LayerNr As Byte, ByVal format As Byte, _
                                      ByVal PtrBuffer As Long, _
                                      ByVal BufferLen As Long) As Boolean

    Dim cmd(11) As Byte

    cmd(0) = &HAD                           ' READ DVD STRUCTURE Op-Code
    cmd(2) = cd.RShift(LBA, 24) And &HFF    ' LBA LSB
    cmd(3) = cd.RShift(LBA, 16) And &HFF
    cmd(4) = cd.RShift(LBA, 8) And &HFF
    cmd(5) = LBA And &HFF                   ' LBA MSB
    cmd(6) = LayerNr                        ' Layer Number
    cmd(7) = format                         ' Information Format
    cmd(8) = BufferLen \ &HFF               ' allocation length
    cmd(9) = BufferLen And &HFF             ' allocation length

    CDRomReadDVDStructure = cd.ExecCMD(strDrv, cmd, 12, False, _
                                       SRB_DIR_IN, PtrBuffer, BufferLen, 10)
End Function

'read DVD Book
Private Function GetDVDBookType(ByVal strDrv As String) As e_CDType
    Dim physdata As t_DVD_Phys
    Dim book As Byte

    'get DVD Book
    If CDRomReadDVDStructure(strDrv, 0, 0, 0, VarPtr(physdata), _
                          Len(physdata) - 1) Then

        '
        With physdata

            If IsBitSet(.BookType, 4) Then book = 1
            If IsBitSet(.BookType, 5) Then book = book Or 2
            If IsBitSet(.BookType, 6) Then book = book Or 4
            If IsBitSet(.BookType, 7) Then book = book Or 8

            Select Case book
                Case 0: GetDVDBookType = ROMTYPE_DVD_ROM   ' DVD-ROM
                Case 1: GetDVDBookType = ROMTYPE_DVD_RAM   ' DVD-RAM
                Case 2: GetDVDBookType = ROMTYPE_DVD_R     ' DVD-R
                Case 3: GetDVDBookType = ROMTYPE_DVD_RW    ' DVD-RW
                Case 9: GetDVDBookType = ROMTYPE_DVD_P_RW  ' DVD+RW
                Case 10: GetDVDBookType = ROMTYPE_DVD_P_R  ' DVD+R
                Case Else: GetDVDBookType = -1             ' unknown
            End Select

        End With

    End If
End Function

'bit in a byte is set?
Public Function IsBitSet(ByVal InByte As Byte, ByVal Bit As Byte) As Boolean
    IsBitSet = ((InByte And (2 ^ Bit)) > 0)
End Function

'format bytes
Public Function FormatFileSize(ByVal dblFileSize As Double, _
    Optional ByVal strFormatMask As String) As String

    Select Case dblFileSize
        Case 0 To 1023               ' Bytes
            FormatFileSize = format(dblFileSize) & " bytes"
        Case 1024 To 1048575         ' KB
            If strFormatMask = Empty Then strFormatMask = "###0"
            FormatFileSize = format(dblFileSize \ 1024, strFormatMask) & " KB"
        Case 1024# ^ 2 To 1073741823 ' MB
            If strFormatMask = Empty Then strFormatMask = "####0.00"
            FormatFileSize = format(dblFileSize \ (1024 ^ 2), strFormatMask) & " MB"
        Case Is > 1073741823#        ' GB
            If strFormatMask = Empty Then strFormatMask = "####0.00"
            FormatFileSize = format(dblFileSize \ (1024 ^ 3), strFormatMask) & " GB"
    End Select

End Function

'remove Chr$(0)s from a string
Private Function RemNulls(ByVal sVal As String) As String
    Dim sBuf As String
    Dim i As Integer

    For i = 1 To Len(sVal)
        If Asc(Mid(sVal, i, 1)) > 0 Then sBuf = sBuf & Mid(sVal, i, 1)
    Next

    RemNulls = sBuf
End Function

'add unsigned DWords
Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
   If Start And &H80000000 Then
      UnsignedAdd = Start + Incr
   ElseIf (Start Or &H80000000) < -Incr Then
      UnsignedAdd = Start + Incr
   Else
      UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
   End If
End Function

'pad a string to the left
Private Function lPAD(ByVal str As String, ByVal str_len As Integer, ByVal char As String) As String
    If Not (Len(str) = str_len Or Len(str) > str_len) Then
        lPAD = String(str_len - Len(str), char) & str
    Else
        lPAD = str
    End If
End Function

'Wave playtime
Public Function GetWAVLength(ByVal strFileName As String) As Long

    '   Waveheader
    Dim wavhdr  As t_WavHdr
    '   filehandle
    Dim FF      As Integer:  FF = FreeFile

    'read Waveheader
    Open strFileName For Binary As #FF

        Get #FF, , wavhdr

    Close #FF

    'length in seconds
    GetWAVLength = FileLen(strFileName) \ wavhdr.BytesPerSec

End Function

'MP3 playtime
'from VB Archiv
Public Function GetMP3Length(ByVal strFileName As String) As Long

    Dim strBuffer As String
    Dim lRet      As Long
    Dim sReturn   As String

    strBuffer = Space$(255)
    lRet = GetShortPathName(strFileName, strBuffer, Len(strBuffer))

    If lRet <> 0 Then _
        strFileName = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)

    ' open MP3 file
    mciSendString "open " & strFileName & _
                  " type MPEGVideo alias axCDDAWriterMP3", 0, 0, 0

    ' get the length of the file in ms
    sReturn = Space$(256)
    lRet = mciSendString("status axCDDAWriterMP3 length", _
                         sReturn, Len(sReturn), 0&)

    ' close the mp3 file
    mciSendString "close axCDDAWriterMP3", 0, 0, 0

    ' return the length in seconds
    GetMP3Length = Val(sReturn) / 1000

End Function
