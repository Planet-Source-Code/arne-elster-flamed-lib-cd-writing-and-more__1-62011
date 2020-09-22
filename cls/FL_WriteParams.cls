VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_WriteParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************
'*** Flamed write parameters mode page reader/writer
'***
'*** read and write write parameters
'*** for CD writing
'***************************************************

' '''''''''''''''''''''''''''''''
' ToDo: MCN
'       ISRC
'       Sub Headers
'       Vendor Specific
' '''''''''''''''''''''''''''''''


' Copied some stuff from MoxCDBurn

Public Enum FL_WriteType
    WT_PACKET = &H0
    WT_TAO = &H1
    WT_SAO = &H2
    WT_RAW = &H3
End Enum

Public Enum FL_DataBlockType
    DB_RAW_2352 = &H0
    DB_RAW_2368 = &H1
    DB_RAW_2448 = &H2
    DB_RAW_2448_PW = &H3
    DB_MODE1_ISO = &H8
    DB_MODE2_ISO = &H9
    DB_MODE2_F1_2048 = &HA
    DB_MODE2_F1_2056 = &HB
    DB_MODE2_F2_2324 = &HC
    DB_MODE2_F2_2332 = &HD
End Enum

Public Enum FL_SessionFormat
    SF_CDDA_DATA = &H0
    SF_CDI = &H10
    SF_CDROM_XA = &H20
End Enum

Private bytWriteType        As Byte
Private blnTestWrite        As Boolean
Private blnLinkSizeValid    As Boolean
Private blnBURNProof        As Boolean
Private bytTrackMode        As Byte
Private blnCopy             As Boolean
Private blnFixedPacket      As Boolean
Private bytMultiSession     As Byte
Private bytDataBlockType    As Byte
Private bytLinkSize         As Byte
Private bytHostAppCode      As Byte
Private bytSessFormat       As Byte
Private lngPacketSize       As Long
Private intAudioPauseLength As Integer
Private bytMCN(14)          As Byte
Private bytISRC(14)         As Byte
Private bytSubHeaderData(3) As Byte
Private bytVendorSpec(3)    As Byte

Private Const MODE_HEADER_SIZE = 8&
Private Const MMC4_MODE_PAGE_WRPAR_SIZE = 56&

Private lngPageSize         As Long
Private lngHeaderSize       As Long
Private lngDataSize         As Long

Private pagebuf(511)        As Byte

Public Property Let WriteType(aval As FL_WriteType)
    bytWriteType = aval
End Property

Public Property Get WriteType() As FL_WriteType
    WriteType = bytWriteType
End Property

Public Property Let TestMode(aval As Boolean)
    blnTestWrite = aval
End Property

Public Property Get TestMode() As Boolean
    TestMode = blnTestWrite
End Property

Public Property Let LinkSizeValid(aval As Boolean)
    blnLinkSizeValid = aval
End Property

Public Property Get LinkSizeValid() As Boolean
    LinkSizeValid = blnLinkSizeValid
End Property

Public Property Let BURNProof(aval As Boolean)
    blnBURNProof = aval
End Property

Public Property Get BURNProof() As Boolean
    BURNProof = blnBURNProof
End Property

Public Property Let TrackMode(aval As Byte)
    bytTrackMode = aval
End Property

Public Property Get TrackMode() As Byte
    TrackMode = bytTrackMode
End Property

Public Property Let Copy(aval As Boolean)
    blnCopy = aval
End Property

Public Property Get Copy() As Boolean
    Copy = blnCopy
End Property

Public Property Let FixedPacket(aval As Boolean)
    blnFixedPacket = aval
End Property

Public Property Get FixedPacket() As Boolean
    FixedPacket = blnFixedPacket
End Property

Public Property Let Multisession(aval As Byte)
    bytMultiSession = aval
End Property

Public Property Get Multisession() As Byte
    Multisession = bytMultiSession
End Property

Public Property Let DataBlockType(aval As FL_DataBlockType)
    bytDataBlockType = aval
End Property

Public Property Get DataBlockType() As FL_DataBlockType
    DataBlockType = bytDataBlockType
End Property

Public Property Let LinkSize(aval As Byte)
    bytLinkSize = aval
End Property

Public Property Get LinkSize() As Byte
    LinkSize = bytLinkSize
End Property

Public Property Let ApplicationCode(aval As Byte)
    bytHostAppCode = aval
End Property

Public Property Get ApplicationCode() As Byte
    ApplicationCode = bytHostAppCode
End Property

Public Property Let SessionFormat(aval As FL_SessionFormat)
    bytSessFormat = aval
End Property

Public Property Get SessionFormat() As FL_SessionFormat
    SessionFormat = bytSessFormat
End Property

Public Property Let PacketSize(aval As Long)
    lngPacketSize = aval
End Property

Public Property Get PacketSize() As Long
    PacketSize = lngPacketSize
End Property

Public Property Let AudioPauseLength(aval As Integer)
    intAudioPauseLength = aval
End Property

Public Property Get AudioPauseLength() As Integer
    AudioPauseLength = intAudioPauseLength
End Property

Public Property Get DataSize() As Long
    DataSize = lngDataSize
End Property

Public Property Get HeaderSize() As Long
    HeaderSize = lngHeaderSize
End Property

Public Property Get PageSize() As Long
    PageSize = lngPageSize
End Property

' Read the WPP
Public Function GetData(DrvID As String) As Boolean

    Dim i           As Long
    Dim blnFound    As Boolean

    ' mode page to buffer
    If Not CDRomModeSense10(DrvID, &H5, VarPtr(pagebuf(0)), UBound(pagebuf) + 1) Then
        Exit Function
    End If

    ' get the size of the mode page
    lngPageSize = (cd.LShift(pagebuf(0), 8) Or _
                    pagebuf(1)) + 2

    ' maybe there are block descriptors
    ' (even if MMC-3 says: Block Descriptor Length = 0)
    ' Well, I don't know, so ...
    ' get the size of the mode page header
    lngHeaderSize = lngPageSize - MMC4_MODE_PAGE_WRPAR_SIZE

    ' look for the start of the mode page data
    If Not (pagebuf(lngHeaderSize + 1) And &H3F) = &H5 Then

        ' look in the whole buffer for &H5
        ' and take the first found pos
        For i = 3 To UBound(pagebuf) + 1

            ' found it?
            If (pagebuf(i) And &H3F) = &H5 Then
                If pagebuf(i + 1) = &H32 Or pagebuf(i + 1) = &H36 Then
                    ' yep, exit function
                    lngHeaderSize = i - 1
                    lngDataSize = lngPageSize - lngHeaderSize + 1
                    GetData = True
                    Exit Function
                End If
            End If

        Next

    Else

        lngDataSize = lngPageSize - lngHeaderSize
        GetData = True

    End If

End Function

Public Function SendData(DrvID As String) As Boolean

    Dim start   As Long

    start = lngHeaderSize + 2

    ' Byte 2
    pagebuf(start + 1) = cd.LShift(Abs(blnBURNProof), 6)
    pagebuf(start + 1) = pagebuf(start + 1) Or cd.LShift(Abs(blnLinkSizeValid), 5)
    pagebuf(start + 1) = pagebuf(start + 1) Or cd.LShift(Abs(blnTestWrite), 4)
    pagebuf(start + 1) = pagebuf(start + 1) Or (bytWriteType And &HF)

    ' Byte 3
    pagebuf(start + 2) = cd.LShift(bytMultiSession, 6)
    pagebuf(start + 2) = pagebuf(start + 2) Or cd.LShift(Abs(blnFixedPacket), 5)
    pagebuf(start + 2) = pagebuf(start + 2) Or cd.LShift(Abs(blnCopy), 4)
    pagebuf(start + 2) = pagebuf(start + 2) Or (bytTrackMode And &HF)

    ' Byte 4
    pagebuf(start + 3) = bytDataBlockType And &HF

    ' Byte 5
    pagebuf(start + 4) = bytLinkSize And &HF

    ' Byte 6
    'pagebuf(start + 5) = 0

    ' Byte 7
    pagebuf(start + 6) = bytHostAppCode

    ' Byte 8
    pagebuf(start + 7) = bytSessFormat

    ' Byte 9
    'pagebuf(start + 8) = 0

    ' Byte 10, 11, 12, 13
    CopyMemory pagebuf(start + 9), lngPacketSize, 4

    ' Byte 14, 15
    pagebuf(start + 13) = cd.RShift(intAudioPauseLength, 8) And &HFF
    pagebuf(start + 14) = intAudioPauseLength And &HFF

    ' Byte 16 - 31
    CopyMemory pagebuf(start + 15), bytMCN(0), 15

    ' Byte 32 - 47
    CopyMemory pagebuf(start + 31), bytISRC(0), 17

    ' Byte 48 - 51
    CopyMemory pagebuf(start + 47), bytSubHeaderData(0), 4

    ' Byte 52 - 55
    CopyMemory pagebuf(start + 51), bytVendorSpec(0), 4

    ' send the page
    SendData = CDRomModeSelect10(DrvID, VarPtr(pagebuf(0)), lngPageSize)

End Function

' Extract parameters from page
Public Sub ExtractData()

    Dim start   As Long

    start = lngHeaderSize + 2

    ' Byte 2
    blnBURNProof = IsBitSet(pagebuf(start + 1), 6)
    blnLinkSizeValid = IsBitSet(pagebuf(start + 1), 5)
    blnTestWrite = IsBitSet(pagebuf(start + 1), 4)
    bytWriteType = pagebuf(start + 1) And &HF

    ' Byte 3
    bytMultiSession = pagebuf(start + 2) And &HC0
    blnFixedPacket = IsBitSet(pagebuf(start + 2), 5)
    blnCopy = IsBitSet(pagebuf(start + 2), 4)
    bytTrackMode = pagebuf(start + 2) And &HF

    ' Byte 4
    bytDataBlockType = pagebuf(start + 3) And &HF

    ' Byte 5
    bytLinkSize = pagebuf(start + 4)

    ' Byte 6
    ' reserved

    ' Byte 7
    bytHostAppCode = pagebuf(start + 6) And &H3F

    ' Byte 8
    bytSessFormat = pagebuf(start + 7)

    ' Byte 9
    ' reserved

    ' Byte 10, 11, 12, 13
    CopyMemory lngPacketSize, pagebuf(start + 9), 4

    ' Byte 14, 15
    CopyMemory intAudioPauseLength, pagebuf(start + 13), 2

    intAudioPauseLength = cd.LShift(pagebuf(start + 13), 8) Or _
                                    pagebuf(start + 14)

    ' Byte 16 - 31
    CopyMemory bytMCN(0), pagebuf(start + 15), 15

    ' Byte 32 - 47
    CopyMemory bytISRC(0), pagebuf(start + 31), 15

    ' Byte 48, 49, 50, 51
    CopyMemory bytSubHeaderData(0), pagebuf(start + 47), 4

    ' Byte 52, 53, 54, 55
    CopyMemory bytVendorSpec(0), pagebuf(start + 51), 4

End Sub
