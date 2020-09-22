VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_CDReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************
'*** Flamed CD Reader class
'***
'*** for reading sectors/MCN/ISRC
'*** from CD-R/ROM/RWs
'****************************************

Public Enum FL_BufferSize
    BS_TOO_BIG = -1
    BS_INVALID_TRACK = 0
End Enum

Public Enum FL_SectorSize
    SS_RAW_AUDIO_MODE0 = 2352
    SS_MODE1 = 2048
    SS_MODE2FORM1 = 2048
    SS_MODE2FORM2 = 2328
    SS_MODE2PLAIN = 2336
    SS_RAW_SUBCHANNELS = 2448
End Enum

Public Enum FL_SUBCHANNELS
    RCD_SUBCH_PW_RAW = &H1                            ' Raw Channels P-W
    RCD_SUBCH_Q = &H2                                 ' only Q channel
    RCD_SUBCH_PW_CORRECTED = &H4                      ' Channels P-W corrected
End Enum

' International Standard Recording Code
Public Function ReadISRC(DriveID As String, _
                Track As Byte) As String

    Dim udtISRC         As t_Q_ISRC
    Dim strISRC         As String

    If Track < 1 Or Track > 100 Then
        Exit Function
    End If

    If Not CDRomReadSubChannel(DriveID, True, _
        True, &H3, Track, VarPtr(udtISRC), Len(udtISRC) - 1) Then
            Exit Function
    End If

    ' valid data?
    If Not (udtISRC.ISRC(0) And &H80) Then
        Exit Function
    End If

    With udtISRC

        ' Country Code
        strISRC = strISRC & Chr$(.ISRC(1))
        strISRC = strISRC & Chr$(.ISRC(2))
        strISRC = strISRC & "-"

        ' Owner Code
        strISRC = strISRC & Chr$(.ISRC(3))
        strISRC = strISRC & Chr$(.ISRC(4))
        strISRC = strISRC & Chr$(.ISRC(5))
        strISRC = strISRC & "-"

        ' Year Of Recording
        strISRC = strISRC & Chr$(.ISRC(6))
        strISRC = strISRC & Chr$(.ISRC(7))
        strISRC = strISRC & "-"

        ' Serial Number
        strISRC = strISRC & Chr$(.ISRC(8))
        strISRC = strISRC & Chr$(.ISRC(9))
        strISRC = strISRC & Chr$(.ISRC(10))
        strISRC = strISRC & Chr$(.ISRC(11))
        strISRC = strISRC & Chr$(.ISRC(12))

    End With

    ReadISRC = strISRC

End Function

' UPC/Bar Code
Public Function ReadMCN(DriveID As String) As String

    Dim udtMCN      As t_Q_MCN
    Dim strMCN      As String
    Dim i           As Integer

    If Not CDRomReadSubChannel(DriveID, True, _
        True, &H2, 0, VarPtr(udtMCN), Len(udtMCN) - 1) Then
            Exit Function
    End If

    ' valid data?
    If Not (udtMCN.MCN(0) And &H80) Then
        Exit Function
    End If

    For i = 1 To 13
        strMCN = strMCN & Chr$(udtMCN.MCN(i))
    Next

    ReadMCN = strMCN

End Function

' read a bunch of sectors from a disk
Public Function ReadSectorsLBA(DriveID As String, _
                LBA As Long, _
                sectors As Integer, _
                Buffer() As Byte, _
                Optional raw As Boolean, _
                Optional subchannels As FL_SUBCHANNELS) As Boolean

    ReadSectorsLBA = CDRomReadCD(DriveID, _
                                 LBA, _
                                 sectors, _
                                 VarPtr(Buffer(0)), _
                                 UBound(Buffer), _
                                 IIf(raw, RCD_RAW, RCD_USRDATA), _
                                 subchannels)

End Function

' determine buffer size for a track
Public Function BufferSize(ByVal DriveID As String, _
                ByVal sectors As Integer, _
                ByVal Track As Integer, _
                Optional ByVal raw As Boolean, _
                Optional ByVal subchannels As FL_SUBCHANNELS) As Long

    Dim multiplier  As Long
    Dim udtInfo     As t_TrackInfo

    ' not more then 25 sectors allowed.
    ' e.g. for a read with raw+96 30 sectors
    ' the size would be 71 KBytes.
    ' As most miniports only allow
    ' a maximum of 64 KBytes, this would
    ' cause trouble.
    If sectors > 25 Then
        BufferSize = BS_TOO_BIG
        Exit Function
    End If

    If Not CDRomTrackInfo(DriveID, Track, udtInfo) Then
        BufferSize = BS_INVALID_TRACK
        Exit Function
    End If

    Select Case udtInfo.DataMode
        Case e_SectorModes.MODE_AUDIO: multiplier = SS_RAW_AUDIO_MODE0
        Case e_SectorModes.MODE_MODE1: multiplier = SS_MODE1
        Case e_SectorModes.MODE_MODE2: multiplier = SS_MODE2PLAIN
        Case e_SectorModes.MODE_MODE2_FORM1: multiplier = SS_MODE2FORM1
        Case e_SectorModes.MODE_MODE2_FORM2: multiplier = SS_MODE2FORM2
    End Select

    ' if raw then the mulitplier ist just 2352
    If raw Then multiplier = 2352

    ' add 96 bytes for corrected/interleaved
    ' or raw subchannels data
    If subchannels = RCD_SUBCH_PW_CORRECTED Or _
       subchannels = RCD_SUBCH_PW_RAW Then
        multiplier = multiplier + 96
    End If

    ' add 16 bytes for formatted Q Subchannel
    If subchannels = RCD_SUBCH_Q Then
        multiplier = multiplier + 16
    End If

    BufferSize = CLng(sectors) * multiplier

End Function
