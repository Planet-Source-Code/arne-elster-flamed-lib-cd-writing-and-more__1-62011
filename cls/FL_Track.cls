VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_TrackInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************
'*** Flamed track info class
'***
'*** for information about tracks
'****************************************

Public Enum FL_TrackModes
    MODE_AUDIO                  ' Audio (or Mode-0)
    MODE_MODE1                  ' Mode-1
    MODE_MODE2                  ' Plain Mode-2
    MODE_MODE2_FORM1            ' Mode-2 Form-1
    MODE_MODE2_FORM2            ' Mode-2 Form-2
End Enum

Private blnLastTrackInSession    As Boolean

Private intSession  As Integer
Private intTrack    As Integer

Private udeMode     As e_SectorModes

Private cStart      As New FL_MSF
Private cEnd        As New FL_MSF
Private cLength     As New FL_MSF

Public Property Get TrackMode() As FL_TrackModes
    TrackMode = udeMode
End Property

Public Property Get LastTrackInSession() As Boolean
    LastTrackInSession = blnLastTrackInSession
End Property

Public Property Get Session() As Integer
    Session = intSession
End Property

Public Property Get Track() As Integer
    Track = intTrack
End Property

Public Property Get mode() As FL_TrackModes
    mode = udeMode
End Property

Public Property Get TrackStart() As FL_MSF
    Set TrackStart = cStart
End Property

Public Property Get TrackEnd() As FL_MSF
    Set TrackEnd = cEnd
End Property

Public Property Get TrackLength() As FL_MSF
    Set TrackLength = cLength
End Property

Public Function GetInfo(DriveID As String, _
                Track As Integer) As Boolean

    Dim udtInfo As t_TrackInfo

    If Not CDRomTestUnitReady(DriveID) Then _
        Exit Function

   If Track < 1 Or Track > 100 Then _
        Exit Function

    GetInfo = CDRomTrackInfo(DriveID, _
                             Track, _
                             udtInfo)

    With udtInfo
        udeMode = .DataMode
        cEnd.LBA = .endLBA
        cStart.LBA = .startLBA
        cLength.LBA = .Length
        blnLastTrackInSession = .LastTrackInSession
        intSession = .Session
        intTrack = .Track
    End With

End Function
