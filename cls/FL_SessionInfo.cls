VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_SessionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************
'*** Flamed session info class
'***
'*** for information about sessions
'****************************************

Private intSessNum        As Integer

Private intFirstTrack     As Integer
Private intLastTrack      As Integer

Private lngFirstTrackLBA  As Long
Private lngLeadOutLBA     As Long

Private lngSize           As Long

Public Property Get SessionNumber() As Integer
    SessionNumber = intSessNum
End Property

Public Property Get FirstTrack() As Integer
    FirstTrack = intFirstTrack
End Property

Public Property Get LastTrack() As Integer
    LastTrack = intLastTrack
End Property

Public Property Get FirstTrackLBA() As Long
    FirstTrackLBA = lngFirstTrackLBA
End Property

Public Property Get LeadOutLBA() As Long
    LeadOutLBA = lngLeadOutLBA
End Property

Public Property Get Size() As Long
    Size = lngSize
End Property

Public Function GetInfo(ByVal DriveID As String, _
                        ByVal Session As Integer) As Boolean

    Dim toc As t_RTOC_STRUCT
    Dim i   As Long

    'read the RAW TOC
    If Not CDRomReadTOC(DriveID, 2, True, 1, VarPtr(toc), Len(toc)) Then Exit Function

    'parse the packets of the TOC
    For i = 0 To ((cd.LShift(toc.dummy(0), 8) Or toc.dummy(1)) \ 11) - 1

        Select Case toc.packet(i).point

            'get the LBA of the first track in the session
            Case intFirstTrack
                If toc.packet(i).sessionNr = Session Then _
                    lngFirstTrackLBA = cd.MSF2LBA(toc.packet(i).pmin, toc.packet(i).psec, toc.packet(i).pframe)

            'found the first track of the session
            Case &HA0
                If toc.packet(i).sessionNr = Session Then _
                    intFirstTrack = toc.packet(i).pmin

            'found the last track of the session
            Case &HA1
                If toc.packet(i).sessionNr = Session Then _
                    intLastTrack = toc.packet(i).pmin

            'found the Lead-Out of the session
            Case &HA2
                If toc.packet(i).sessionNr = Session Then _
                    lngLeadOutLBA = cd.MSF2LBA(toc.packet(i).pmin, toc.packet(i).psec, toc.packet(i).pframe)

        End Select

    Next

    intSessNum = Session
    lngSize = (LeadOutLBA - FirstTrackLBA) * 2352

    GetInfo = True

End Function
