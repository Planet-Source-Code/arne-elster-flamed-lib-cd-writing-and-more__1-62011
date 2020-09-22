VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_CDText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************
'*** Flamed CD-Text reader
'***
'*** reads CD-Text from CDs
'****************************************

Private strAlbum    As String
Private strArtist   As String
Private strTrack()  As String

Public Property Get Album() As String
    Album = strAlbum
End Property

Public Property Get Artist() As String
    Artist = strArtist
End Property

Public Property Get TrackCount() As Integer
    TrackCount = UBound(strTrack) + 1
End Property

Public Property Get Track(index As Integer) As String
    Track = strTrack(index)
End Property

Public Function ReadCDText(DriveID As String) As Boolean

    Dim strs()  As String
    Dim I       As Integer
    Dim J       As Integer

    If Not CDRomHasCDText(DriveID) Then Exit Function
    strs() = CDRomReadCDText(DriveID)

    strAlbum = strs(0)
    strArtist = strs(1)

    For I = 2 To UBound(strs)
        If Not strs(I) = vbNullString Then
            ReDim Preserve strTrack(J)
            strTrack(J) = strs(I)
            J = J + 1
        End If
    Next

    ReadCDText = True

End Function
