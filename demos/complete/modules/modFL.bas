Attribute VB_Name = "modFL"
Option Explicit

' help structure for managing tracks to grab
Public Type t_AudioTrack
    Album       As String
    Artist      As String
    Title       As String
    no          As Integer
    grab        As Boolean
    startLBA    As Long
    endLBA      As Long
    lenLBA      As Long
End Type

Public Type t_AudioTracks
    Track(98)   As t_AudioTrack
    count       As Integer
End Type

' global Flamed Manager
Public cManager     As New FL_Manager
' global drive ID
Public strDrvID     As String
