VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_FreeDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************
'*** Flamed FreeDB query class
'***
'*** query FreeDB database
'****************************************

Public Enum FL_CDDBState
    CDDB_HELLO
    CDDB_QUERY
    CDDB_RESULT
    CDDB_DATA
    CDDB_CLOSE
End Enum

Public Event Status(Status As FL_CDDBState)

Private WithEvents sck  As CSocket
Attribute sck.VB_VarHelpID = -1

Private WithEvents tmr  As clsTimer
Attribute tmr.VB_VarHelpID = -1

Private cCDInfo         As New FL_CDInfo

' handshake
Private Const MSG_OK_IO             As String = "200"
Private Const MSG_OK_READ           As String = "201"
Private Const MSG_PERM_DENIED       As String = "432"
Private Const MSG_TOO_MANY_USERS    As String = "433"
Private Const MSG_SERVER_LOAD       As String = "434"

Private Const MSG_HANDSHAKE_SUCCESS As String = "200"
Private Const MSG_HANDSHAKE_FAILED  As String = "431"
Private Const MSG_HANDSHAKE_ALREADY As String = "402"

' query database
Private Const MSG_EXACT_MATCH       As String = "200"
Private Const MSG_INEXACT_MATCH     As String = "211"
Private Const MSG_NO_MATCH          As String = "202"
Private Const MSG_DB_CORRUPT        As String = "403"
Private Const MSG_NO_HANDSHAKE      As String = "409"

Private Const MSG_ENTRY_FOLLOWS     As String = "210"
Private Const MSG_ENTRY_NOT_FOUND   As String = "401"
Private Const MSG_SERVER_ERROR      As String = "402"

Private Const CDDB_SERVER   As String = "freedb.org"
Private Const CDDB_PORT     As Integer = 8880
Private Const CDDB_USER     As String = "clsCDDB"
Private Const CDDB_HOST     As String = "local.host.com"
Private Const CDDB_CLIENT   As String = "clsCDDB"
Private Const CDDB_VERSION  As String = "v0.5"

Private blnFinished         As Boolean

Private State               As FL_CDDBState

Private strResult           As String

Private strArtist           As String
Private strAlbum            As String
Private strTracks()         As String

Private blnCantClose        As Boolean
Private blnCancel           As Boolean

Private lngTimeout          As Long

Private m_DrvID             As String

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" _
Alias "InternetGetConnectedStateExA" ( _
    ByRef lpdwFlags As Long, _
    ByVal lpszConnectionName As String, _
    ByVal dwNameLen As Long, _
    ByVal dwReserved As Long _
) As Long

Private Const INTERNET_CONNECTION_MODEM = &H1&
Private Const INTERNET_CONNECTION_LAN = &H2&
Private Const INTERNET_CONNECTION_PROXY = &H4&
Private Const INTERNET_RAS_INSTALLED = &H10&
Private Const INTERNET_CONNECTION_OFFLINE = &H20&
Private Const INTERNET_CONNECTION_CONFIGURED = &H40&

' VB-City
Public Function IsConnectedToInternet() As Boolean

    Dim sNameBuf    As String, msg  As String
    Dim dwFlags     As Long, lPos   As Long

    sNameBuf = String$(513, 0)

    If InternetGetConnectedStateEx(dwFlags, sNameBuf, 512&, 0&) Then
        IsConnectedToInternet = True
    End If

End Function

Public Function CDDBID() As String
    CDDBID = CDRomCDDB_ID(m_DrvID)
End Function

Public Property Get DriveID() As String
    DriveID = m_DrvID
End Property

Public Property Let DriveID(aval As String)
    m_DrvID = aval
End Property

Public Property Get Timeout() As Long
    Timeout = lngTimeout / 1000
End Property

Public Property Let Timeout(Seconds As Long)
    lngTimeout = Seconds * 1000
End Property

Public Function Query(Cancel As Boolean) As Boolean

    On Error Resume Next

    blnFinished = False
    strResult = vbNullString

    If CDRomGetTOC(DriveID).Tracks = 0 Then Exit Function

    ' set timeout
    tmr.Interval = lngTimeout
    tmr.Enabled = True

    ' byref cancel param
    blnCancel = Cancel

    ' connect to FreeDB
    sck.Connect CDDB_SERVER, CDDB_PORT

    Do While Not blnFinished

        If Cancel Then

            blnCancel = Cancel

            Do While blnCantClose
                DoEvents
            Loop

            sck.CloseSocket
            Cancel = False

            Exit Function

        End If

        If sck.State = sckClosed Or _
           sck.State = sckError Then
            blnFinished = True
        End If

        DoEvents
    Loop

    ' really recieved a track listing?
    If Right$(strResult, 3) = "." & vbCrLf Then
        ParseResult
        Query = True
    End If

End Function

Private Sub Class_Initialize()

    Set sck = New CSocket
    Set tmr = New clsTimer

    lngTimeout = 8& * 1000&

End Sub

Private Sub sck_OnClose()

    RaiseEvent Status(State)

    blnFinished = True
    tmr.Enabled = False

End Sub

Private Sub sck_OnConnect()
    State = CDDB_HELLO
End Sub

Private Sub sck_OnDataArrival(ByVal bytesTotal As Long)

    Dim strBuffer   As String

    If blnCancel Then Exit Sub

    blnCantClose = True

    sck.GetData strBuffer

    ' data arrived, reset timeout
    tmr.Enabled = False
    tmr.Interval = lngTimeout
    tmr.Enabled = True

    Select Case State

        ' handshake
        Case CDDB_HELLO

            RaiseEvent Status(State)

            If IsHandshake(strBuffer) Then
                sck.SendData HelloString()
            Else
                sck.CloseSocket
                State = CDDB_CLOSE
                Exit Sub
            End If

            State = CDDB_QUERY

        ' query
        Case CDDB_QUERY

            RaiseEvent Status(State)

            If IsHandshake(strBuffer) Then
                sck.SendData QueryString()
            Else
                sck.CloseSocket
                State = CDDB_CLOSE
                Exit Sub
            End If

            State = CDDB_RESULT

        ' result
        Case CDDB_RESULT

            RaiseEvent Status(State)

            Select Case Resp(strBuffer)

                Case MSG_EXACT_MATCH
                    ' wait for data
                    sck.SendData QueryResultString(strBuffer)
                    State = CDDB_DATA

                Case MSG_INEXACT_MATCH
                    ' take the first best
                    sck.SendData QueryInexactResultString(strBuffer)
                    State = CDDB_DATA

                Case MSG_NO_MATCH
                    ' quit
                    sck.CloseSocket
                    State = CDDB_CLOSE

                Case MSG_DB_CORRUPT
                    ' quit
                    sck.CloseSocket
                    State = CDDB_CLOSE

                Case MSG_NO_HANDSHAKE
                    ' quit
                    sck.CloseSocket
                    State = CDDB_CLOSE

                Case Else
                    ' quit
                    sck.CloseSocket
                    State = CDDB_CLOSE

            End Select

        Case CDDB_DATA

            RaiseEvent Status(State)

            ' save queried data
            strResult = strResult & strBuffer
            If Right$(strBuffer, 3) = "." & vbCrLf Then
                ' quit
                sck.CloseSocket
                State = CDDB_CLOSE
            End If

        Case CDDB_CLOSE
            ' close the connection

    End Select

    blnCantClose = False

End Sub

Private Function IsHandshake(str As String) As Boolean

    Select Case Resp(str)
        Case MSG_OK_IO: IsHandshake = True
        Case MSG_OK_READ: IsHandshake = True
        Case Else:
    End Select

End Function

Private Function Resp(str As String) As String
    Resp = Left$(str, 3)
End Function

Private Function HelloString() As String

    HelloString = "cddb hello " & _
                  CDDB_USER & " " & _
                  CDDB_HOST & " " & _
                  CDDB_CLIENT & " " & _
                  CDDB_VERSION & vbCrLf

End Function

Private Function QueryString() As String

    Dim strOffsets  As String
    Dim i           As Long
    Dim lngSeconds  As Long
    Dim toc         As t_TOC_STRUCT
    Dim nfo         As t_TrackInfo

    ' read the disc's TOC
    CDRomReadTOC m_DrvID, 0, True, 0, VarPtr(toc), Len(toc) - 1

    With toc

        For i = 0 To .LastTrack - 1

            With .TocTrack(i)

                ' add start offsets
                strOffsets = strOffsets & " " & _
                             cd.MSF2LBA(.addr(1), .addr(2), .addr(3))

            End With

        Next

        With .TocTrack(i)

            ' save the play length
            lngSeconds = cd.MSF2LBA(.addr(1), .addr(2), .addr(3)) / 75

        End With

    End With

    QueryString = "cddb query " & _
                  CDRomCDDB_ID(m_DrvID) & " " & _
                  CDRomGetTOC(m_DrvID).Tracks & _
                  strOffsets & " " & _
                  lngSeconds & vbCrLf

End Function

Private Function QueryResultString(str As String) As String

    Dim strSplt()   As String

    Dim strCode     As String
    Dim strCat      As String
    Dim strDiscID   As String
    Dim strTitle    As String

    strSplt = Split(str, " ")

    strCode = strSplt(0)
    strCat = strSplt(1)
    strDiscID = strSplt(2)
    strTitle = strSplt(3)

    QueryResultString = "cddb read " & _
                        strCat & " " & _
                        strDiscID & vbCrLf

End Function

Private Function QueryInexactResultString(str As String) As String

    Dim strLines()  As String
    Dim strSplit()  As String

    Dim strCat      As String
    Dim strDiscID   As String

    ' get the first result for the inexact query
    ' and query it :)

    strLines = Split(str, vbCrLf)
    strSplit = Split(strLines(1), " ")

    strCat = strSplit(0)
    strDiscID = strSplit(1)

    QueryInexactResultString = "cddb read " & _
                               strCat & " " & _
                               strDiscID & vbCrLf

End Function

Private Sub ParseResult()

    Dim strLines()  As String
    Dim Title       As String
    Dim name        As String
    Dim Value       As String
    Dim Track       As Integer
    Dim i           As Long

    ReDim strTracks(CDRomGetTOC(m_DrvID).Tracks) As String

    strLines = Split(strResult, vbCrLf)

    ' go through every line
    For i = 0 To UBound(strLines)

        If IsNumeric(Left$(strLines(i), 3)) Then GoTo NextItem
        If Left$(strLines(i), 1) = "#" Then GoTo NextItem
        If Left$(strLines(i), 1) = "." Then GoTo NextItem
        If Trim$(strLines(i)) = "" Then GoTo NextItem

        ' get name and value for the current field
        name = UCase$(Trim$(Left$(strLines(i), InStr(strLines(i), "=") - 1)))
        Value = Trim$(Mid$(strLines(i), Len(name) + 2))

        Select Case True

            ' save the artist/album
            Case name = "DTITLE"
                Title = Title & Value

            ' save a track title
            Case name Like "TTITLE*"
                Track = CInt(Mid(name, 7))
                strTracks(Track) = strTracks(Track) & Value

        End Select

NextItem:

    Next

    strArtist = Left$(Title, InStr(Title, "/") - 1)
    strAlbum = Trim$(Mid$(Title, Len(strArtist) + 2))
    strArtist = Trim$(strArtist)

End Sub

Property Get Artist() As String
    Artist = strArtist
End Property

Property Get Album() As String
    Album = strAlbum
End Property

Property Get Tracks() As Integer
    Tracks = UBound(strTracks)
End Property

Property Get Track(trk As Integer) As String
    Track = strTracks(trk - 1)
End Property

Private Sub tmr_Timer()
    ' timeout
    sck.CloseSocket
End Sub
