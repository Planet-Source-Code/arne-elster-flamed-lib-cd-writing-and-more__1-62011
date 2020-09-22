VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_ImageConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum FL_BIN2ISO_RET
    BIN2ISO_UNKNOWN
    BIN2ISO_NOT_RAW
    BIN2ISO_INVALID_MODE
    BIN2ISO_CANCELED
    BIN2ISO_OK
End Enum

Public Event Progress(ByVal Percent As Integer, Cancel As Boolean)

' convert raw M1/M2F1 BIN to ISO
Public Function ConvertBINtoISO( _
            ByVal strIN As String, _
            ByVal strOUT As String) _
            As FL_BIN2ISO_RET

    On Error GoTo ExitFnc

    '   IN file handle          OUT file handle
    Dim ffIN        As Integer, ffOUT   As Integer
    '   string buffer           userdata
    Dim strBuffer   As String, btUser() As Byte
    '   buffer for a m1/m2f1 sector
    Dim btSector(2047)  As Byte

    '   bytes to skip before/after userdata
    Dim lngSkipBefore   As Long
    Dim lngSkipAfter    As Long

    '   cancel bool
    Dim blnCancel       As Boolean

    '   counters
    Dim i               As Long, j      As Long



    ' open files for read/write access
    ffIN = FreeFile
    Open strIN For Binary As #ffIN
    ffOUT = FreeFile
    Open strOUT For Binary As #ffOUT


    ' has SYNC pattern?
    strBuffer = Space(Len(SYNCPATTERN) / 2)
    Get #ffIN, 1, strBuffer

    For i = 1 To Len(strBuffer)
        strBuffer = strBuffer & format(Hex$(Asc(Mid$(strBuffer, i, 1))), "00")
    Next
    strBuffer = Mid$(strBuffer, 13)

    ' DATA
    If strBuffer = SYNCPATTERN Then

        strBuffer = Space(19)
        Get #ffIN, 1, strBuffer

        ' check for valid mode (1 or 2f1)
        Select Case Asc(Mid$(strBuffer, 16, 1))

            Case 1
                lngSkipBefore = 16      ' skip SYNC/header
                lngSkipAfter = 290      ' skip EDC/ECC

            Case 2
                If (Asc(Mid$(strBuffer, 19)) And &H20) = 0 Then
                    lngSkipBefore = 24  ' skip SYNC/header/subheader
                    lngSkipAfter = 280  ' skip EDC/ECC
                Else
                    ConvertBINtoISO = BIN2ISO_INVALID_MODE
                    GoTo ExitFnc
                End If

            Case Else
                ConvertBINtoISO = BIN2ISO_INVALID_MODE

        End Select
                
    Else
        ' AUDIO/MODE-0 or simply not raw
        ConvertBINtoISO = BIN2ISO_NOT_RAW
        GoTo ExitFnc
    End If

    ReDim btUser(2352& * 6& - 1&) As Byte

    ' seek back to start
    Seek #ffIN, 1
    Do While Not Seek(ffIN) >= LOF(ffIN)

        If Seek(ffIN) + (UBound(btUser) + 1) > LOF(ffIN) Then
            ReDim btUser(LOF(ffIN) - Seek(ffIN)) As Byte
        End If

        ' get a chunk and extract userdata
        Get #ffIN, , btUser

        i = 0
        For j = 0 To (UBound(btUser) + 1) \ 2352& - 1&
            i = j * 2352 + lngSkipBefore
            CopyMemory btSector(0), btUser(i), 2048
            Put #ffOUT, , btSector
        Next

        RaiseEvent Progress(Seek(ffIN) / LOF(ffIN) * 100, blnCancel)

        If blnCancel Then
            ConvertBINtoISO = BIN2ISO_CANCELED
            GoTo ExitFnc
        End If

        DoEvents

    Loop

    ConvertBINtoISO = BIN2ISO_OK

ExitFnc:
    ' close files
    Close #ffIN, #ffOUT

End Function
