VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_SubChannel16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************
'*** Flamed Q Sub-Channel decoder
'*** supports mode 1, 2 and 3
'****************************************

' NOTE: Looks like Sub-Channels don't
'       really have something like an
'       error correction.

Private blnCRCValid As Boolean

Private cCRC        As New clsCRC

Private relmsf      As New FL_MSF
Private absmsf      As New FL_MSF

Private udtQ        As t_Qch

Private udtQ1       As t_Qch_Mode1
Private udtQ2       As t_Qch_Mode2
Private udtQ3       As t_Qch_Mode3

Private udeMode     As FL_SubChannelQModes

Private strMCN(12)  As String

Private strISRCC(1) As String
Private strISRCO(2) As String
Private strISRCY(1) As String
Private strISRCS(4) As String

Public Property Get CRCValid() As Boolean
    CRCValid = blnCRCValid
End Property

Public Property Get QMode() As FL_SubChannelQModes
    QMode = udeMode
End Property

Public Property Get ISRC() As String
    ISRC = strISRCC(0) & strISRCC(1) & "-" & _
           strISRCO(0) & strISRCO(1) & strISRCO(2) & "-" & _
           strISRCY(0) & strISRCY(1) & "-" & _
           strISRCS(0) & strISRCS(1) & strISRCS(2) & strISRCS(3) & strISRCS(4)
End Property

Public Property Get MCN() As String()
    MCN = strMCN
End Property

Public Property Get AbsTime() As FL_MSF
    absmsf.F = udtQ1.afrm
    absmsf.s = udtQ1.asec
    absmsf.M = udtQ1.amin
    Set AbsTime = absmsf
End Property

Public Property Get RelTime() As FL_MSF
    relmsf.F = udtQ1.frm
    relmsf.s = udtQ1.sec
    relmsf.M = udtQ1.min
    Set RelTime = relmsf
End Property

Public Property Get TrackIndex() As Byte
    TrackIndex = udtQ1.index
End Property

Public Property Get TrackNo() As Byte
    TrackNo = udtQ1.TNO
End Property

Public Function GetCh(Data() As Byte, offset As Long) As Boolean

    CopyMemory udtQ, Data(offset), 12

    Select Case (udtQ.CTLADR And &HF)
        Case &H1:   ' Mode 1 Channel
            udeMode = QCH_Mode1
        Case &H2:   ' Mode 2 Channel
            udeMode = QCH_Mode2
        Case &H3:   ' Mode 3 Channel
            udeMode = QCH_Mode3
        Case Else:  ' unspoorted
            udeMode = QCH_Others
            Exit Function
    End Select

    Select Case udeMode
        Case QCH_Mode1
            CopyMemory udtQ1, udtQ, 12
        Case QCH_Mode2
            CopyMemory udtQ2, udtQ, 12
        Case QCH_Mode3
            CopyMemory udtQ3, udtQ, 12
    End Select

    ' check Q channel
    blnCRCValid = CRCCheck

    GetCh = True

End Function

' ported from CDR-DAO SubChannel.cc

Private Function CRCCheck() As Boolean

    Dim bytes(9)    As Byte
    Dim CRC         As Long

    ' CRC is optional
    If udtQ.CRC(0) = 0 And udtQ.CRC(1) = 0 Then
        CRCCheck = True
        Exit Function
    End If

    ' create CRC-16 hash
    CopyMemory bytes(0), udtQ, 10
    CRC = cCRC.CRC16A(bytes)

    CRCCheck = CRC = (cd.LShift(udtQ.CRC(0), 8) Or (udtQ.CRC(1)))

End Function

Private Sub DecodeISRC()

    Dim d   As Byte

    With udtQ3

        d = cd.RShift(.ISRC(0), 2) And &H3F
        strISRCC(0) = isrc2ascii(d)

        d = cd.LShift(.ISRC(0) And &H3, 4) Or (cd.RShift(.ISRC(1), 4) And &HF)
        strISRCC(1) = isrc2ascii(d)

        d = cd.LShift(.ISRC(1) And &HF, 2) Or (cd.RShift(.ISRC(2), 6) And &H3)
        strISRCO(0) = isrc2ascii(d)

        d = .ISRC(2) And &H3F
        strISRCO(1) = isrc2ascii(d)

        d = cd.RShift(.ISRC(3), 2) And &H3F
        strISRCO(2) = isrc2ascii(d)

        strISRCY(0) = Chr$((cd.RShift(.ISRC(4), 4) And &HF) + Asc("0"))
        strISRCY(1) = Chr$((.ISRC(4) And &HF) + Asc("0"))
        strISRCS(0) = Chr$((cd.RShift(.ISRC(5), 4) And &HF) + Asc("0"))
        strISRCS(1) = Chr$((.ISRC(5) And &HF) + Asc("0"))
        strISRCS(2) = Chr$((cd.RShift(.ISRC(6), 4) And &HF) + Asc("0"))
        strISRCS(3) = Chr$((.ISRC(6) And &HF) + Asc("0"))
        strISRCS(4) = Chr$((cd.RShift(.ISRC(7), 4) And &HF) + Asc("0"))

    End With

End Sub

Private Sub DecodeMCN()

    With udtQ2

         strMCN(0) = Chr$((cd.RShift(.MCN(0), 4) And &HF) + Asc("0"))
         strMCN(1) = Chr$((.MCN(0) And &HF) + Asc("0"))
         strMCN(2) = Chr$((cd.RShift(.MCN(1), 4) And &HF) + Asc("0"))
         strMCN(3) = Chr$((.MCN(1) And &HF) + Asc("0"))
         strMCN(4) = Chr$((cd.RShift(.MCN(2), 4) And &HF) + Asc("0"))
         strMCN(5) = Chr$((.MCN(2) And &HF) + Asc("0"))
         strMCN(6) = Chr$((cd.RShift(.MCN(3), 4) And &HF) + Asc("0"))
         strMCN(7) = Chr$((.MCN(3) And &HF) + Asc("0"))
         strMCN(8) = Chr$((cd.RShift(.MCN(4), 4) And &HF) + Asc("0"))
         strMCN(9) = Chr$((.MCN(4) And &HF) + Asc("0"))
        strMCN(10) = Chr$((cd.RShift(.MCN(5), 4) And &HF) + Asc("0"))
        strMCN(11) = Chr$((.MCN(5) And &HF) + Asc("0"))
        strMCN(12) = Chr$((cd.RShift(.MCN(6), 4) And &HF) + Asc("0"))

    End With

End Sub

Private Function bcd(d As Integer) As Byte

    If d >= 0 And d <= 99 Then
        bcd = cd.LShift((d / 10), 4) Or (d Mod 10)
    Else
        bcd = d
    End If

End Function

Private Function bcd2int(d As Byte) As Integer

    Dim d1  As Byte, d2     As Byte

    d1 = d And &HF
    d2 = cd.RShift(d, 4)

    If (d1 <= 9) And (d2 <= 9) Then
        bcd2int = d2 * 10 + d1
    Else
        bcd2int = d
    End If

End Function

Private Function isBcd(d As Byte) As Boolean

    If ((d And &HF) <= 9) And (cd.RShift(d, 4) <= 9) Then
        isBcd = True
    End If

End Function

Private Function isrc2ascii(c As Byte) As String

    If c <= 9 Then
        isrc2ascii = Chr$(Asc("0") + c)
    ElseIf (c >= 17 And c <= 42) Then
        isrc2ascii = Chr$(Asc("A") + (c - 17))
    End If

End Function
