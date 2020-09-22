VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_SubChannel96"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' RAW Sub-Channel data
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      '  7     6     5     4     3     2     1     0 '
' Byte |  P  |  Q  |  R  |  S  |  T  |  U  |  V  |  W '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  0   |  P1    Q1    R1    S1    T1    U1    V1    W1'
'  1   |  P2    Q2    R2    S2    T2    U2    V2    W2'
' ...  |                      ...                     '
' 94   |                                              '
' 95   |                                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Interleaved Sub-Channel data
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Byte |      '     '     '     '     '     '     '      '
'  0   |  P1    P2     P3    P4    P5    P6    P7    P8  |
' ...  |  ...                                            |
' 11   |  P89   P90    P91   P92   P93   P94   P95   P96 |
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  0   |  Q1    Q2     Q3    Q4    Q5    Q6    Q7    Q8  |
' ...  |  ...                                            |
' 11   |  Q89   Q90    Q91   Q92   Q93   Q94   Q95   Q96 |
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'...                                                     '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Q sub-channel modes
Public Enum FL_SubChannelQModes
    QCH_Mode1 = &H1     ' current position
    QCH_Mode2 = &H2     ' MCN
    QCH_Mode3 = &H3     ' ISRC
    QCH_Others = &H4    ' unsupported
End Enum

' sub-channel columns
Public Enum FL_SubChannelLetters
    CH_P = 7            ' P channel
    CH_Q = 6            ' Q channel
    CH_R = 5            ' ...
    CH_S = 4
    CH_T = 3
    CH_U = 2            ' ...
    CH_V = 1            ' V channel
    CH_W = 0            ' W channel
End Enum

' NOT YET TESTED!
Public Function InterleaveRAW96RWPackets(Data() As Byte, _
                offset As Long) As Byte()

    Dim ret(95) As Byte
    Dim I       As Long

    For I = offset To offset + 95
        ret(I - offset) = Data(I) And &H3F
    Next

    InterleaveRAW96RWPackets = ret

End Function

Public Function InterleaveRAW96QCh(Data() As Byte, _
                ByVal offset As Long) As Byte()

    Dim I       As Long
    Dim ndx     As Integer
    Dim ret(11) As Byte

    For I = offset To offset + 95 Step 8

        ndx = (I - offset) \ 8
        ret(ndx) = ret(ndx) Or cd.LShift(Abs(cd.IsBitSet(Data(I + 0), CH_Q)), 7)
        ret(ndx) = ret(ndx) Or cd.LShift(Abs(cd.IsBitSet(Data(I + 1), CH_Q)), 6)
        ret(ndx) = ret(ndx) Or cd.LShift(Abs(cd.IsBitSet(Data(I + 2), CH_Q)), 5)
        ret(ndx) = ret(ndx) Or cd.LShift(Abs(cd.IsBitSet(Data(I + 3), CH_Q)), 4)
        ret(ndx) = ret(ndx) Or cd.LShift(Abs(cd.IsBitSet(Data(I + 4), CH_Q)), 3)
        ret(ndx) = ret(ndx) Or cd.LShift(Abs(cd.IsBitSet(Data(I + 5), CH_Q)), 2)
        ret(ndx) = ret(ndx) Or cd.LShift(Abs(cd.IsBitSet(Data(I + 6), CH_Q)), 1)
        ret(ndx) = ret(ndx) Or cd.LShift(Abs(cd.IsBitSet(Data(I + 7), CH_Q)), 0)

    Next

    InterleaveRAW96QCh = ret

End Function
