VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' DOESNT WORK?

' from some mvp page...
Public Function CRC16A(Buffer() As Byte) As Long

    Dim I       As Long, Temp As Long
    Dim J       As Integer
    Dim CRC     As Long

    For I = 0 To UBound(Buffer)

        Temp = Buffer(I) * &H100&
        CRC = CRC Xor Temp

        For J = 0 To 7
            If (CRC And &H8000&) Then
                CRC = ((CRC * 2) Xor &H1021&) And &HFFFF&
            Else
                CRC = (CRC * 2) And &HFFFF&
            End If
        Next J

    Next I
    CRC16A = CRC And &HFFFF

End Function


'Public Function DoCRC16(bytes() As Byte) As Long
'
'    Dim CRC As Long
'    Dim I   As Long
'
'    For I = 0 To UBound(bytes)
'
'        CRC = m_CRC16Table((cd.RShift(CRC, 8) And &HFFF&) Xor bytes(I))
'        CRC = (cd.LShift(CRC, 8) And &HFFF&) Xor (CRC And &HFFF&)
'
'    Next
'
'    DoCRC16 = CRC
'
'End Function
'
'Private Sub InitializeCRC16()
'
'    m_CRC16Table(0) = &H0
'    m_CRC16Table(1) = &H1021
'    m_CRC16Table(2) = &H2042
'    m_CRC16Table(3) = &H3063
'    m_CRC16Table(4) = &H4084
'    m_CRC16Table(5) = &H50A5
'    m_CRC16Table(6) = &H60C6
'    m_CRC16Table(7) = &H70E7
'    m_CRC16Table(8) = &H8108
'    m_CRC16Table(9) = &H9129
'    m_CRC16Table(10) = &HA14A
'    m_CRC16Table(11) = &HB16B
'    m_CRC16Table(12) = &HC18C
'    m_CRC16Table(13) = &HD1AD
'    m_CRC16Table(14) = &HE1CE
'    m_CRC16Table(15) = &HF1EF
'    m_CRC16Table(16) = &H1231
'    m_CRC16Table(17) = &H210
'    m_CRC16Table(18) = &H3273
'    m_CRC16Table(19) = &H2252
'    m_CRC16Table(20) = &H52B5
'    m_CRC16Table(21) = &H4294
'    m_CRC16Table(22) = &H72F7
'    m_CRC16Table(23) = &H62D6
'    m_CRC16Table(24) = &H9339
'    m_CRC16Table(25) = &H8318
'    m_CRC16Table(26) = &HB37B
'    m_CRC16Table(27) = &HA35A
'    m_CRC16Table(28) = &HD3BD
'    m_CRC16Table(29) = &HC39C
'    m_CRC16Table(30) = &HF3FF
'    m_CRC16Table(31) = &HE3DE
'    m_CRC16Table(32) = &H2462
'    m_CRC16Table(33) = &H3443
'    m_CRC16Table(34) = &H420
'    m_CRC16Table(35) = &H1401
'    m_CRC16Table(36) = &H64E6
'    m_CRC16Table(37) = &H74C7
'    m_CRC16Table(38) = &H44A4
'    m_CRC16Table(39) = &H5485
'    m_CRC16Table(40) = &HA56A
'    m_CRC16Table(41) = &HB54B
'    m_CRC16Table(42) = &H8528
'    m_CRC16Table(43) = &H9509
'    m_CRC16Table(44) = &HE5EE
'    m_CRC16Table(45) = &HF5CF
'    m_CRC16Table(46) = &HC5AC
'    m_CRC16Table(47) = &HD58D
'    m_CRC16Table(48) = &H3653
'    m_CRC16Table(49) = &H2672
'    m_CRC16Table(50) = &H1611
'    m_CRC16Table(51) = &H630
'    m_CRC16Table(52) = &H76D7
'    m_CRC16Table(53) = &H66F6
'    m_CRC16Table(54) = &H5695
'    m_CRC16Table(55) = &H46B4
'    m_CRC16Table(56) = &HB75B
'    m_CRC16Table(57) = &HA77A
'    m_CRC16Table(58) = &H9719
'    m_CRC16Table(59) = &H8738
'    m_CRC16Table(60) = &HF7DF
'    m_CRC16Table(61) = &HE7FE
'    m_CRC16Table(62) = &HD79D
'    m_CRC16Table(63) = &HC7BC
'    m_CRC16Table(64) = &H48C4
'    m_CRC16Table(65) = &H58E5
'    m_CRC16Table(66) = &H6886
'    m_CRC16Table(67) = &H78A7
'    m_CRC16Table(68) = &H840
'    m_CRC16Table(69) = &H1861
'    m_CRC16Table(70) = &H2802
'    m_CRC16Table(71) = &H3823
'    m_CRC16Table(72) = &HC9CC
'    m_CRC16Table(73) = &HD9ED
'    m_CRC16Table(74) = &HE98E
'    m_CRC16Table(75) = &HF9AF
'    m_CRC16Table(76) = &H8948
'    m_CRC16Table(77) = &H9969
'    m_CRC16Table(78) = &HA90A
'    m_CRC16Table(79) = &HB92B
'    m_CRC16Table(80) = &H5AF5
'    m_CRC16Table(81) = &H4AD4
'    m_CRC16Table(82) = &H7AB7
'    m_CRC16Table(83) = &H6A96
'    m_CRC16Table(84) = &H1A71
'    m_CRC16Table(85) = &HA50
'    m_CRC16Table(86) = &H3A33
'    m_CRC16Table(87) = &H2A12
'    m_CRC16Table(88) = &HDBFD
'    m_CRC16Table(89) = &HCBDC
'    m_CRC16Table(90) = &HFBBF
'    m_CRC16Table(91) = &HEB9E
'    m_CRC16Table(92) = &H9B79
'    m_CRC16Table(93) = &H8B58
'    m_CRC16Table(94) = &HBB3B
'    m_CRC16Table(95) = &HAB1A
'    m_CRC16Table(96) = &H6CA6
'    m_CRC16Table(97) = &H7C87
'    m_CRC16Table(98) = &H4CE4
'    m_CRC16Table(99) = &H5CC5
'    m_CRC16Table(100) = &H2C22
'    m_CRC16Table(101) = &H3C03
'    m_CRC16Table(102) = &HC60
'    m_CRC16Table(103) = &H1C41
'    m_CRC16Table(104) = &HEDAE
'    m_CRC16Table(105) = &HFD8F
'    m_CRC16Table(106) = &HCDEC
'    m_CRC16Table(107) = &HDDCD
'    m_CRC16Table(108) = &HAD2A
'    m_CRC16Table(109) = &HBD0B
'    m_CRC16Table(110) = &H8D68
'    m_CRC16Table(111) = &H9D49
'    m_CRC16Table(112) = &H7E97
'    m_CRC16Table(113) = &H6EB6
'    m_CRC16Table(114) = &H5ED5
'    m_CRC16Table(115) = &H4EF4
'    m_CRC16Table(116) = &H3E13
'    m_CRC16Table(117) = &H2E32
'    m_CRC16Table(118) = &H1E51
'    m_CRC16Table(119) = &HE70
'    m_CRC16Table(120) = &HFF9F
'    m_CRC16Table(121) = &HEFBE
'    m_CRC16Table(122) = &HDFDD
'    m_CRC16Table(123) = &HCFFC
'    m_CRC16Table(124) = &HBF1B
'    m_CRC16Table(125) = &HAF3A
'    m_CRC16Table(126) = &H9F59
'    m_CRC16Table(127) = &H8F78
'    m_CRC16Table(128) = &H9188
'    m_CRC16Table(129) = &H81A9
'    m_CRC16Table(130) = &HB1CA
'    m_CRC16Table(131) = &HA1EB
'    m_CRC16Table(132) = &HD10C
'    m_CRC16Table(133) = &HC12D
'    m_CRC16Table(134) = &HF14E
'    m_CRC16Table(135) = &HE16F
'    m_CRC16Table(136) = &H1080
'    m_CRC16Table(137) = &HA1
'    m_CRC16Table(138) = &H30C2
'    m_CRC16Table(139) = &H20E3
'    m_CRC16Table(140) = &H5004
'    m_CRC16Table(141) = &H4025
'    m_CRC16Table(142) = &H7046
'    m_CRC16Table(143) = &H6067
'    m_CRC16Table(144) = &H83B9
'    m_CRC16Table(145) = &H9398
'    m_CRC16Table(146) = &HA3FB
'    m_CRC16Table(147) = &HB3DA
'    m_CRC16Table(148) = &HC33D
'    m_CRC16Table(149) = &HD31C
'    m_CRC16Table(150) = &HE37F
'    m_CRC16Table(151) = &HF35E
'    m_CRC16Table(152) = &H2B1
'    m_CRC16Table(153) = &H1290
'    m_CRC16Table(154) = &H22F3
'    m_CRC16Table(155) = &H32D2
'    m_CRC16Table(156) = &H4235
'    m_CRC16Table(157) = &H5214
'    m_CRC16Table(158) = &H6277
'    m_CRC16Table(159) = &H7256
'    m_CRC16Table(160) = &HB5EA
'    m_CRC16Table(161) = &HA5CB
'    m_CRC16Table(162) = &H95A8
'    m_CRC16Table(163) = &H8589
'    m_CRC16Table(164) = &HF56E
'    m_CRC16Table(165) = &HE54F
'    m_CRC16Table(166) = &HD52C
'    m_CRC16Table(167) = &HC50D
'    m_CRC16Table(168) = &H34E2
'    m_CRC16Table(169) = &H24C3
'    m_CRC16Table(170) = &H14A0
'    m_CRC16Table(171) = &H481
'    m_CRC16Table(172) = &H7466
'    m_CRC16Table(173) = &H6447
'    m_CRC16Table(174) = &H5424
'    m_CRC16Table(175) = &H4405
'    m_CRC16Table(176) = &HA7DB
'    m_CRC16Table(177) = &HB7FA
'    m_CRC16Table(178) = &H8799
'    m_CRC16Table(179) = &H97B8
'    m_CRC16Table(180) = &HE75F
'    m_CRC16Table(181) = &HF77E
'    m_CRC16Table(182) = &HC71D
'    m_CRC16Table(183) = &HD73C
'    m_CRC16Table(184) = &H26D3
'    m_CRC16Table(185) = &H36F2
'    m_CRC16Table(186) = &H691
'    m_CRC16Table(187) = &H16B0
'    m_CRC16Table(188) = &H6657
'    m_CRC16Table(189) = &H7676
'    m_CRC16Table(190) = &H4615
'    m_CRC16Table(191) = &H5634
'    m_CRC16Table(192) = &HD94C
'    m_CRC16Table(193) = &HC96D
'    m_CRC16Table(194) = &HF90E
'    m_CRC16Table(195) = &HE92F
'    m_CRC16Table(196) = &H99C8
'    m_CRC16Table(197) = &H89E9
'    m_CRC16Table(198) = &HB98A
'    m_CRC16Table(199) = &HA9AB
'    m_CRC16Table(200) = &H5844
'    m_CRC16Table(201) = &H4865
'    m_CRC16Table(202) = &H7806
'    m_CRC16Table(203) = &H6827
'    m_CRC16Table(204) = &H18C0
'    m_CRC16Table(205) = &H8E1
'    m_CRC16Table(206) = &H3882
'    m_CRC16Table(207) = &H28A3
'    m_CRC16Table(208) = &HCB7D
'    m_CRC16Table(209) = &HDB5C
'    m_CRC16Table(210) = &HEB3F
'    m_CRC16Table(211) = &HFB1E
'    m_CRC16Table(212) = &H8BF9
'    m_CRC16Table(213) = &H9BD8
'    m_CRC16Table(214) = &HABBB
'    m_CRC16Table(215) = &HBB9A
'    m_CRC16Table(216) = &H4A75
'    m_CRC16Table(217) = &H5A54
'    m_CRC16Table(218) = &H6A37
'    m_CRC16Table(219) = &H7A16
'    m_CRC16Table(220) = &HAF1
'    m_CRC16Table(221) = &H1AD0
'    m_CRC16Table(222) = &H2AB3
'    m_CRC16Table(223) = &H3A92
'    m_CRC16Table(224) = &HFD2E
'    m_CRC16Table(225) = &HED0F
'    m_CRC16Table(226) = &HDD6C
'    m_CRC16Table(227) = &HCD4D
'    m_CRC16Table(228) = &HBDAA
'    m_CRC16Table(229) = &HAD8B
'    m_CRC16Table(230) = &H9DE8
'    m_CRC16Table(231) = &H8DC9
'    m_CRC16Table(232) = &H7C26
'    m_CRC16Table(233) = &H6C07
'    m_CRC16Table(234) = &H5C64
'    m_CRC16Table(235) = &H4C45
'    m_CRC16Table(236) = &H3CA2
'    m_CRC16Table(237) = &H2C83
'    m_CRC16Table(238) = &H1CE0
'    m_CRC16Table(239) = &HCC1
'    m_CRC16Table(240) = &HEF1F
'    m_CRC16Table(241) = &HFF3E
'    m_CRC16Table(242) = &HCF5D
'    m_CRC16Table(243) = &HDF7C
'    m_CRC16Table(244) = &HAF9B
'    m_CRC16Table(245) = &HBFBA
'    m_CRC16Table(246) = &H8FD9
'    m_CRC16Table(247) = &H9FF8
'    m_CRC16Table(248) = &H6E17
'    m_CRC16Table(249) = &H7E36
'    m_CRC16Table(250) = &H4E55
'    m_CRC16Table(251) = &H5E74
'    m_CRC16Table(252) = &H2E93
'    m_CRC16Table(253) = &H3EB2
'    m_CRC16Table(254) = &HED1
'    m_CRC16Table(255) = &H1EF0
'
'End Sub
'
'Private Sub Class_Initialize()
'    InitializeCRC16
'End Sub
