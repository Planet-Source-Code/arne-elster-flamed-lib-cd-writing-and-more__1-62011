VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMP3Header"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************
'*** MP3 header reader
'***
'*** from ActiveVB
'*******************************************

Private anBitrateLookup(7, 15) As Integer
Private alFreqLookup(3, 7) As Long
Private avFrameRates(3)

'#####################################################################

Private m_sFilePath As String
Private m_lFileSize As Long
Private m_bytVersion As Byte
Private m_bytLayer As Byte
Private m_bCRCProtected As Boolean
Private m_lBitrate As Long
Private m_lFrequency As Long
Private m_bPadding As Boolean
Private m_bPrivateBit As Boolean
Private m_bytChannelMode As Byte
Private m_bytChannelModeExtention As Byte
Private m_bCopyright As Boolean
Private m_bOriginal As Boolean
Private m_bytEmphasis As Byte
Private m_sVersionText As String
Private m_sLayerText As String
Private m_sChannelModeText As String
Private m_sEmphasisText As String
Private m_lFrameSize As Long
Private m_lFrames As Long
Private m_lSeconds As Long
Private m_ID3V2 As Boolean
Private m_ID3V1 As Boolean
Private m_HeaderPosition As Long

'#####################################################################

Public Property Get FilePath() As String
  FilePath = m_sFilePath
End Property

Public Property Let FilePath(sValue As String)
  Call ZeroValues
  m_sFilePath = sValue
End Property

Public Property Get FileSize() As String
  FileSize = m_lFileSize
End Property

Public Property Get Version() As Byte
  Version = m_bytVersion
End Property

Public Property Get layer() As Byte
  layer = m_bytLayer
End Property

Public Property Get CRCProtected() As Boolean
  CRCProtected = m_bCRCProtected
End Property

Public Property Get Bitrate() As Long
  Bitrate = m_lBitrate
End Property

Public Property Get Frequency() As Long
  Frequency = m_lFrequency
End Property

Public Property Get Padding() As Boolean
  Padding = m_bPadding
End Property

Public Property Get PrivateBit() As Boolean
  PrivateBit = m_bPrivateBit
End Property

Public Property Get ChannelMode() As Byte
  ChannelMode = m_bytChannelMode
End Property

Public Property Get ChannelModeExtention() As Byte
  ChannelModeExtention = m_bytChannelModeExtention
End Property

Public Property Get Copyright() As Boolean
  Copyright = m_bCopyright
End Property

Public Property Get Original() As Boolean
  Original = m_bOriginal
End Property

Public Property Get Emphasis() As Byte
  Emphasis = m_bytEmphasis
End Property

Public Property Get VersionText() As String
  VersionText = m_sVersionText
End Property

Public Property Get LayerText() As String
  LayerText = m_sLayerText
End Property

Public Property Get ChannelModeText() As String
  ChannelModeText = m_sChannelModeText
End Property

Public Property Get EmphasisText() As String
  EmphasisText = m_sEmphasisText
End Property

Public Property Get FrameSize() As Long
  FrameSize = m_lFrameSize
End Property

Public Property Get Frames() As Long
  Frames = m_lFrames
End Property

Public Property Get Seconds() As Long
  Seconds = m_lSeconds
End Property

Public Property Get HeaderPosition() As Long
  HeaderPosition = m_HeaderPosition
End Property

Public Property Get ID3V1() As Boolean
  ID3V1 = m_ID3V1
End Property

Public Property Get ID3V2() As Boolean
  ID3V2 = m_ID3V2
End Property

'#####################################################################

Private Sub Class_Initialize()
  Dim asBitrateCore() As String, asFreqCore() As String
  Dim sBitrateData As String, sFreqData As String
  Dim nBitRate As Integer, nVerLayer As Integer, nFreq As Integer

  sBitrateData = "999,999,999,999,999,999," & _
                 "032,032,032,032,008,008," & _
                 "064,048,040,048,016,016," & _
                 "096,056,048,056,024,024," & _
                 "128,064,056,064,032,032," & _
                 "160,080,064,080,040,040," & _
                 "192,096,080,096,048,048," & _
                 "224,112,096,112,056,056," & _
                 "256,128,112,128,064,064," & _
                 "288,160,128,144,080,080," & _
                 "320,192,160,160,096,096," & _
                 "352,224,192,176,112,112," & _
                 "384,256,224,192,128,128," & _
                 "416,320,256,224,144,144," & _
                 "448,384,320,256,160,160," & _
                 "999,999,999,999,999,999"

  asBitrateCore = Split(sBitrateData, ",")

  For nBitRate = 1 To 14
    For nVerLayer = 0 To 2
      anBitrateLookup(7 - nVerLayer, nBitRate) = _
      Val(asBitrateCore((nBitRate * 6) + nVerLayer))
    Next
    For nVerLayer = 0 To 2
      anBitrateLookup(3 - nVerLayer, nBitRate) = _
      Val(asBitrateCore((nBitRate * 6) + 3 + nVerLayer))
    Next
  Next

  sFreqData = "44100,22050,11025," & _
              "48000,24000,12000," & _
              "32000,16000,08000," & _
              "99999,99999,99999"

  asFreqCore = Split(sFreqData, ",")

  For nFreq = 0 To 3
    alFreqLookup(3, nFreq) = Val(asFreqCore((nFreq * 3)))
    alFreqLookup(2, nFreq) = Val(asFreqCore((nFreq * 3) + 1))
    alFreqLookup(0, nFreq) = Val(asFreqCore((nFreq * 3) + 2))
  Next

  avFrameRates(0) = 38.5
  avFrameRates(1) = 32.5
  avFrameRates(2) = 27.8
  avFrameRates(3) = 0
End Sub

'#####################################################################

Private Sub ZeroValues()
  m_lFileSize = 0
  m_bytVersion = 0
  m_bytLayer = 0
  m_bCRCProtected = False
  m_lBitrate = 0
  m_lFrequency = 0
  m_bPadding = False
  m_bPrivateBit = False
  m_bytChannelMode = 0
  m_bytChannelModeExtention = 0
  m_bCopyright = False
  m_bOriginal = False
  m_bytEmphasis = 0
  m_sVersionText = ""
  m_sLayerText = ""
  m_sChannelModeText = ""
  m_sEmphasisText = ""
  m_lFrameSize = 0
  m_lFrames = 0
  m_lSeconds = 0
End Sub

Public Function GetFileInfos()
  Dim nFile As Integer
  Dim i As Long, z As Integer
  Dim sInput As String, sMP3bitsString As String
  Dim nBit1 As Integer, nBit2 As Integer
  Dim nBitD1 As Integer, nBitD2 As Integer
  Dim dSHIFT, LayerType, FrameSize
  Dim mp3_ID1, mp3_bitrate, mp3_protection, mp3_frequency
  Dim aBytes(3) As Byte
  Dim ID3V2Len As Long
  Dim sID3Len As String
  Dim ID3Position As Long
  Dim HDPos As Long
  
  m_ID3V2 = False
  GetFileInfos = -1
  If Not FileExists(m_sFilePath) Then Exit Function
  GetFileInfos = 0
  nFile = FreeFile
  Open m_sFilePath For Binary As #nFile
    Seek #nFile, LOF(nFile) - 127
    sInput = Input(128, #nFile)
    If Left(sInput, 3) = "TAG" Then
      m_ID3V1 = True
    End If
    Seek #nFile, 1
    
    'Einlesen der ersten vier Kilobytes um
    'den Header der Datei zu finden
    sInput = Input(8192, #nFile)
    
    'Wird für die Berechnung der Trackduration benötigt
    m_lFileSize = LOF(nFile)
    
    ' Ist ein ID3V2-Tag vorhanden?
    If Left$(sInput, 3) = "ID3" Then
      ID3Position = 1
      m_ID3V2 = True
    End If
    
    If ID3Position Then
      
      ' Bytes mit Längen-Info des Tags lesen
      sID3Len = Mid$(sInput, ID3Position + 6, 4)
      
      ' länge des Tags berechnen
      ID3V2Len = &H200000 * Asc(Left$(sID3Len, 1)) + _
        &H4000 * Asc(Mid$(sID3Len, 2, 1)) + _
        &H80 * Asc(Mid$(sID3Len, 3, 1)) + _
        Asc(Mid$(sID3Len, 4, 1))
      
      ' Tag überspringen
      Seek #nFile, ID3Position + ID3V2Len + 10
      
      'wird benötigt zur Berechnung der Headerposition
      HDPos = ID3Position + ID3V2Len + 10
      m_lFileSize = m_lFileSize - (ID3Position + ID3V2Len + 10)
      
      ' neuen Einlesen
      sInput = Input(8192, #nFile)
    End If
  Close #nFile
  

  i = 0
  Do Until i = 8191
ReEnter:
    i = i + 1
    nBit1 = Asc(Mid(sInput, i, 1))
    nBit2 = Asc(Mid(sInput, i + 1, 1))
    If nBit1 = &HFF And (nBit2 And &HE0) = &HE0 Then
      
      '20 HeadersBits auslesen - es sind die
      'letzen 20 Bits der nexten 3 Bytes
      sMP3bitsString = Mid(sInput, i + 1, 3)
      m_HeaderPosition = HDPos + i - 1
      Exit Do
    End If
    
    'Wir haben die Sync nicht gefunden, deshalb
    'verschieben wir das ganze um 4Bits nach links
    dSHIFT = ShiftBits(Mid(sInput, i, 3))
    nBitD1 = Asc(Left(dSHIFT, 1))
    nBitD2 = Asc(Right(dSHIFT, 1))
    
    If nBitD1 = &HFF And (nBitD2 And &HE0) = &HE0 Then
      '20 HeaderBits auslesen - es sind die
      'ersten 20 Bits der nexten 3 Bytes
      sMP3bitsString = Mid(sInput, i + 2, 3)
      m_HeaderPosition = HDPos + i - 1
      Exit Do
    End If
  Loop

  If i = 8191 Then Exit Function 'Header wurde nicht gefunden!
                                 ' -> beenden der Routine

  For z = 1 To 3
    aBytes(z) = Asc(Mid(sMP3bitsString, z))
  Next

  'Die ersten 20 Bits von sMP3bitsString sind die
  'Headerinformationen für diesen Frame
  '1te Bit: ID | 0 = MPEG-2 | 1 = MPEG-1
  m_bytVersion = (&H18 And aBytes(1)) / 8
  mp3_ID1 = (m_bytVersion And 1)
  
  'folgende 2 Bits sind der Layer
  m_bytLayer = (&H6 And aBytes(1)) / 2
  
  'folgendes Bit ist Protection
  mp3_protection = &H1 And aBytes(1)
  m_bCRCProtected = mp3_protection <> 0
  
  'folgende 4 Bits sind die Bitrate
  mp3_bitrate = (&HF0 And aBytes(2)) / 16
  LayerType = (mp3_ID1 * 4) Or m_bytLayer
  m_lBitrate = 1000 * CLng((anBitrateLookup(LayerType, mp3_bitrate)))
  
  'folgende 2 Bits sind die Frequenz
  mp3_frequency = (&HC And aBytes(2)) / 4
  m_lFrequency = alFreqLookup(m_bytVersion, mp3_frequency)
  If m_lFrequency = 99999 Or m_lFrequency = 0 Or m_lBitrate = 0 Then
    i = i + 4
    GoTo ReEnter
  End If
  
  'folgendes Bit ist das Padding Bit
  m_bPadding = ((&H2 And aBytes(2)) / 2) = 1
  
  'folgendes Bit ist das Private Bit
  m_bPrivateBit = ((&H10 And aBytes(3)) / 2) = 1
  
  'folgende 2 Bit sind der Channel mode
  m_bytChannelMode = (&HC0 And aBytes(3)) / 64
  
  'folgende 2 Bits sind die Channel Mode Extention
  m_bytChannelModeExtention = (&H30 And aBytes(3)) / 16
  
  'folgendes Bit ist der Copyright Flag
  m_bCopyright = ((&H8 And aBytes(3)) / 8) = 1
  
  'folgendes Bit ist das Original Flag
  m_bOriginal = ((&H4 And aBytes(3)) / 4) = 1
  
  'folgendes Bit ist das Emphasis Flag
  m_bytEmphasis = &H3 And aBytes(3)

  Select Case m_bytVersion
    Case 0
      m_sVersionText = "MPEG-2.5"
    Case 1
    Case 2
      m_sVersionText = "MPEG-2.0"
    Case 3
      m_sVersionText = "MPEG-1.0"
  End Select

  Select Case m_bytLayer
    Case 1
      m_sLayerText = "Layer III"
      FrameSize = (144 * (m_lBitrate / m_lFrequency))
    Case 2
      m_sLayerText = "Layer II"
      FrameSize = (144 * (m_lBitrate / m_lFrequency))
    Case 3
      m_sLayerText = "Layer I"
      FrameSize = ((12 * (m_lBitrate / m_lFrequency) + _
      Abs(m_bPadding))) * 4
  End Select

  Select Case m_bytChannelMode
    Case 0
      m_sChannelModeText = "Stereo"
    Case 1
      m_sChannelModeText = "Joint Stereo (Stereo)"
      If m_bytVersion < 3 Then FrameSize = Fix(FrameSize) / 2
      If m_bytVersion = 0 Then FrameSize = Fix(FrameSize) / 2
    Case 2
      m_sChannelModeText = "Dual Channel (Stereo)"
    Case 3
      m_sChannelModeText = "Single Channel (Mono)"
      If m_bytVersion < 3 Then FrameSize = Fix(FrameSize) / 2
  End Select

  Select Case m_bytEmphasis
    Case 0
      m_sEmphasisText = "None"
    Case 1
      m_sEmphasisText = "50/15 ms"
    Case 2
      m_sEmphasisText = "reserved"
    Case 3
      m_sEmphasisText = "CIT J.17"
  End Select

  'Ausrechnen der Frameanzahl und der Spieldauer
  m_lFrameSize = Fix(FrameSize)
  m_lFrames = m_lFileSize / Fix(FrameSize)
  m_lSeconds = m_lFrames / avFrameRates(mp3_frequency)
End Function

Private Function ShiftBits(sInput As String) As String
  Dim nSD1, nSD2, nSD3, nDO1, nDO2 As Integer
  
  nSD1 = Asc(Left(sInput, 1))
  nSD2 = Asc(Mid(sInput, 2, 1))
  nSD3 = Asc(Right(sInput, 1))

  nDO1 = ((nSD1 And &HF) * 16) Or ((nSD2 And &HF0) / 16)
  nDO2 = ((nSD2 And &HF) * 16) Or ((nSD3 And &HF0) / 16)
  ShiftBits = Chr(nDO1) + Chr(nDO2)
End Function

Private Function FileExists(ByVal FileName As String) As Boolean
  On Error Resume Next
  FileExists = (GetAttr(FileName) >= vbNormal)
End Function
