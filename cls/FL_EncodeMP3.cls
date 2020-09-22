VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_EncoderMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************
'*** Flamed MP3 encoder using the ACM
'*******************************************

Public Enum FL_MP3ENC_BITRATES
     [96 kBits] = 12000
    [112 kBits] = 14000
    [128 kBits] = 16000
    [160 kBits] = 20000
    [192 kBits] = 24000
    [224 kBits] = 28000
    [256 kBits] = 32000
    [320 kBits] = 40000
End Enum

Private wfxIN   As WAVEFORMATEX
Private wfxOUT  As WAVEFORMATEX

Private ashBUF  As ACMSTREAMHEADER
Private bytBUF() As Byte

Private hStream As Long

Public Property Get OutputBitrate() As FL_MP3ENC_BITRATES
    OutputBitrate = wfxOUT.nAvgBytesPerSec
End Property

Public Property Let OutputBitrate(aval As FL_MP3ENC_BITRATES)
    wfxOUT.nAvgBytesPerSec = aval
End Property

Public Property Get StreamHandle() As Long
    StreamHandle = hStream
End Property

Public Property Get FileExtension() As String
    FileExtension = "mp3"
End Property

Public Function Init(InputBuffer() As Byte, _
                     InputLen As Long) As Boolean

    On Error GoTo ErrorHandler

    Dim lngRet          As Long
    Dim OutputLength    As Long
    Dim hDrv            As Long

    Dim btWFX()         As Byte

    ' take the first best wfx structure
    GetCompatibleCodecs wfxIN, wfxOUT
    With udtFoundFormats(0)
        btWFX = .btWaveFormatEx
        hDrv = .haid
    End With

    ' open a stream
    lngRet = acmStreamOpen(hStream, _
                           0, _
                           wfxIN, _
                           btWFX(0), _
                           0, _
                           0, _
                           0, _
                           0)

    If lngRet <> MMSYSERR_NOERROR Then Exit Function

    ' get the stream output size
    lngRet = acmStreamSize(hStream, InputLen, OutputLength, ACM_STREAMSIZEF_SOURCE)
    If lngRet <> MMSYSERR_NOERROR Then Exit Function

    ' prepare the conversion buffer
    ReDim bytBUF(OutputLength - 1) As Byte

    With ashBUF
        .cbStruct = Len(ashBUF)
        .cbSrcLength = InputLen
        .pbSrc = VarPtr(InputBuffer(0))
        .cbDstLength = OutputLength
        .pbDst = VarPtr(bytBUF(0))
    End With

    If MMSYSERR_NOERROR <> acmStreamPrepareHeader(hStream, ashBUF, 0) Then
        DeInit
        Exit Function
    End If

    Init = True

ErrorHandler:

End Function

Public Function Encode(ByVal InLen As Long, Optional blnLast As Boolean = False) As Byte()

    Dim btOut() As Byte
    Dim lngOut  As Long
    Dim lngRet  As Long

    ashBUF.cbSrcLength = InLen

    ' convert buffer
    lngRet = acmStreamConvert(hStream, ashBUF, ACM_STREAMCONVERTF_BLOCKALIGN)
    If lngRet <> MMSYSERR_NOERROR Then Exit Function

    ' something got converted?
    If ashBUF.cbDstLengthUsed > 0 Then

        lngOut = ashBUF.cbDstLengthUsed

        ReDim btOut(lngOut - 1) As Byte
        CopyMemory btOut(0), ByVal ashBUF.pbDst, ashBUF.cbDstLengthUsed

        ' last conversion, get everything out of the stream
        If blnLast Then

            lngRet = acmStreamConvert(hStream, ashBUF, ACM_STREAMCONVERTF_END)
            If lngRet <> MMSYSERR_NOERROR Then Exit Function

            ReDim Preserve btOut(lngOut + ashBUF.cbDstLengthUsed - 1) As Byte
            CopyMemory btOut(lngOut), ByVal ashBUF.pbDst, ashBUF.cbDstLengthUsed

        End If

        Encode = btOut

    End If

End Function

Public Function DeInit() As Boolean
    DeInit = DeInit And (MMSYSERR_NOERROR = acmStreamUnprepareHeader(hStream, ashBUF, 0))
    DeInit = DeInit And (MMSYSERR_NOERROR = acmStreamClose(hStream, 0))
    hStream = 0
End Function

Private Sub Class_Initialize()

    ' red book audio
    With wfxIN
        .cbSize = 0
        .nBlockAlign = 4
        .nChannels = 2
        .nSamplesPerSec = 44100
        .wBitsPerSample = 16
        .nAvgBytesPerSec = .nSamplesPerSec * .wBitsPerSample * 1 / 8 * .nChannels
        .wFormatTag = WAVE_FORMAT_PCM
    End With

    With wfxOUT
        .nChannels = wfxIN.nChannels
        .nSamplesPerSec = wfxIN.nSamplesPerSec
        .wFormatTag = WAVE_FORMAT_MPEGLAYER3
    End With

    OutputBitrate = [192 kBits]

End Sub
