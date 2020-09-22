VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_DecoderWAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************
'*** Flamed WAV decoder using the ACM
'*******************************************

Private wfxIN   As WAVEFORMATEX
Private wfxOUT  As WAVEFORMATEX

Private ashBUF  As ACMSTREAMHEADER

Private bytBUF() As Byte

Private hStream As Long

Public Property Get InputBlockAlign() As Integer
    InputBlockAlign = wfxIN.nBlockAlign
End Property

Public Property Let InputBlockAlign(aval As Integer)
    wfxIN.nBlockAlign = aval
End Property

Public Property Get InputChannels() As Integer
    InputChannels = wfxIN.nChannels
End Property

Public Property Let InputChannels(aval As Integer)
    With wfxIN
        .nChannels = aval
        .nAvgBytesPerSec = .nSamplesPerSec * .wBitsPerSample * 1 / 8 * .nChannels
    End With
End Property

Public Property Get InputBitsPerSample() As Integer
    InputBitsPerSample = wfxIN.wBitsPerSample
End Property

Public Property Let InputBitsPerSample(aval As Integer)
    With wfxIN
        .wBitsPerSample = aval
        .nAvgBytesPerSec = .nSamplesPerSec * .wBitsPerSample * 1 / 8 * .nChannels
    End With
End Property

Public Property Get InputSampleRate() As Long
    InputSampleRate = wfxOUT.nSamplesPerSec
End Property

Public Property Let InputSampleRate(aval As Long)
    With wfxIN
        .nSamplesPerSec = aval
        .nAvgBytesPerSec = .nSamplesPerSec * .wBitsPerSample * 1 / 8 * .nChannels
    End With
End Property

Public Property Get StreamHandle() As Long
    StreamHandle = hStream
End Property

Public Property Get FileExtension() As String
    FileExtension = "wav"
End Property

Public Function Init(InputBuffer() As Byte, _
                     InputLen As Long) As Boolean

    Dim lngRet          As Long
    Dim OutputLength    As Long

    ' open a stream
    lngRet = acmStreamOpen(hStream, _
                           0, _
                           wfxIN, _
                           wfxOUT, _
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

End Function

Public Function Decode(ByVal InLen As Long, Optional blnLast As Boolean = False) As Byte()

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

        Decode = btOut

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
        .cbSize = 0
        .nBlockAlign = 4
        .nChannels = 2
        .nSamplesPerSec = 44100
        .wBitsPerSample = 16
        .nAvgBytesPerSec = .nSamplesPerSec * .wBitsPerSample * 1 / 8 * .nChannels
        .wFormatTag = WAVE_FORMAT_PCM
    End With

End Sub
