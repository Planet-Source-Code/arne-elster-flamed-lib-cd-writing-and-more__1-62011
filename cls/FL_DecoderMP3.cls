VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_DecoderMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************
'*** Flamed MP3 decoder using the ACM
'*******************************************

Private wfxIN   As MPEGLAYER3WAVEFORMAT
Private wfxOUT  As WAVEFORMATEX

Private ashBUF  As ACMSTREAMHEADER
Private bytBUF() As Byte

Private hStream As Long

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

    Dim i               As Integer
    Dim j               As Integer

    Dim btWFX()         As Byte
    Dim wfx             As WAVEFORMATEX

    For i = 0 To acmDrivers.count - 1
        With acmDrivers.drivers(i)
            For j = 0 To .FormatTagCount - 1
                With .FormatTag(j)
                    If .FormatTag = WAVE_FORMAT_MPEGLAYER3 Then
                        If .wfx.nSamplesPerSec = wfxOUT.nSamplesPerSec Then
                            If .wfx.nChannels = wfxOUT.nChannels Then
                                wfx = .wfx
                                GoTo ExitLoops
                            End If
                        End If
                    End If
                End With
            Next
        End With
    Next

ExitLoops:

    ' open a stream
    lngRet = acmStreamOpen(hStream, _
                           0, _
                           wfx, _
                           wfxOUT, _
                           0, _
                           0, _
                           0, _
                           0)

    If lngRet <> MMSYSERR_NOERROR Then Exit Function

    ' get the stream output size
    lngRet = acmStreamSize(hStream, InputLen, OutputLength, ACM_STREAMSIZEF_SOURCE)
    If lngRet <> MMSYSERR_NOERROR Then Exit Function

    OutputLength = OutputLength + 3 * OutputLength / 2

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

Public Property Get InputChannels() As Integer
    InputChannels = wfxIN.wfx.nChannels
End Property

Public Property Let InputChannels(aval As Integer)
    wfxIN.wfx.nChannels = aval
End Property

Public Property Get InputSampleRate() As Long
    InputSampleRate = wfxIN.wfx.nSamplesPerSec
End Property

Public Property Let InputSampleRate(aval As Long)
    wfxIN.wfx.nSamplesPerSec = aval
End Property

Private Sub Class_Initialize()

    ' red book audio
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
