Attribute VB_Name = "modWinMM"
Option Explicit

Private Type mmioinfo
   dwFlags              As Long
   fccIOProc            As Long
   pIOProc              As Long
   wErrorRet            As Long
   htask                As Long
   cchBuffer            As Long
   pchBuffer            As String
   pchNext              As String
   pchEndRead           As String
   pchEndWrite          As String
   lBufOffset           As Long
   lDiskOffset          As Long
   adwInfo(4)           As Long
   dwReserved1          As Long
   dwReserved2          As Long
   hmmio                As Long
End Type

Private Type WAVEHDR
   lpData               As Long
   dwBufferLength       As Long
   dwBytesRecorded      As Long
   dwUser               As Long
   dwFlags              As Long
   dwLoops              As Long
   lpNext               As Long
   Reserved             As Long
End Type

Private Type WAVEINCAPS
   wMid                 As Integer
   wPid                 As Integer
   vDriverVersion       As Long
   szPname              As String * 32
   dwFormats            As Long
   wChannels            As Integer
End Type

Private Type MMCKINFO
   ckid                 As Long
   ckSize               As Long
   fccType              As Long
   dwDataOffset         As Long
   dwFlags              As Long
End Type

Public Type ChunkInfo
    start               As Long
    length              As Long
End Type

Private Const MMIO_READ         As Long = &H0
Private Const MMIO_FINDCHUNK    As Long = &H10
Private Const MMIO_FINDRIFF     As Long = &H20

Private Const SEEK_CUR          As Long = 1
Private Const SEEK_END          As Long = 2
Private Const SEEK_SET          As Long = 0

Private Declare Function mmioClose Lib "winmm.dll" ( _
        ByVal hmmio As Long, _
        ByVal uFlags As Long) As Long

Private Declare Function mmioDescend Lib "winmm.dll" ( _
        ByVal hmmio As Long, _
        lpck As MMCKINFO, _
        lpckParent As MMCKINFO, _
        ByVal uFlags As Long) As Long

Private Declare Function mmioDescendParent Lib "winmm.dll" _
Alias "mmioDescend" ( _
        ByVal hmmio As Long, _
        lpck As MMCKINFO, _
        ByVal X As Long, _
        ByVal uFlags As Long) As Long

Private Declare Function mmioOpen Lib "winmm.dll" _
Alias "mmioOpenA" ( _
        ByVal szFileName As String, _
        lpmmioinfo As mmioinfo, _
        ByVal dwOpenFlags As Long) As Long

Private Declare Function mmioRead Lib "winmm.dll" ( _
        ByVal hmmio As Long, _
        ByVal pch As Long, _
        ByVal cch As Long) As Long

Private Declare Function mmioReadString Lib "winmm.dll" _
Alias "mmioRead" ( _
        ByVal hmmio As Long, _
        ByVal pch As String, _
        ByVal cch As Long) As Long

Private Declare Function mmioSeek Lib "winmm.dll" ( _
        ByVal hmmio As Long, _
        ByVal lOffset As Long, _
        ByVal iOrigin As Long) As Long

Private Declare Function mmioStringToFOURCC Lib "winmm.dll" _
Alias "mmioStringToFOURCCA" ( _
        ByVal sz As String, _
        ByVal uFlags As Long) As Long

Private Declare Function mmioAscend Lib "winmm.dll" ( _
        ByVal hmmio As Long, _
        lpck As MMCKINFO, _
        ByVal uFlags As Long) As Long

Public Function GetWavFormat(ByVal strFile As String) As WAVEFORMATEX

    Dim FF  As Integer
    FF = FreeFile

    Dim lngStart    As Long
    Dim udtWFX      As WAVEFORMATEX

    With GetWavChunkPos(strFile, "fmt ")
        lngStart = .start
    End With

    Open strFile For Binary As #FF
        Get #FF, lngStart + 1, udtWFX
    Close #FF

    udtWFX.cbSize = 0

    GetWavFormat = udtWFX

End Function

Public Function GetWavChunkPos(ByVal strFile As String, ByVal strChunk As String) As ChunkInfo

    Dim hMmioIn             As Long
    Dim lR                  As Long
    Dim mmckinfoParentIn    As MMCKINFO
    Dim mmckinfoSubchunkIn  As MMCKINFO
    Dim mmioinf             As mmioinfo

    ' Open the input file
    hMmioIn = mmioOpen(strFile, mmioinf, MMIO_READ)
    If hMmioIn = 0 Then
        Exit Function
    End If

    ' Check if this is a wave file
    mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
    lR = mmioDescendParent(hMmioIn, mmckinfoParentIn, 0, MMIO_FINDRIFF)
    If Not (lR = MMSYSERR_NOERROR) Then
        mmioClose hMmioIn, 0
       Exit Function
    End If

    ' Find the data subchunk
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC(strChunk, 0)
    lR = mmioDescend(hMmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
    If Not (lR = MMSYSERR_NOERROR) Then
        mmioClose hMmioIn, 0
        Exit Function
    End If

    GetWavChunkPos.start = mmioSeek(hMmioIn, 0, SEEK_CUR)
    GetWavChunkPos.length = mmckinfoSubchunkIn.ckSize

    mmioClose hMmioIn, 0

End Function
