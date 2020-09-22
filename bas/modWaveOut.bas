Attribute VB_Name = "modWaveOut"
Option Explicit

' waveOut functions

Public Declare Function waveOutOpen Lib "winmm.dll" ( _
                 lphWaveOut As Long, _
                 ByVal uDeviceID As Long, _
                 lpFormat As WAVEFORMAT, _
                 ByVal dwCallback As Long, _
                 ByVal dwInstance As Long, _
                 ByVal dwFlags As Long) As Long

Public Declare Function waveOutPrepareHeader Lib "winmm.dll" ( _
                 ByVal hWaveOut As Long, _
                 lpWaveOutHdr As WAVEHDR, _
                 ByVal uSize As Long) As Long

Public Declare Function waveOutWrite Lib "winmm.dll" ( _
                 ByVal hWaveOut As Long, _
                 lpWaveOutHdr As WAVEHDR, _
                 ByVal uSize As Long) As Long

Public Declare Function waveOutUnprepareHeader Lib "winmm.dll" ( _
                 ByVal hWaveOut As Long, _
                 lpWaveOutHdr As WAVEHDR, _
                 ByVal uSize As Long) As Long

Public Declare Function waveOutClose Lib "winmm.dll" ( _
                 ByVal hWaveOut As Long) As Long

Public Declare Function waveOutReset Lib "winmm.dll" ( _
                 ByVal hWaveOut As Long) As Long

Public Declare Function waveOutSetVolume Lib "winmm.dll" ( _
                 ByVal hWaveOut As Long, _
                 ByVal dwVolume As Long) As Long

Public Declare Function waveOutPause Lib "winmm.dll" ( _
                 ByVal hWaveOut As Long) As Long

Public Declare Function waveOutRestart Lib "winmm.dll" ( _
                 ByVal hWaveOut As Long) As Long

Public Declare Function SendMessageA Lib "user32" ( _
                 ByVal hWnd As Long, _
                 ByVal wMsg As Long, _
                 ByVal wParam As Long, _
                 ByVal lParam As Any) As Long

Public Declare Function CreateWindowEx Lib "user32" _
Alias "CreateWindowExA" ( _
    ByVal dwExStyle As Long, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hWndParent As Long, _
    ByVal hMenu As Long, _
    ByVal hInstance As Long, _
    ByVal lpParam As Long _
) As Long

Public Declare Function DestroyWindow Lib "user32" ( _
    ByVal hWnd As Long _
) As Long

Public Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long

Public Declare Function CallWindowProc Lib "user32" _
Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Public Const MM_WOM_DONE       As Long = &H3BD
Public Const WHDR_DONE         As Long = &H1

Public Type WAVEHDR
        lpData                  As Long
        dwBufferLength          As Long
        dwBytesRecorded         As Long
        dwUser                  As Long
        dwFlags                 As Long
        dwLoops                 As Long
        lpNext                  As Long
        reserved                As Long
End Type

Public Type WAVEFORMAT
        wFormatTag              As Integer
        nChannels               As Integer
        nSamplesPerSec          As Long
        nAvgBytesPerSec         As Long
        nBlockAlign             As Integer
        wBitsPerSample          As Integer
        cbSize                  As Integer
End Type

Public Type Stereo
    l                           As Integer
    R                           As Integer
End Type

Public Const GWL_WNDPROC       As Long = (-4)
