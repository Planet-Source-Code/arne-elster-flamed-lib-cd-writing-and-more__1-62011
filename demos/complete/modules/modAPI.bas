Attribute VB_Name = "modAPI"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" ( _
    ByVal ms As Long)

Public Declare Sub CopyMemory Lib "kernel32" Alias _
  "RtlMoveMemory" ( _
  pDst As Any, _
  pSrc As Any, _
  ByVal ByteLen As Long)

