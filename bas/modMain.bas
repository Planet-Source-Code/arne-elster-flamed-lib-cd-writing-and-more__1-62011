Attribute VB_Name = "modMain"
Option Explicit

' stop the execution for a given time
Public Declare Sub Sleep Lib "kernel32" ( _
    ByVal dwMS As Long _
)

' "The IsBadReadPtr function verifies that
'  the calling process has read access to
'  the specified range of memory."
Public Declare Function IsBadReadPtr Lib "kernel32" ( _
    ByVal lp As Long, _
    ByVal ucb As Long _
) As Long

Private Declare Function GetTempPath Lib "kernel32" _
Alias "GetTempPathA" ( _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String _
) As Long

Public Function GetTempDir() As String

    Dim strBuffer As String

    strBuffer = Space(255)
    GetTempDir = Left$(strBuffer, GetTempPath(255, strBuffer))
    If Right$(GetTempDir, 1) <> "\" Then GetTempDir = GetTempDir & "\"

End Function
