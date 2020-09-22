Attribute VB_Name = "modFileHandling"
Option Explicit

' file handling functions

Public Const MOVEFILE_REPLACE_EXISTING = &H1
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_BEGIN = 0
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

Public Const FILE_FLAG_WRITE_THROUGH = &H80000000
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000

Public Declare Function ResetEvent Lib "kernel32" ( _
    ByVal hEvent As Long _
) As Long

Public Declare Function GetOverlappedResult Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpOverlapped As overlapped, _
    lpNumberOfBytesTransferred As Long, _
    ByVal bWait As Boolean _
) As Long

Public Declare Function CreateEvent Lib "kernel32" _
Alias "CreateEventA" ( _
    lpEventAttributes As Any, _
    ByVal bManualReset As Long, _
    ByVal bInitialState As Long, _
    ByVal lpName As String _
) As Long

Public Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long _
) As Long

Public Declare Function CreateFile Lib "kernel32" _
Alias "CreateFileA" ( _
        ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, _
        lpSecurityAttributes As Any, _
        ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long _
) As Long

Public Declare Function WriteFile Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpBuffer As Any, _
        ByVal nNumberOfBytesToWrite As Long, _
        lpNumberOfBytesWritten As Long, _
        lpOverlapped As Any _
) As Long

Public Declare Function SetFilePointer Lib "kernel32" ( _
        ByVal hFile As Long, _
        ByVallDistanceToMove As Long, _
        lpDistanceToMoveHigh As Long, _
        ByVal dwMoveMethod As Long _
) As Long

Public Declare Function ReadFile Lib "kernel32" ( _
        ByVal hFile As Long, _
        lpBuffer As Any, _
        ByVal nNumberOfBytesToRead As Long, _
        lpNumberOfBytesRead As Long, _
        lpOverlapped As Any _
) As Long

Public Type overlapped
    Internal                As Long
    InternalHigh            As Long
    offset                  As Long
    OffsetHigh              As Long
    hEvent                  As Long
End Type

Public Function FileExists(ByVal strFile As String) As Boolean
   FileExists = Dir(strFile) <> ""
End Function

Public Function PathFromPathFile(ByVal s As String) As String
    PathFromPathFile = Left$(s, InStrRev(s, "\"))
End Function

Public Function FileFromPathFile(ByVal s As String) As String
    FileFromPathFile = Mid$(s, InStrRev(s, "\") + 1)
End Function
