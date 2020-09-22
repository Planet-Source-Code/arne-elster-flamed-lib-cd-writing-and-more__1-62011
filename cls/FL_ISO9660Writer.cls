VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_ISO9660Writer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ToDo: strict Level 1 or Level 2
'       Joliet support

'*****************************************
'** clsISO9660Writer
'**
'** creates ISO9660 Level-1 Images.
'**
'** by [rm_code] 2005
'
'                    http://rm_code.dl.am/
'*****************************************


Private Declare Function FindFirstFile Lib "kernel32" _
Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA _
) As Long

Private Declare Function FindNextFile Lib "kernel32" _
Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA _
) As Long

Private Declare Function GetFileAttributes Lib "kernel32" _
Alias "GetFileAttributesA" ( _
    ByVal lpFileName As String _
) As Long

Private Declare Function FindClose Lib "kernel32" ( _
    ByVal hFindFile As Long _
) As Long

Private files()                         As String
Private filecnt                         As Long

Private dirs()                          As String
Private dircnt                          As Long

Private strVolID                        As String
Private strSysID                        As String
Private strAppID                        As String
Private strPubID                        As String

Private strStartDir                     As String
Private strISOStartDir                  As String

Private Const MAX_PATH                  As Long = 260
Private Const MAXDWORD                  As Long = &HFFFF
Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const FILE_ATTRIBUTE_ARCHIVE    As Long = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10
Private Const FILE_ATTRIBUTE_HIDDEN     As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
Private Const FILE_ATTRIBUTE_READONLY   As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM     As Long = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY  As Long = &H100

Private Type DirRecord
    Length As Byte                              ' length of the record
    ExtAttrRecLen As Byte                       ' ?
    LocExt As String * 8                        ' sector with associated data
    DataLen As String * 8                       ' data length (in Bytes)
    RecDateTime As String * 7                   ' date of the record
    FileFlags As Byte                           ' Flags (file or directory)
    FileUnitSize As Byte                        ' ?
    InterleaveGapSize As Byte                   ' ?
    VolSeqNum As String * 4                     ' ?
    FileIDLen As Byte                           ' length of the name
End Type

Private Type RootDirRecord
    Length As Byte                              ' length of the record
    ExtAttrRecLen As Byte                       ' ?
    LocExt As String * 8                        ' sector with associated data
    DataLen As String * 8                       ' data length (in Bytes)
    RecDateTime As String * 7                   ' date of the record
    FileFlags As Byte                           ' Flags (file or directory)
    FileUnitSize As Byte                        ' ?
    InterleaveGapSize As Byte                   ' ?
    VolSeqNum As String * 4                     ' ?
    FileIDLen As Byte                           ' length of the name
    FileID(0) As Byte                           ' filename
End Type

Private Type ISO_VOL_DESCRIPTOR
    Vol_Desc_Type As Byte                       'Volume Descriptor (1)
    Std_ID        As String * 6                 'CD001
    Vol_Desc_Ver  As Byte                       'Volume Descriptor Version (0)
End Type

Private Type ISO_PRIMARY_DESCRIPTOR

    System_ID As String * 32                    ' System ID
    Volume_ID As String * 32                    ' Volume ID

    unused2(1 To 8) As Byte                     ' reserved

    Vol_Space_Size As String * 8                ' Volume size
    unused3(1 To 32) As Byte                    ' reserved
    Vol_Set_Size As String * 4                  ' Volume Set Size (1)
    Vol_Seq_Number As String * 4                ' Volume Sequence Number (1)
    Logical_Block_Size As String * 4            ' Logische Sectorsize (2048)

    Path_Table_Size As String * 8               ' path table size
    Location_TypeL_Path_Table As String * 4     ' location of the little endian path table
    Location_TypeL_Opt_Path_Table As String * 4 ' location of the optional little endian path table
    Location_TypeM_Path_Table As String * 4     ' location of the big endian path table
    Location_TypeM_Opt_Oath_Table As String * 4 ' location of the optional big endian path table
    Dir_Root_Record As RootDirRecord            ' Root Directory Record ("\")

    Vol_Set_ID As String * 128                  ' Volume Set Identifier
    Publisher_ID As String * 128                ' Publisher Identifier
    Data_Preparer_ID As String * 128            ' Data Preparer Identifier
    App_ID As String * 128                      ' Application Identifier

    Copyright_File_ID As String * 37            ' Copyright File Identifier
    Abstract_File_ID As String * 37             ' Abstract File Identifier
    Bibliographic_File_ID As String * 37        ' Bibliographic File Idenfitier

    Vol_Creation_Date As String * 17            ' creation date
    Vol_Modification_Date As String * 17        ' modification date
    Vol_Expiration_Date As String * 17          ' expiration date
    Vol_Effective_Date As String * 17           ' effective date

    File_Structure_Ver As Byte                  ' file structure version
    rsvd As Byte                                ' reserved
    app_use(1 To 512) As Byte                   ' reserved
    rsvd2(1 To 653) As Byte                     ' reserved
End Type

Private Type PTableRecord
    Length As Byte                              ' length of the record
    ExtAttrRecLen As Byte                       ' ?
    FirstSector As String * 4                   ' first sector of the dir
    RecNumParentDir As String * 2               ' record number of the parent dir
End Type

Private Type FILETIME
    dwLowDateTime               As Long
    dwHighDateTime              As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes            As Long
    ftCreationTime              As FILETIME
    ftLastAccessTime            As FILETIME
    ftLastWriteTime             As FILETIME
    nFileSizeHigh               As Long
    nFileSizeLow                As Long
    dwReserved0                 As Long
    dwReserved1                 As Long
    cFileName                   As String * MAX_PATH
    cAlternate                  As String * 14
End Type

Private Enum eFileFlags
    FLAG_HIDDEN = &H1                           ' hidden
    FLAG_DIRECTORY = &H2                        ' directory
    FLAG_ASSOCIATED = &H4                       ' associated
    FLAG_RECFMTSPEC = &H8                       ' format specific
    FLAG_PERM = &H16                            ' permission needed
    FLAG_NOTFINAL = &H80                        ' ?
End Enum

Private Enum eVolDesc
    VOL_DESC_BOOT_RECORD                        ' Boot Record
    VOL_DESC_PRIM_VOL_DESCRIPTOR                ' Primary Volume Descriptor
    VOL_DESC_SUPPL_VOL_DESCRIPTOR               ' Supplementary Volume Descriptor
    VOL_DESC_VOL_PART_DESCRIPTOR                ' Volume Partition Descriptor
    VOL_DESC_SET_TERMINATOR                     ' Volume Descriptor Set Terminator
End Enum

Public Enum eFileDetail
    FD_Filename = 0
    FD_ISOPath
    FD_FullISOPath
    FD_LocalPath
    FD_FileSize
    FD_StartLBA
End Enum

Private Const a_chars As String = "!""%&'()*+,-./0123456789:;<>=?ABCDEFGHIJKLMNOPQRSTUVWXYZ_ "
Private Const d_chars As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_.-"

Public Event Progress(ByVal lngMax As Long, ByVal lngValue As Long)
Public Event ISOStart()

'Publisher
Public Property Get PublisherID() As String

    PublisherID = strPubID

End Property

'Publisher
Public Property Let PublisherID(ByVal str As String)

    strPubID = Left$(To_A_Chars(str), 128)

End Property

'Application
Public Property Get AppID() As String

    AppID = strAppID

End Property

'Application
Public Property Let AppID(ByVal str As String)

    strAppID = Left$(To_A_Chars(str), 128)

End Property

'System
Public Property Get SystemID() As String

    SystemID = strSysID

End Property

'System
Public Property Let SystemID(ByVal str As String)

    strSysID = Left$(To_D_Chars(str), 32)

End Property

'Title
Public Property Get VolumeID() As String

    VolumeID = strVolID

End Property

'Title
Public Property Let VolumeID(ByVal str As String)

    strVolID = Left$(To_D_Chars(str), 32)

End Property

Private Sub RaisePrg(ByVal max As Long, ByVal fh As Integer)

    If max < Seek(fh) Then max = Seek(fh)

    RaiseEvent Progress(max, Seek(fh))

End Sub

Private Sub DelEx(file As String)
    On Error Resume Next
    Kill file
End Sub

'create image
Public Function CreateISO(ByVal strPath As String, _
                Optional blnOnlyHeader As Boolean) As Boolean

    '   Volume Descriptor              Primary Volume Descriptor
    Dim vd      As ISO_VOL_DESCRIPTOR, pvd      As ISO_PRIMARY_DESCRIPTOR
    '   Filehandle
    Dim fh      As Integer:            fh = FreeFile
    '   Start LBA
    Dim startLBA    As Long:           Dim i    As Long
    '   data reset
    Dim strParts    As String
    '   ISO size
    Dim lngISOSize  As Long


    'delete the image, if exists
    If FileExists(strPath) Then DelEx strPath

    RaiseEvent ISOStart

    'open the file
    Open strPath For Binary As #fh


        'ISO size
        lngISOSize = ISOSize


        '32 KB emptyness
        Put #fh, , CStr(String(32768, Chr$(0)))


        'primary volume Descriptor
        With vd

            .Vol_Desc_Type = VOL_DESC_PRIM_VOL_DESCRIPTOR
            .Std_ID = "CD001" & Chr$(1)

        End With

        'Primary Volume Descriptor
        With pvd

            .System_ID = strSysID
            .Volume_ID = strVolID

            'sectorsize always 2048
            .Logical_Block_Size = To16BitBothByte(2048)

            'location of the little endian path table
            .Location_TypeL_Path_Table = To32BitMS(17 + GetCDFSSize(GetPathTableSize) \ 2048)
            'location of the big endian path table
            .Location_TypeM_Path_Table = To32BitLS(17)
            'pathtablesize in bytes
            .Path_Table_Size = To32BitBothByte(GetPathTableSize)

            'filesystem (ISO) size
            .Vol_Space_Size = To32BitBothByte(lngISOSize \ 2048)

            '?
            .Vol_Seq_Number = To16BitBothByte(1)
            .Vol_Set_Size = To16BitBothByte(1)

            'About
            .Data_Preparer_ID = Space(128)
            .Publisher_ID = strPubID
            .Bibliographic_File_ID = Space(37)
            .Abstract_File_ID = Space(37)
            .App_ID = strPubID
            .Vol_Set_ID = Space(128)

            .Vol_Modification_Date = CreateDateTime
            .Vol_Expiration_Date = CreateDateTime
            .Vol_Effective_Date = CreateDateTime
            .Vol_Creation_Date = CreateDateTime

            'Root Directory Record
            With .Dir_Root_Record

                'location of the Root Directory Record content
                .LocExt = To32BitBothByte(17 + (GetCDFSSize(GetPathTableSize) * 2) \ 2048)
                'size of it
                .DataLen = To32BitBothByte(GetCDFSSize(GetDirSize("\")))
                'creation date
                .RecDateTime = CreateDateTimeRec
                '?
                .VolSeqNum = To16BitBothByte(1)
                'directory
                .FileFlags = FLAG_DIRECTORY
                'name of the root is 1 byte big
                .FileIDLen = 1
                'Directory Record length
                .Length = 34

            End With

        End With

        Put #fh, , vd
        Put #fh, , pvd

        'fill the sector with zeros
        FillSector fh

        RaisePrg lngISOSize, fh

        'startsectors of the directory contents
        CreateDirStartOffsets
        'startsectors of the files
        CreateFileStartOffsets

        'write the path table (LS and MS)
        WritePathTable fh, 0
        WritePathTable fh, 1

        RaisePrg lngISOSize, fh

        'write a directory record for all directories
        For i = 0 To dircnt - 1

            ' define the start LBA
            startLBA = LOF(fh) / 2048

            ' write it
            WriteDirectoryRecord fh, FullDirFromIndex(i), startLBA

            RaisePrg lngISOSize, fh

        Next

        If Not blnOnlyHeader Then

            ' write all files to the image
            For i = 0 To filecnt - 1

                If Not WriteFile(fh, GetFileDetailByIndex(i, FD_LocalPath) & _
                                 GetFileDetailByIndex(i, FD_Filename)) Then

                    'Debug.Print "Konnte Datei mit Index " & i & " nicht lesen"

                End If

                RaisePrg lngISOSize, fh
    
            Next

        End If

    Close #fh


    'reset the data
    For i = 0 To filecnt - 1

        files(i) = Join(Split(files(i), Chr$(0), 3), Chr$(0))

    Next

    For i = 0 To dircnt - 1

        dirs(i) = Join(Split(dirs(i), Chr$(0), 2), Chr$(0))

    Next


    RaiseEvent Progress(1, 1)


    'sucess
    CreateISO = True

End Function

'write a file
Private Function WriteFile(ByVal fh As Integer, _
                           ByVal file As String) As Boolean

    On Error GoTo ErrHandler

    Dim strBuf  As String

    Dim FF      As Integer: FF = FreeFile



    'read the file content
    Open file For Binary As #FF

        strBuf = Space(LOF(FF))

        'when the file is empty, we need to write at least 1 byte
        If strBuf = "" Then strBuf = " "

        Get #FF, , strBuf

    Close #FF

    'and write it to the new file
    Put #fh, , strBuf

    'fill the last sector with zeros
    FillSector fh


    WriteFile = True


ErrHandler:
End Function

'write a directory record
Private Function WriteDirectoryRecord(ByVal fh As Integer, _
                                      ByVal strISOPath As String, _
                                      ByVal startLBA As Long) As Long

    Dim SubFiles()  As String
    Dim SubDirs()   As String
    Dim DirRec      As DirRecord
    Dim i           As Long

    'read the directory content
    SubFiles = GetFiles(strISOPath)
    SubDirs = GetDirs(strISOPath)

    '".\"
    With DirRec
        .DataLen = To32BitBothByte(GetCDFSSize(GetDirSize(strISOPath)))
        .FileFlags = FLAG_DIRECTORY Or FLAG_HIDDEN
        .LocExt = To32BitBothByte(startLBA)
        .RecDateTime = CreateDateTimeRec
        .VolSeqNum = To16BitBothByte(1)
        .FileIDLen = 1
        .Length = 34
    End With

    Put #fh, , DirRec: Put #fh, , CByte(0)

    '"..\"
    With DirRec
        .LocExt = To32BitBothByte(Split(dirs(IndexFromFullDir(GetParentDir(strISOPath))), Chr$(0))(2))
        .DataLen = To32BitBothByte(GetCDFSSize(GetDirSize(GetParentDir(strISOPath))))
    End With

    Put #fh, , DirRec: Put #fh, , CByte(1)

    'if there are directories...
    If DirCount(strISOPath) > 0 Then

        For i = 0 To UBound(SubDirs)

            '... add them
            With DirRec

                .Length = 34 + Len(SubDirs(i))
                .FileIDLen = Len(SubDirs(i))

                .LocExt = To32BitBothByte(Split(dirs(IndexFromFullDir(strISOPath & SubDirs(i))), Chr$(0))(2))
                .DataLen = To32BitBothByte(GetCDFSSize(GetDirSize(strISOPath & SubDirs(i))))

                .RecDateTime = CreateDateTimeRec
                .VolSeqNum = To16BitBothByte(1)

                .FileFlags = FLAG_DIRECTORY

            End With

            'check if we are near to the sector bound
            If CheckOffset(fh, DirRec.Length) Then FillSector fh

            Put #fh, , DirRec
            Put #fh, , CStr(SubDirs(i))
            Put #fh, , CByte(0)

        Next

    End If


    'if there are files...
    If FileCount(strISOPath) > 0 Then

        For i = 0 To UBound(SubFiles)

            '... add them
            With DirRec

                .Length = 34 + Len(SubFiles(i)) + 2

                .FileIDLen = Len(SubFiles(i)) + 2

                .DataLen = To32BitBothByte(GetFileDetailByPath(strISOPath & SubFiles(i), FD_FileSize))
                .LocExt = To32BitBothByte(GetFileDetailByPath(strISOPath & SubFiles(i), FD_StartLBA))

                .RecDateTime = CreateDateTimeRec
                .VolSeqNum = To16BitBothByte(1)

                .FileFlags = 0

            End With

            'check if we are near to a sector bound
            If CheckOffset(fh, DirRec.Length) Then FillSector fh

            Put #fh, , DirRec
            Put #fh, , CStr(SubFiles(i) & ";1")
            Put #fh, , CByte(0)

        Next

    End If

    FillSector fh

End Function

'near to the sector bound?
Private Function CheckOffset(ByVal fh As Integer, _
                             ByVal BytesToAdd As Long) As Boolean

    CheckOffset = LOF(fh) + BytesToAdd > GetCDFSSize(LOF(fh))

End Function

'calculate the directory start LBAs
Private Sub CreateDirStartOffsets()

    Dim i       As Long, j      As Long


    'startsector for root
    j = 17 + (GetCDFSSize(GetPathTableSize) * 2 / 2048)

    For i = 0 To dircnt - 1

        'add startpos to dir info
        dirs(i) = dirs(i) & Chr$(0) & _
                  CStr(j) & Chr$(0) & _
                  CStr(i + 1)

        'next position
        j = j + GetCDFSSize(GetDirSize(FullDirFromIndex(i))) / 2048

    Next

End Sub

'calculate file start LBAs
Private Sub CreateFileStartOffsets()

    Dim i       As Long, j      As Long
    Dim k       As Long

    'end of the last directory record
    j = Split(dirs(UBound(dirs)), Chr$(0))(2) + GetCDFSSize(GetDirSize(FullDirFromIndex(UBound(dirs)))) / 2048

    For i = 0 To filecnt - 1

        'append the start LBA to the file entry
        files(i) = files(i) & Chr$(0) & CStr(j)

        'next position
        k = GetCDFSSize(GetFileDetailByIndex(i, FD_FileSize)) / 2048

        j = j + k

    Next

End Sub

'write the path table
Private Sub WritePathTable(ByVal fh As Integer, ByVal order As Integer)

    Dim PTable  As PTableRecord
    Dim i       As Long
    Dim func    As String

    'create the Root Entry
    With PTable

        'LSB
        If order = 0 Then

            .FirstSector = To32BitLS(Split(dirs(0), Chr$(0))(2))
            .RecNumParentDir = To16BitLS(1)

        'MSB
        Else

            .FirstSector = To32BitMS(Split(dirs(0), Chr$(0))(2))
            .RecNumParentDir = To16BitMS(1)

        End If

        .Length = 1

    End With

    Put #fh, , PTable
    Put #fh, , CByte(0)
    Put #fh, , CByte(0)

    For i = 1 To dircnt - 1

        'LSB
        If order = 0 Then

            'Recordnumber of parentdirectory
            PTable.RecNumParentDir = To16BitLS(Split(dirs(IndexFromFullDir(GetParentDir(FullDirFromIndex(i)))), Chr$(0))(3))
            'startsector of the Directory Record
            PTable.FirstSector = To32BitLS(Split(dirs(i), Chr$(0))(2))

        'MSB
        Else

            'Recordnumber of parentdirectory
            PTable.RecNumParentDir = To16BitMS(Split(dirs(IndexFromFullDir(GetParentDir(FullDirFromIndex(i)))), Chr$(0))(3))
            'startsector of the Directory Record
            PTable.FirstSector = To32BitMS(Split(dirs(i), Chr$(0))(2))

        End If

        'length of the dir name
        PTable.Length = Len(GetLastItem(RemSlash(FullDirFromIndex(i))))

        Put #fh, , PTable

        Put #fh, , CStr(GetLastItem(RemSlash(FullDirFromIndex(i))))

        'append a null char to odd namelengths
        If PTable.Length Mod 2 Then Put #fh, , CByte(0)

    Next

    'fill the sector with zeros
    FillSector fh

End Sub

'fill the current sector with null chars
Private Function FillSector(ByVal fh As Integer) As Integer

    If GetCDFSSize(LOF(fh)) > LOF(fh) Then

        Put #fh, , CStr(String(GetCDFSSize(LOF(fh)) - LOF(fh), CByte(0)))

    End If

End Function

'calculate a directory record size
Private Function GetDirSize(ByVal strISOPath As String) As Long

    Const DirSize   As Byte = 34

    Dim strFiles()  As String
    Dim strDirs()   As String

    Dim lngSize     As Long
    Dim bytCurr     As Byte
    Dim i           As Long
    Dim j           As Long

    'dir content
    strFiles = GetFiles(strISOPath)
    strDirs = GetDirs(strISOPath)

    '"." and ".."
    lngSize = lngSize + DirSize * 2

    If DirCount(strISOPath) > 0 Then

        'subdirs
        For j = 0 To UBound(strDirs)

            If lngSize + (DirSize + Len(GetLastItem(FullDirFromIndex(j)))) > GetCDFSSize(lngSize) Then _
                lngSize = GetCDFSSize(lngSize)

            lngSize = lngSize + (DirSize + Len(GetLastItem(FullDirFromIndex(j))))

        Next

    End If


    If FileCount(strISOPath) > 0 Then

        'files
        For j = 0 To UBound(strFiles)

            If lngSize + (DirSize + Len(GetFileDetailByIndex(j, FD_Filename))) > GetCDFSSize(lngSize) Then _
                lngSize = GetCDFSSize(lngSize)

            lngSize = lngSize + (DirSize + Len(GetFileDetailByIndex(j, FD_Filename)))

            If Len(GetFileDetailByIndex(j, FD_Filename)) Mod 2 Then _
                lngSize = lngSize + 1

        Next

    End If

    GetDirSize = lngSize

End Function

'aprox. ISO size
Public Property Get ISOSize() As Long

    Dim i   As Integer


    ISOSize = CLng(16) * 2048 + GetCDFSSize(GetPathTableSize) * 2

    For i = 0 To filecnt - 1

        ISOSize = ISOSize + CLng(GetFileDetailByIndex(i, FD_FileSize))

    Next

End Property

'calculate filetable size
Private Function GetFileTableSize() As Long

    On Error Resume Next

    Const DirSize   As Byte = 34

    Dim strFiles()  As String
    Dim strDirs()   As String

    Dim lngSize     As Long
    Dim bytCurr     As Byte
    Dim i           As Long
    Dim j           As Long

    For i = 0 To dircnt - 1

        'dir content
        strFiles = GetFiles(FullDirFromIndex(i))
        strDirs = GetDirs(FullDirFromIndex(i))

        '"." and ".."
        lngSize = lngSize + DirSize * 2

        'subdirs
        For j = 0 To UBound(strDirs)

            If lngSize + (DirSize + Len(GetLastItem(FullDirFromIndex(j)))) > GetCDFSSize(lngSize) Then _
                lngSize = GetCDFSSize(lngSize)

            lngSize = lngSize + (DirSize + Len(GetLastItem(FullDirFromIndex(j))))

        Next

        'files
        For j = 0 To UBound(strFiles)

            If lngSize + (DirSize + Len(GetFileDetailByIndex(j, FD_Filename))) > GetCDFSSize(lngSize) Then _
                lngSize = GetCDFSSize(lngSize)

            lngSize = lngSize + (DirSize + Len(GetFileDetailByIndex(j, FD_Filename)))

        Next

    Next

    GetFileTableSize = lngSize

End Function

Private Function GetPathTableSize() As Long

    Dim lngSize     As Long
    Dim bytCurr     As Byte
    Dim i           As Long

    'Root
    lngSize = 10

    For i = 1 To dircnt - 1

        'a path table entry is 8 bytes + name length of the directory big
        bytCurr = Len(GetLastItem(RemSlash(FullDirFromIndex(i))))
        If bytCurr Mod 2 Then bytCurr = bytCurr + 1
        bytCurr = bytCurr + 8

        'check if we are near to a sector bound
        If bytCurr + lngSize > GetCDFSSize(lngSize) Then _
            lngSize = GetCDFSSize(lngSize)

        lngSize = lngSize + bytCurr

    Next

    GetPathTableSize = lngSize

End Function

'Private Function RoundUp(ByVal val As Double) As Long
'
'    If Fix(val) < val Then RoundUp = CLng(val) + 1
'
'End Function

'Private Function GetCDFSSize(ByVal length As Long) As Long
'
'    GetCDFSSize = RoundUp(length / 2048) * 2048
'    If GetCDFSSize = 0 Then GetCDFSSize = 2048
'
'End Function

'round bytes up to sectorsize
Private Function GetCDFSSize(ByVal Length As Long) As Long

    Dim i As Byte, j As Long

    While Not i = &HFF

        j = j + 1

        GetCDFSSize = j * 2048

        If GetCDFSSize >= Length Then i = &HFF

    Wend

End Function

'return subdirs
Public Function GetDirs(ByVal strISOPath As String) As String()

    '   counters
    Dim i           As Long, j       As Long
    '   Array buffer           path buffer
    Dim strBuf()    As String, strPath  As String

    strISOPath = AddSlash(strISOPath)

    For i = 0 To dircnt - 1

        strPath = FullDirFromIndex(i)

        'don't allow root or same dir
        If Not (strPath = "\" Or StrComp(strPath, strISOPath, vbTextCompare) = 0) Then

            'found dir?
            If StrComp(GetParentDir(strPath), strISOPath, vbTextCompare) = 0 Then

                'add it
                ReDim Preserve strBuf(j) As String
                strBuf(j) = GetLastItem(RemSlash(strPath))
                j = j + 1

            End If

        End If

    Next

    GetDirs = strBuf

End Function

'return files of a dir
Public Function GetFiles(ByVal strISOPath As String) As String()

    '   counters
    Dim i           As Long, j       As Long
    '   Array buffer
    Dim strBuf()    As String

    strISOPath = AddSlash(strISOPath)

    For i = 0 To filecnt - 1

        'file is in the same dir?
        If StrComp(GetFileDetailByIndex(i, FD_ISOPath), _
                   strISOPath) = 0 Then

            ReDim Preserve strBuf(j) As String
            strBuf(j) = GetFileDetailByIndex(i, FD_Filename)
            j = j + 1

        End If

    Next

    GetFiles = strBuf

End Function

'number of files in a dir
Public Function FileCount(ByVal strISOPath As String) As Long

    Dim i           As Long


    strISOPath = AddSlash(strISOPath)

    For i = 0 To filecnt - 1

        'file is in the same dir?
        If StrComp(GetFileDetailByIndex(i, FD_ISOPath), _
                   strISOPath) = 0 Then

            FileCount = FileCount + 1

        End If

    Next

End Function

'number of sub dirs
Public Function DirCount(ByVal strISOPath As String) As Long

    Dim i           As Long
    Dim strPath     As String

    strISOPath = AddSlash(strISOPath)

    For i = 0 To dircnt - 1

        strPath = FullDirFromIndex(i)

        'root or same dir not allowed
        If Not (strPath = "\" Or StrComp(strPath, strISOPath, vbTextCompare) = 0) Then

'---------------
'            'strISOPath kann Parent zu aktuellem Verzeichnis sein?
'            If Left$(strPath, Len(strISOPath)) = strISOPath Then
'
'                'Pfadebenen vergleichen
'                If PathLayers(strISOPath) = PathLayers(strPath) - 1 Then
'
'                    DirCount = DirCount + 1
'
'                End If
'
'            End If
'---------------

            If StrComp(GetParentDir(strPath), strISOPath, vbTextCompare) = 0 Then
                DirCount = DirCount + 1
            End If

        End If

    Next

End Function

'remove a dir
Public Function RemDir(ByVal strISOPath As String) As Boolean

    '   counters
    Dim i           As Long, j   As Long
    '   buffer
    Dim strBuf()    As String
    '   ISO path
    Dim strPath     As String

    strISOPath = AddSlash(strISOPath)

    'root cannot be deleted
    If strISOPath = "\" Then Exit Function

    If Not ISODirExists(strISOPath) Then Exit Function

    For i = 0 To dircnt - 1

        strPath = FullDirFromIndex(i)

        'ignore dirs to delete
        If Not Left$(LCase$(strPath), Len(strISOPath)) = LCase$(strISOPath) Then

            ReDim Preserve strBuf(j) As String
            strBuf(j) = dirs(i)
            j = j + 1

        Else

            dircnt = dircnt - 1

        End If

    Next

    dirs = strBuf


    ReDim strBuf(0)
    j = 0


    For i = 0 To filecnt - 1

        strPath = GetFileDetailByIndex(i, FD_ISOPath)

        If Not Left$(LCase$(strPath), Len(strISOPath)) = LCase$(strISOPath) Then

            ReDim Preserve strBuf(j) As String
            strBuf(j) = files(i)
            j = j + 1

        Else

            filecnt = filecnt - 1

        End If

    Next

    files = strBuf

    RemDir = True

End Function

'create a new dir
Public Function NewDir(ByVal strISOPath As String, _
                       ByVal strNewDir As String) As Boolean

    strISOPath = AddSlash(strISOPath)
    strNewDir = AddSlash(strNewDir)

    'Parent exists?
    If Not ISODirExists(strISOPath) Then Exit Function

    'dir exists already?
    If ISODirExists(strISOPath & strNewDir) Then Exit Function

    ReDim Preserve dirs(dircnt) As String

    'insert a new dir
    dirs(dircnt) = RemSlash(strNewDir) & Chr$(0) & _
                   strISOPath

    dircnt = dircnt + 1

    NewDir = True

End Function

'remove a file
Public Function RemFile(ByVal strISOPath As String) As Boolean

    '   counters
    Dim i           As Long:  Dim j     As Long
    '   Array buffer
    Dim strBuf()    As String

    'file exists?
    If Not ISOFileExists(strISOPath) Then Exit Function

    'only 1 file in the project?
    If filecnt = 1 Then

        files(0) = ""
        filecnt = 0

        RemFile = True

        Exit Function

    End If
    
    ReDim strBuf(filecnt - 2) As String

    For i = 0 To filecnt - 1

        'ignore the file to delete
        If Not StrComp(GetFileDetailByIndex(i, FD_FullISOPath), _
                       strISOPath) = 0 Then

            strBuf(j) = files(i)
            j = j + 1

        End If

    Next

    files = strBuf
    filecnt = filecnt - 1

    RemFile = True

End Function

Public Function AddDir(ByVal strLocalDir As String, _
                       ByVal strISOPath As String) As Boolean

    strLocalDir = AddSlash(strLocalDir)
    strISOPath = AddSlash(strISOPath)

    'dir exists?
    If Not DirExists(strLocalDir) Then Exit Function

    NewDir strISOPath, GetLastItem(RemSlash(strLocalDir))

    strISOStartDir = strISOPath & GetLastItem(RemSlash(strLocalDir))
    strStartDir = strLocalDir

    'search for dirs and files
    FindFilesAPI strStartDir, "*.*", 0, 0

    AddDir = True

End Function

'add a file
Public Function AddFile(ByVal strLocalFile As String, _
                        ByVal strISOPath As String) As Boolean

    strISOPath = AddSlash(strISOPath)


    'file exists?
    If Not FileExists(strLocalFile) Then _
        Exit Function

    'ISO dir exists?
    If Not ISODirExists(strISOPath) Then _
        Exit Function

    'ISO file exists already?
    If ISOFileExists(strISOPath & GetLastItem(strLocalFile)) Then _
        Exit Function



    ReDim Preserve files(filecnt) As String

    'add the file
    files(filecnt) = GetLastItem(strLocalFile) & Chr$(0) & _
                     strISOPath & Chr$(0) & _
                     GetAllButLastItem(strLocalFile) & Chr$(0) & _
                     FileLen(strLocalFile)

    filecnt = filecnt + 1

    AddFile = True

End Function

'return file information
Public Function GetFileDetailByPath(ByVal ISOPath As String, _
                                    ByVal detail As eFileDetail) As String

    Dim strParts()  As String

    strParts = Split(files(FileIndexByPath(ISOPath)), Chr$(0))

    Select Case detail
        Case FD_FullISOPath: GetFileDetailByPath = strParts(1) & strParts(0)
        Case FD_LocalPath:   GetFileDetailByPath = strParts(2)
        Case FD_Filename:    GetFileDetailByPath = strParts(0)
        Case FD_FileSize:    GetFileDetailByPath = strParts(3)
        Case FD_StartLBA:    GetFileDetailByPath = strParts(4)
        Case FD_ISOPath:     GetFileDetailByPath = strParts(1)
    End Select

End Function

'return all files in the image
Public Function GetLocalFiles() As String()

    Dim ret()   As String
    Dim i       As Long

    ReDim ret(UBound(files) + 1) As String

    For i = 0 To UBound(ret) - 1
        ret(i + 1) = AddSlash(GetFileDetailByIndex(i, FD_LocalPath)) & _
                              GetFileDetailByIndex(i, FD_Filename)
    Next

    GetLocalFiles = ret

End Function

'return file information
Private Function GetFileDetailByIndex(ByVal index As Long, _
                                      ByVal detail As eFileDetail) As String

    Dim strParts()  As String

    strParts = Split(files(index), Chr$(0))

    Select Case detail
        Case FD_FullISOPath: GetFileDetailByIndex = strParts(1) & strParts(0)
        Case FD_LocalPath:   GetFileDetailByIndex = strParts(2)
        Case FD_Filename:    GetFileDetailByIndex = strParts(0)
        Case FD_FileSize:    GetFileDetailByIndex = strParts(3)
        Case FD_StartLBA:    GetFileDetailByIndex = strParts(4)
        Case FD_ISOPath:     GetFileDetailByIndex = strParts(1)
    End Select

End Function

'return the index of a file
Private Function FileIndexByPath(ByVal strISOPath As String) As Long

    Dim i   As Long

    For i = 0 To UBound(files)

        If StrComp(GetFileDetailByIndex(i, FD_FullISOPath), _
                   strISOPath, vbTextCompare) = 0 Then

            FileIndexByPath = i
            Exit Function

        End If

    Next
End Function

Private Function GetAllButLastItem(ByVal str As String, _
                                   Optional ByVal Delimeter As String = "\" _
                                  ) As String

    Dim strParts()  As String
    Dim i           As Long

    strParts = Split(str, Delimeter)

    For i = 0 To UBound(strParts) - 1
        GetAllButLastItem = GetAllButLastItem & strParts(i) & Delimeter
    Next

End Function

Private Function GetLastItem(ByVal str As String, _
                             Optional ByVal Delimeter As String = "\" _
                            ) As String

    Dim strParts() As String

    strParts = Split(str, Delimeter)
    GetLastItem = strParts(UBound(strParts))

End Function

'file exists already in the project?
Public Function ISOFileExists(ByVal strPathFile As String) As Boolean

    Dim i   As Long

    For i = 0 To filecnt - 1

        If StrComp(GetFileDetailByIndex(i, FD_FullISOPath), _
                   strPathFile) = 0 Then

            ISOFileExists = True
            Exit Function

        End If

    Next

End Function

'dir exists already in the project?
Public Function ISODirExists(ByVal strPath As String) As Boolean

    Dim i   As Long

    For i = 0 To dircnt - 1

        If StrComp(FullDirFromIndex(i), strPath, vbTextCompare) = 0 Then
            ISODirExists = True
            Exit Function
        End If

    Next

End Function

Private Function GetParentDir(ByVal str As String) As String

    str = RemSlash(str)

    If str = "" Then GetParentDir = "\": Exit Function

    GetParentDir = AddSlash(Left$(str, Len(str) - Len(GetLastItem(str))))

End Function

Private Function IndexFromFullDir(ByVal strISOPath As String) As Long

    Dim i   As Long

    For i = 0 To dircnt - 1

        If StrComp(RemSlash(FullDirFromIndex(i)), RemSlash(strISOPath), vbTextCompare) = 0 Then

            IndexFromFullDir = i
            Exit Function

        End If

    Next

End Function

Private Function FullDirFromIndex(ByVal index As Long) As String

    FullDirFromIndex = AddSlash(Split(dirs(index), Chr$(0))(1) & _
                                Split(dirs(index), Chr$(0))(0))

End Function

'Pfadebenen
'Private Function PathLayers(ByVal strPath As String, _
'                            Optional ByVal Delimeter As String = "\" _
'                           ) As Long
'
'    'Pfadebenen
'    PathLayers = UBound(Split(strPath, Delimeter))
'
'End Function

Private Function RemSlash(ByVal str As String) As String

    RemSlash = IIf(Right$(str, 1) = "\", Left$(str, Len(str) - 1), str)

End Function

Private Function AddSlash(ByVal str As String) As String

    AddSlash = IIf(Right$(str, 1) = "\", str, str & "\")

End Function

Private Function DirExists(ByVal strPath As String) As Boolean

    DirExists = CBool(GetAttr(strPath) And vbDirectory)

End Function

Private Function StripNulls(OriginalStr As String) As String

    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If

    StripNulls = OriginalStr

End Function

Private Function FindFilesAPI(Path As String, _
                              SearchStr As String, _
                              FileCount As Integer, _
                              DirCount As Integer)

    'KPD-Team 1999
    'E-Mail: KPDTeam@Allapi.net
    'URL: http://www.allapi.net/

    Dim FileName    As String
    Dim DirName     As String
    Dim dirNames()  As String

    Dim nDir        As Integer
    Dim i           As Integer
    Dim Cont        As Integer

    Dim hSearch     As Long

    Dim WFD         As WIN32_FIND_DATA



    If Right(Path, 1) <> "\" Then Path = Path & "\"


    nDir = 0
    ReDim dirNames(nDir) As String
    Cont = True

    hSearch = FindFirstFile(Path & "*", WFD)

    If hSearch <> INVALID_HANDLE_VALUE Then

        Do While Cont

            DirName = StripNulls(WFD.cFileName)

            If (DirName <> ".") And (DirName <> "..") Then

                ' found a dir?
                If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then

                    dirNames(nDir) = DirName
                    DirCount = DirCount + 1
                    nDir = nDir + 1
                    ReDim Preserve dirNames(nDir)

                End If

            End If

            Cont = FindNextFile(hSearch, WFD)

        Loop

        Cont = FindClose(hSearch)

    End If

    hSearch = FindFirstFile(Path & SearchStr, WFD)

    Cont = True

    If hSearch <> INVALID_HANDLE_VALUE Then

        While Cont

            FileName = StripNulls(WFD.cFileName)

            If (FileName <> ".") And (FileName <> "..") Then

                'found a file
                FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow

                If Not DirExists(Path & FileName) Then
                    FileCount = FileCount + 1
                    AddFile Path & FileName, AddSlash(strISOStartDir) & Mid$(Path, Len(strStartDir) + 1)
                End If

            End If

            Cont = FindNextFile(hSearch, WFD)

        Wend

        Cont = FindClose(hSearch)

    End If

    If nDir > 0 Then

        For i = 0 To nDir - 1

            NewDir AddSlash(strISOStartDir) & Mid$(Path, Len(strStartDir) + 1), dirNames(i)

            FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", _
                                                       SearchStr, _
                                                       FileCount, _
                                                       DirCount)

        Next i

    End If

End Function

Private Function To16BitLS(ByVal l As Long) As String

    Dim b0 As Byte, b1 As Byte

    b0 = (l And &HFF&)
    b1 = (l And &HFF00&) \ &H100&

    To16BitLS = Chr(b1) & Chr(b0)

End Function

Private Function To16BitMS(ByVal l As Long) As String

    Dim b0 As Byte, b1 As Byte

    b0 = (l And &HFF&)
    b1 = (l And &HFF00&) \ &H100&

    To16BitMS = Chr(b0) & Chr(b1)

End Function

Private Function To16BitBothByte(ByVal l As Long) As String

    Dim b0 As Byte, b1 As Byte

    b0 = (l And &HFF&)
    b1 = (l And &HFF00&) \ &H100&

    To16BitBothByte = Chr(b0) & Chr(b1) & Chr(b1) & Chr(b0)

End Function

Private Function To32BitMS(ByVal l As Long) As String

    Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte

    b0 = (l And &HFF&)
    b1 = (l And &HFF00&) \ &H100&
    b2 = (l And &HFF0000) \ &H10000
    b3 = (l And &H7F000000) \ &H1000000 - 128 * (l < 0)

    To32BitMS = Chr(b0) & Chr(b1) & Chr(b2) & Chr(b3)

End Function

Private Function To32BitLS(ByVal l As Long) As String

    Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte

    b0 = (l And &HFF&)
    b1 = (l And &HFF00&) \ &H100&
    b2 = (l And &HFF0000) \ &H10000
    b3 = (l And &H7F000000) \ &H1000000 - 128 * (l < 0)

    To32BitLS = Chr(b3) & Chr(b2) & Chr(b1) & Chr(b0)

End Function

Private Function To32BitBothByte(ByVal l As Long) As String

    Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte

    b0 = (l And &HFF&)
    b1 = (l And &HFF00&) \ &H100&
    b2 = (l And &HFF0000) \ &H10000
    b3 = (l And &H7F000000) \ &H1000000 - 128 * (l < 0)

    To32BitBothByte = Chr(b0) & Chr(b1) & Chr(b2) & Chr(b3) & _
                      Chr(b3) & Chr(b2) & Chr(b1) & Chr(b0)

End Function

Private Function To_D_Chars(ByVal str As String) As String

    Dim sBuf As String, i As Integer

    str = UCase(str)

    For i = 1 To Len(str)

        If InStr(1, d_chars, Mid(str, i, 1)) > 0 Then _
            sBuf = sBuf & Mid(str, i, 1)

    Next

    To_D_Chars = sBuf

End Function

Private Function To_A_Chars(ByVal str As String) As String

    Dim sBuf As String, i As Integer

    str = UCase(str)

    For i = 1 To Len(str)

        If InStr(1, a_chars, Mid(str, i, 1)) > 0 Then _
            sBuf = sBuf & Mid(str, i, 1)

    Next

    To_A_Chars = sBuf

End Function

Private Function CreateDateTime() As String
    CreateDateTime = format(Year(Date), "0000") & _
                     format(Month(Date), "00") & _
                     format(Day(Date), "00") & _
                     format(Hour(Time), "00") & _
                     format(Minute(Time), "00") & _
                     format(Second(Time), "00") & _
                     "00" & Chr(0)
End Function

Private Function CreateDateTimeRec() As String
    CreateDateTimeRec = Chr(Year(Date) - 1900) & _
                        Chr(Month(Date)) & _
                        Chr(Day(Date)) & _
                        Chr(Hour(Time)) & _
                        Chr(Minute(Time)) & _
                        Chr(Second(Time)) & Chr(0)
End Function

Public Function LoadProject(ByVal strLocal As String) As Boolean

    On Error GoTo ErrHandler

    Dim FF      As Integer:  FF = FreeFile
    Dim str     As String:  str = Space(7)
    Dim lng     As Long

    Open strLocal For Binary As FF

        Get FF, , str

        If Not str = "ISO9660" Then Exit Function

        'load dirs
        Get FF, , dircnt
        Get FF, , lng

        str = Space(lng)

        Get FF, , str

        dirs = Split(str, Chr$(255))

        'load files
        Get FF, , filecnt
        Get FF, , lng

        str = Space(lng)

        Get FF, , str

        files = Split(str, Chr$(255))

    Close FF

    LoadProject = True

ErrHandler:

End Function

Public Function SaveProject(ByVal strLocal As String) As Boolean

    On Error GoTo ErrHandler

    Dim FF   As Integer: FF = FreeFile

    If FileExists(strLocal) Then Kill strLocal

    Open strLocal For Binary As FF


        Put FF, , CStr("ISO9660")
        Put FF, , dircnt
        Put FF, , CLng(Len(Join(dirs, Chr$(255))))
        Put FF, , CStr(Join(dirs, Chr$(255)))
        Put FF, , filecnt
        Put FF, , CLng(Len(Join(files, Chr$(255))))
        Put FF, , CStr(Join(files, Chr$(255)))

    Close FF

    SaveProject = True

ErrHandler:

End Function

Public Sub ClearDirsFiles()

    ReDim files(0) As String
    filecnt = 0
    ReDim dirs(0) As String
    dircnt = 0

    Class_Initialize

End Sub

Private Sub Class_Initialize()

    'create root dir
    ReDim dirs(dircnt) As String
    dirs(dircnt) = "" & Chr$(0) & "\"
    dircnt = dircnt + 1

End Sub
