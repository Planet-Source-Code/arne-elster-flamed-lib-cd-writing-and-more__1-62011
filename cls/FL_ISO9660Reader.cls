VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_ISO9660Reader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'****************************************
'*** ISO9660 Reader Class
'****************************************

Option Explicit

Private Type ISO_VOL_DESCRIPTOR
    Vol_Desc_Type                 As Byte          ' Volume Descriptor Type (1)
    Std_ID                        As String * 6    ' Standard Identifier (CD001 )
    Vol_Desc_Ver                  As Byte          ' Volume Descriptor Version (0)
End Type

Private Type ISO_HEADER
     zeros                        As String        ' 32767 Chr$(0)s
     Vol_Desc                     As ISO_VOL_DESCRIPTOR ' Primary Volume Descriptor (PVD)
End Type

Private Type PTableRecord
    Length                        As Byte          ' length of the record
    ExtAttrRecLen                 As Byte          ' Extended Attribute Record length
    FirstSector                   As String * 4    ' first sector of the directory
    RecNumParentDir               As String * 2    ' Record Number of the parent
End Type

Private Type DirRecord
    Length                        As Byte          ' length of the record
    ExtAttrRecLen                 As Byte          ' Extended Attribute Record length
    LocExt                        As String * 8    ' sector with the extent
    DataLen                       As String * 8    ' data length
    RecDateTime                   As String * 7    ' date of the record
    FileFlags                     As Byte          ' file flags
    FileUnitSize                  As Byte          ' File Unit size
    InterleaveGapSize             As Byte          ' InterleaveGapSize
    VolSeqNum                     As String * 4    ' Volume Sequence Number
    FileIDLen                     As Byte          ' filename length
End Type

Private Type RootDirRecord
    Length                        As Byte          ' length of the record
    ExtAttrRecLen                 As Byte          ' Extended Attribute Record length
    LocExt                        As String * 8    ' sector with the extent
    DataLen                       As String * 8    ' data length
    RecDateTime                   As String * 7    ' date of the record
    FileFlags                     As Byte          ' file flags
    FileUnitSize                  As Byte          ' File Unit size
    InterleaveGapSize             As Byte          ' InterleaveGapSize
    VolSeqNum                     As String * 4    ' Volume Sequence Number
    FileIDLen                     As Byte          ' filename length
    FileID(0)                     As Byte          ' filename
End Type

Private Type ISO_PRIMARY_DESCRIPTOR
    System_ID                     As String * 32   ' System ID
    Volume_ID                     As String * 32   ' Volume ID

    unused2(1 To 8)               As Byte          ' 8 unused Bytes

    Vol_Space_Size                As String * 8    ' volume size
    unused3(1 To 32)              As Byte          ' 32 unused Bytes
    Vol_Set_Size                  As String * 4    ' Volume Set size (1)
    Vol_Seq_Number                As String * 4    ' Volume Sequence Number (1)
    Logical_Block_Size            As String * 4    ' logical blocksize (2048)

    Path_Table_Size               As String * 8    ' path table size
    Location_TypeL_Path_Table     As String * 4    ' location of type L path table
    Location_TypeL_Opt_Path_Table As String * 4    ' location of opt. Type L path table
    Location_TypeM_Path_Table     As String * 4    ' location of Type M path table
    Location_TypeM_Opt_Oath_Table As String * 4    ' location of opt. Type M path table
    Dir_Root_Record               As RootDirRecord ' Directory Record for Root

    Vol_Set_ID                    As String * 128  ' Volume Set Identifier
    Publisher_ID                  As String * 128  ' Publisher Identifier
    Data_Preparer_ID              As String * 128  ' Data Preparer Identifier
    App_ID                        As String * 128  ' Application Identifier

    Copyright_File_ID             As String * 37   ' Copyright File Identifier
    Abstract_File_ID              As String * 37   ' Abstract File Identifier
    Bibliographic_File_ID         As String * 37   ' Bibliographic File Idenfitier

    Vol_Creation_Date             As String * 17   ' creation date
    Vol_Modification_Date         As String * 17   ' modification date
    Vol_Expiration_Date           As String * 17   ' expiration date
    Vol_Effective_Date            As String * 17   ' effective date

    File_Structure_Ver            As Byte          ' file structure version
    rsvd                          As Byte          ' reserved
    app_use(1 To 512)             As Byte          ' reserved
    rsvd2(1 To 653)               As Byte          ' reserved
End Type

Private Enum eFileFlags
    FLAG_HIDDEN = &H1                              ' Record Flag Hidden
    FLAG_DIRECTORY = &H2                           ' Record Flag Directory
    FLAG_ASSOCIATED = &H4                          ' Record Flag associated file
    FLAG_RECFMTSPEC = &H8                          ' Record Flag Fecord Format Specific
    FLAG_PERM = &H16                               ' Record Flag Permission
    FLAG_NOTFINAL = &H80                           ' Record Flag Not Final Version
End Enum

Private Enum eVolDesc
    VOL_DESC_BOOT_RECORD                           ' Boot Record
    VOL_DESC_PRIM_VOL_DESCRIPTOR                   ' Primary Volume Descriptor
    VOL_DESC_SUPPL_VOL_DESCRIPTOR                  ' Supplementary Volume Descriptor
    VOL_DESC_VOL_PART_DESCRIPTOR                   ' Volume Partition Descriptor
    VOL_DESC_SET_TERMINATOR                        ' Volume Descriptor Set Terminator
End Enum

Public Event ReadProgress(ByVal Percent As Integer)

Private sVolID         As String                   ' Volume ID
Private sSysID         As String                   ' System ID
Private sAppID         As String                   ' Appplication ID
Private sPubID         As String                   ' Publisher ID
Private sPrepID        As String                   ' Data Preparer ID
Private sVolCreateDate As String                   ' creation date
Private sVolModDate    As String                   ' modification date
Private sVolExpDate    As String                   ' expiration date
Private sVolEffDate    As String                   ' effective date
Private sAbstrFile     As String                   ' Abstract File ID
Private sBiblioFile    As String                   ' Bibliographic File ID
Private sCopyFile      As String                   ' Copyright File ID
Private lVolSize       As Long                     ' Volume size
Private sCurrFile      As String                   ' current image

Private dirs()         As String                   ' directories
Private files()        As String                   ' files
Private filecnt        As Long                     ' number of files

Public Property Get CurrentISO() As String
    CurrentISO = sCurrFile
End Property

Public Property Get VolumeID() As String
    VolumeID = sVolID
End Property

Public Property Get SystemID() As String
    SystemID = sSysID
End Property

Public Property Get ApplicationID() As String
    ApplicationID = sAppID
End Property

Public Property Get PublisherID() As String
    PublisherID = sPubID
End Property

Public Property Get DataPreparerID() As String
    DataPreparerID = sPrepID
End Property

Public Property Get VolumeCreationDate() As String
    VolumeCreationDate = sVolCreateDate
End Property

Public Property Get VolumeModificationDate() As String
    VolumeModificationDate = sVolModDate
End Property

Public Property Get VolumeExpirationDate() As String
    VolumeExpirationDate = sVolExpDate
End Property

Public Property Get VolumeEffectiveDate() As String
    VolumeEffectiveDate = sVolEffDate
End Property

Public Property Get AbstractFile() As String
    AbstractFile = sAbstrFile
End Property

Public Property Get BibliographicFile() As String
    BibliographicFile = sBiblioFile
End Property

Public Property Get CopyrightFile() As String
    CopyrightFile = sCopyFile
End Property

Public Property Get VolumeSize() As Long
    VolumeSize = lVolSize
End Property

Public Function ReadISO(ByVal sPath As String) As Boolean
    Dim lPathTableStart As Long, lFileTableStart As Long
    Dim lPathTable As Long, lFileTable As Long
    Dim PrimDesc As ISO_PRIMARY_DESCRIPTOR
    Dim FF As Integer: FF = FreeFile
    Dim hdr As ISO_HEADER

    ' skip zeros and read PVD
    Open sPath For Binary As #FF
        Seek #FF, 32767
        Get #FF, , hdr
        Get #FF, , PrimDesc
    Close #FF

    ' must begin with CD001
    If Not Left$(hdr.Vol_Desc.Std_ID, 5) = "CD001" Then Exit Function

    ' reinit vars
    sCurrFile = sPath

    ReDim files(0): filecnt = 0
    ReDim dirs(0)

    sVolID = Trim$(PrimDesc.Volume_ID)
    sSysID = Trim$(PrimDesc.System_ID)
    sAppID = Trim$(PrimDesc.App_ID)
    sPubID = Trim$(PrimDesc.Publisher_ID)
    sPrepID = Trim$(PrimDesc.Data_Preparer_ID)
    sAbstrFile = Trim$(PrimDesc.Abstract_File_ID)
    sBiblioFile = Trim$(PrimDesc.Bibliographic_File_ID)
    sCopyFile = Trim$(PrimDesc.Copyright_File_ID)
    sVolCreateDate = VolDateToStr(PrimDesc.Vol_Creation_Date)
    sVolModDate = VolDateToStr(PrimDesc.Vol_Modification_Date)
    sVolExpDate = VolDateToStr(PrimDesc.Vol_Effective_Date)
    sVolEffDate = VolDateToStr(PrimDesc.Vol_Effective_Date)
    lVolSize = From32BitMS(Left$(PrimDesc.Vol_Space_Size, 4))

    ' read little endian path table
    lPathTable = From32BitMS(Left$(PrimDesc.Path_Table_Size, 4))
    lPathTableStart = From32BitMS(PrimDesc.Location_TypeL_Path_Table)
    ParseDirs ReadSector(sPath, lPathTableStart, GetCDFSSize(lPathTable) \ 2048), lPathTable

    ' read file system
    lFileTableStart = From32BitMS(Left$(PrimDesc.Dir_Root_Record.LocExt, 4))
    lFileTable = From32BitMS(Left$(PrimDesc.Dir_Root_Record.DataLen, 4))
    ParseFileTable sPath, lFileTableStart, lFileTable, "\"

    ' finished :)
    ReadISO = True

End Function

' parse file table
Private Sub ParseFileTable(ByVal sPath As String, ByVal lStart As Long, _
                           ByVal lLen As Long, ByVal sParent As String)

    On Error Resume Next

    Dim lRecLen As Byte, LocExt As Long, DataLen As Long, isDir As Boolean
    Dim sBuf As String, sName As String, sNameLen As Integer
    Dim dirz() As String, splt() As String
    Dim i As Long, dircnt As Long

    ' get content of directory record
    sBuf = ReadSector(sPath, lStart, lLen \ 2048)

    While i > -1

        ' get record length
        lRecLen = Asc(Mid$(sBuf, i + 1, 1))

        If lRecLen = 0 Then
            ' found no record, move to next position
            lRecLen = 1
        Else
            ' read record info
            LocExt = From32BitMS(Mid$(sBuf, i + 3, 4))              ' Startsektor des Inhalts
            DataLen = From32BitMS(Mid$(sBuf, i + 11, 4))            ' Datenlänge des Records
            isDir = Asc(Mid$(sBuf, i + 26, 1)) And FLAG_DIRECTORY   ' Verzeichnis oder Datei?
            sNameLen = Asc(Mid$(sBuf, i + 33, 1))                   ' Namenlänge
            sName = Mid$(sBuf, i + 34, sNameLen)                    ' Name

            ' name empty?
            If Not sName = vbNullString Then

                ' if first byte is greater then 1 we have a valid name
                If Asc(sName) > 1 Then
                    If isDir Then                                   ' found a directory

                        ReDim Preserve dirz(dircnt)
                        dirz(dircnt) = sName & Chr(0) & LocExt & Chr(0) & DataLen
                        dircnt = dircnt + 1

                    Else

                        ReDim Preserve files(filecnt)               ' found a file
                        If Right$(sName, 2) = ";1" Then sName = Left$(sName, Len(sName) - 2)
                        files(filecnt) = AddSlash(sParent) & sName & Chr(0) & LocExt & Chr(0) & DataLen
                        filecnt = filecnt + 1

                    End If
                End If
            End If
        End If

        ' move to next position
        i = i + lRecLen

        ' reached end of record?
        If i >= lLen Then i = -1

    Wend

    ' go through all found directories
    For i = 0 To UBound(dirz)
        DoEvents
        splt = Split(dirz(i), Chr(0))
        ParseFileTable sPath, splt(1), splt(2), AddSlash(sParent) & AddSlash(splt(0))
    Next

End Sub

' parse path table
Private Sub ParseDirs(ByVal sData As String, ByVal lDataLen As Long)
    Dim iLen As Byte, ParentDir As Long, sName As String
    Dim dircnt As Integer: dircnt = 1
    Dim i As Integer: i = 1

    While i > -1

        iLen = Asc(Mid$(sData, i, 1))                               ' length of directory name
        ParentDir = From16BitMS(Mid$(sData, i + 6, 2))              ' Record ID of parent
        sName = Mid$(sData, i + 8, iLen)                            ' directory name
        If sName = Chr(0) Then sName = "\"                          ' Chr(0) = "\" = Root

        ' name empty?
        If Not sName = vbNullString Then
            ReDim Preserve dirs(dircnt)

            If Not sName = "\" Then
                dirs(dircnt) = dirs(ParentDir) & AddSlash(sName)    ' add directory
            Else
                dirs(dircnt) = sName                                ' add root
            End If
        End If

        dircnt = dircnt + 1                                         ' new directory

        If i >= lDataLen Then Exit Sub                              ' end of path table?

        i = i + iLen + 8                                            ' move to next position

        If iLen Mod 2 Then i = i + 1                                ' no odd positions
        If iLen = 0 Then i = i + 1                                  ' we need to move forward
        If ParentDir = 0 Then i = i + 1                             ' we need to move forward
    Wend
End Sub

' a directory has subdirs?
Public Function HasSubDirs(ByVal sPath As String) As Boolean
    Dim i As Integer

    sPath = LCase(sPath)

    ' go through all found dirs
    For i = 1 To UBound(dirs)

        ' parent of current dir is sPath?
        If LCase(GetUpperDir(dirs(i))) = sPath Then

            ' yes, has subdirs
            HasSubDirs = True
            Exit Function

        End If

    Next
End Function

' directory has files?
Public Function HasSubFiles(ByVal sPath As String) As Boolean
    Dim i As Integer

    sPath = LCase(AddSlash(sPath))

    ' go through all found files
    For i = 0 To UBound(files)

        ' parent of file is sPath?
        If LCase(GetUpperDir(AddSlash(files(i)))) = sPath Then

            ' yes, has files
            HasSubFiles = True
            Exit Function

        End If

    Next
End Function

' return sub directories of sPath
Public Function GetSubDirs(ByVal sPath As String) As String()
    Dim dircnt As Integer, i As Integer
    Dim dirz() As String

    sPath = LCase(sPath)

    ' go through all found dirs
    For i = 1 To UBound(dirs)

        ' exclude root
        If Not (dirs(i) = "\" And sPath = "\") Then

            ' parent of directory is sPath?
            If LCase(GetUpperDir(dirs(i))) = sPath Then

                ' yes, add found directory
                ReDim Preserve dirz(dircnt)
                dirz(dircnt) = dirs(i)
                dircnt = dircnt + 1

            End If
        End If

    Next

    ' return found directories
    GetSubDirs = dirz

End Function

' return files in a directory
Public Function GetSubFiles(ByVal sPath As String) As String()
    On Error Resume Next

    Dim i As Integer, fcnt As Long
    Dim filez() As String

    sPath = LCase(AddSlash(sPath))

    ' go through all found files
    For i = 0 To UBound(files)

        ' parent of file is sPath?
        If LCase(GetUpperDir(AddSlash(files(i)))) = sPath Then

            ' yes, add file
            ReDim Preserve filez(fcnt)
            filez(fcnt) = Mid$(Left$(files(i), InStr(1, files(i), ChrW$(0)) - 1), Len(sPath) + 1)
            fcnt = fcnt + 1

        End If

    Next

    ' return found files
    GetSubFiles = filez

End Function

' return size of a file in the image
Public Function GetFilesize(ByVal sFile As String) As Long
    On Error Resume Next

    Dim sCurr As String, parts() As String
    Dim i As Integer

    sFile = LCase(sFile)

    ' go through all found files
    For i = 0 To UBound(files)

        ' match?
        If LCase(Left$(files(i), InStr(1, files(i), ChrW$(0)) - 1)) = sFile Then

            ' get size of file
            parts = Split(files(i), Chr(0))
            GetFilesize = parts(2)
            Exit Function

        End If

    Next
End Function

' read file from image to string
Public Function ReadFileToStr(ByVal sFile As String) As String
    Dim sCurr As String, parts() As String
    Dim i As Integer

    sFile = LCase(sFile)

    ' go through all found files
    For i = 0 To UBound(files)

        ' match?
        If LCase(Left$(files(i), InStr(1, files(i), ChrW$(0)) - 1)) = sFile Then

            ' get file info
            parts = Split(files(i), Chr(0))

            ' read file and return it
            ReadFileToStr = ReadFileStr(sCurrFile, parts(1) * 2048, parts(2))
            Exit Function

        End If

    Next
End Function

' extract file
Public Sub ReadFileToFile(ByVal sFile As String, ByVal sDest As String)
    Dim i As Integer, FF As Integer: FF = FreeFile
    Dim sCurr As String, parts() As String

    sFile = LCase(sFile)

    ' go through all found files
    For i = 0 To UBound(files)

        ' match?
        If LCase(Left$(files(i), InStr(1, files(i), ChrW$(0)) - 1)) = sFile Then

            ' get file info
            parts = Split(files(i), Chr(0))

            ' open destination
            Open sDest For Binary As #FF

                ' read file from image and save it to destination file
                Put #FF, , ReadFileStr(sCurrFile, parts(1) * 2048, parts(2))

            Close #FF

            Exit Sub
        End If

    Next

End Sub

Private Function AddSlash(ByVal sVal As String) As String
    If Not Right$(sVal, 1) = "\" Then sVal = sVal & "\"
    AddSlash = sVal
End Function

Private Function GetCDFSSize(ByVal Length As Long) As Long
    Dim i As Integer, j As Long

    While Not i = -1
        j = j + 1
        If j * 2048 >= Length Then i = -1
    Wend

    GetCDFSSize = j * 2048
End Function

' << Operator
Private Function LShift(ByVal pnValue As Long, ByVal pnShift As Long) As Long
    LShift = pnValue * (2 ^ pnShift)
End Function

' >> Operator
Private Function RShift(ByVal pnValue As Long, ByVal pnShift As Long) As Long
    RShift = CLng(pnValue \ (2 ^ pnShift))
End Function

' 16 Bit LSB to VB6 Integer
Private Function From16BitLS(l As String) As Integer
    Dim b0 As Byte, b1 As Byte
    b0 = Asc(Right(l, 1))
    b1 = Asc(Left(l, 1))
    From16BitLS = b0 Or LShift(b1, 8)
End Function

' 16 Bit MSB to VB6 Integer
Private Function From16BitMS(l As String) As Integer
    Dim b0 As Byte, b1 As Byte
    b0 = Asc(Right(l, 1))
    b1 = Asc(Left(l, 1))
    From16BitMS = b1 Or LShift(b0, 8)
End Function

'32 Bit LSB to VB6 Long
Private Function From32BitLS(l As String) As Long
    Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte
    b0 = Asc(Left$(l, 1))
    b1 = Asc(Mid$(l, 2, 1))
    b2 = Asc(Mid$(l, 3, 1))
    b3 = Asc(Right$(l, 1))
    From32BitLS = LShift(b0, 24) Or LShift(b1, 16) Or LShift(b2, 8) Or b3
End Function

'32 Bit MSB to VB6 Long
Private Function From32BitMS(l As String) As Long
    Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte
    b0 = Asc(Left$(l, 1))
    b1 = Asc(Mid$(l, 2, 1))
    b2 = Asc(Mid$(l, 3, 1))
    b3 = Asc(Right$(l, 1))
    From32BitMS = LShift(b3, 24) Or LShift(b2, 16) Or LShift(b1, 8) Or b0
End Function

' format Volume Descriptor date
Private Function VolDateToStr(ByVal sDate As String) As String
    Dim sYear As String, sMonth As String, sDay As String
    Dim sHour As String, sMinute As String, sSecond As String

    sYear = Left$(sDate, 4)
    sMonth = Mid$(sDate, 5, 2)
    sDay = Mid$(sDate, 7, 2)
    sHour = Mid$(sDate, 10, 2)
    sMinute = Mid$(sDate, 13, 2)
    sSecond = Mid$(sDate, 16, 2)

    VolDateToStr = sDay & "," & sMonth & "," & sYear & " " & sHour & ":" & sMinute & ":" & sSecond
End Function

Private Function GetDirname(ByVal sDir As String) As String
    Dim parts() As String
    parts = Split(AddSlash(sDir), "\")
    GetDirname = parts(UBound(parts) - 1)
End Function

Private Function GetUpperDir(ByVal sDir As String) As String
    Dim parts() As String, i As Integer

    sDir = AddSlash(sDir)
    parts = Split(sDir, "\")

    For i = 0 To UBound(parts) - 2
        GetUpperDir = GetUpperDir & parts(i) & "\"
    Next

    GetUpperDir = AddSlash(GetUpperDir)
End Function

Private Function IsSubDir(ByVal sDir1 As String, ByVal sDir2 As String) As Boolean
    If GetUpperDir(sDir2) = sDir1 Then IsSubDir = True
End Function

' read sectors from Mode-1 ISO image
Private Function ReadSector(ByVal sPath As String, ByVal lLBA As Long, _
                            Optional ByVal lSectors As Long = 1) As String

    Dim FF As Integer: FF = FreeFile

    Open sPath For Binary As #FF
        ReadSector = Space(lSectors * 2048)
        Seek FF, lLBA * 2048 + 1
        Get #FF, , ReadSector
    Close #FF
End Function

' read a part from a file
Private Function ReadFileStr(ByVal sFile As String, ByVal Start As Long, _
                             ByVal Length As Long) As String

    '   64 KB/read
    Const ChunkSize     As Long = 64& * 1024&
    '   String Bilder
    Dim cString         As New clsStringBuilder
    '   String Buffer
    Dim strBuffer       As String

    cString.ChunkSize = ChunkSize

    Dim FF As Integer: FF = FreeFile

    Open sFile For Binary As #FF

        Start = Start + 1
        Length = Start + Length

        ' seek to start of requested part
        Seek #FF, Start

        Do While Not Seek(FF) >= Length

            ' next read will go over the bounds?
            If Seek(FF) + ChunkSize > Length Then
                strBuffer = Space(Length - Seek(FF))
            Else
                strBuffer = Space(ChunkSize)
            End If

            ' get data
            Get #FF, , strBuffer
            ' append string data to buffer
            cString.Append strBuffer

            RaiseEvent ReadProgress((Seek(FF) - Start) / (Length - Start) * 100)

        Loop

    Close #FF

    ' return string buffer
    ReadFileStr = cString.ToString

End Function
