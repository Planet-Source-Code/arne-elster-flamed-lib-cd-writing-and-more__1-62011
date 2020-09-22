VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDataCD 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Flamed v4 - Data CD Project"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   495
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox cboDrv 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   675
      Style           =   2  'Dropdown-Liste
      TabIndex        =   9
      Top             =   690
      Width           =   6540
   End
   Begin prjFlamed.XP_ProgressBar prgUsed 
      Height          =   240
      Left            =   750
      TabIndex        =   7
      Top             =   4875
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   6011071
      Scrolling       =   9
      Value           =   50
   End
   Begin VB.CommandButton cmdDrvNfo 
      Caption         =   "Drive information"
      Height          =   315
      Left            =   4125
      TabIndex        =   6
      Top             =   5250
      Width           =   1440
   End
   Begin MSComDlg.CommonDialog dlgISO 
      Left            =   4350
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "ISO images (*.iso)|*.iso"
      Flags           =   2
   End
   Begin MSComctlLib.ImageList img 
      Left            =   3225
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataCD.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataCD.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataCD.frx":0B44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgFLD 
      Left            =   3825
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "FLD projects (*.fld)|*.fld"
      Flags           =   2
   End
   Begin prjFlamed.isButton cmdOptions 
      Height          =   360
      Left            =   150
      TabIndex        =   5
      Top             =   120
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   635
      Icon            =   "frmDataCD.frx":10DE
      Style           =   4
      Caption         =   " "
      IconAlign       =   0
      iNonThemeStyle  =   4
      USeCustomColors =   -1  'True
      BackColor       =   8296877
      HighlightColor  =   8296877
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write project"
      Default         =   -1  'True
      Height          =   330
      Left            =   5775
      TabIndex        =   4
      Top             =   5250
      Width           =   1365
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "Back"
      Height          =   330
      Left            =   225
      TabIndex        =   3
      Top             =   5250
      Width           =   1365
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   3690
      Left            =   3150
      TabIndex        =   1
      Top             =   1125
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   6509
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2381
      EndProperty
   End
   Begin MSComctlLib.TreeView lstDirs 
      Height          =   3690
      Left            =   75
      TabIndex        =   2
      Top             =   1125
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   6509
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   21
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img"
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "Drive:"
      Height          =   195
      Left            =   150
      TabIndex        =   10
      Top             =   750
      Width           =   435
   End
   Begin VB.Label lblUsed 
      AutoSize        =   -1  'True
      Caption         =   "Used:"
      Height          =   195
      Left            =   225
      TabIndex        =   8
      Top             =   4875
      Width           =   420
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data CD Project"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   750
      TabIndex        =   0
      Top             =   120
      Width           =   1980
   End
   Begin VB.Shape shpHdr 
      BackColor       =   &H00336471&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   7440
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuVD 
         Caption         =   "Volume Descriptors..."
      End
      Begin VB.Menu mnuS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear project"
      End
      Begin VB.Menu mnuS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSavePrj 
         Caption         =   "Save project..."
      End
      Begin VB.Menu mnuLoadPrj 
         Caption         =   "Load project..."
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveToISO 
         Caption         =   "Save to ISO..."
      End
   End
   Begin VB.Menu mnuRClick 
      Caption         =   "RClick"
      Visible         =   0   'False
      Begin VB.Menu mnuNewDir 
         Caption         =   "New directory"
      End
      Begin VB.Menu mnuRemDir 
         Caption         =   "Remove directory"
      End
   End
   Begin VB.Menu mnuRClickF 
      Caption         =   "RClickF"
      Visible         =   0   'False
      Begin VB.Menu mnuRemFiles 
         Caption         =   "Remove selected file(s)"
      End
   End
End
Attribute VB_Name = "frmDataCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cISO     As FL_ISO9660Writer
Attribute cISO.VB_VarHelpID = -1
Private WithEvents cISOCD   As FL_CDDataWriter
Attribute cISOCD.VB_VarHelpID = -1

Private cDrvNfo             As New FL_DriveInfo
Private cCDNfo              As New FL_CDInfo
Private cTrkNfo             As New FL_TrackInfo

' some public properties to be set by frmDataCDSettings/frmVD

Public Property Let TestMode(aval As Boolean)
    cISOCD.TestMode = aval
End Property

Public Property Let Finalize(aval As Boolean)
    cISOCD.NextSessionAllowed = Not aval
End Property

Public Property Let EjectDisk(aval As Boolean)
    cISOCD.EjectAfterWrite = aval
End Property

Public Property Let OnTheFly(aval As Boolean)
    cISOCD.OnTheFly = aval
End Property

Public Property Get VolumeID() As String
    VolumeID = cISO.VolumeID
End Property

Public Property Let VolumeID(val As String)
    cISO.VolumeID = val
    Me.Caption = "Flamed v4 - Data CD Project - " & cISO.VolumeID
    lstDirs.Nodes(1).Text = "\ [" & cISO.VolumeID & "]"
End Property

Public Property Get SystemID() As String
    SystemID = cISO.SystemID
End Property

Public Property Let SystemID(val As String)
    cISO.SystemID = val
End Property

Public Property Get AppID() As String
    AppID = cISO.AppID
End Property

Public Property Let AppID(val As String)
    cISO.AppID = val
End Property

Public Property Get PublisherID() As String
    PublisherID = cISO.PublisherID
End Property

Public Property Let PublisherID(val As String)
    cISO.PublisherID = val
End Property

Private Sub cboDrv_Click()
    strDrvID = cManager.DrvChr2DrvID(Left$(cboDrv.List(cboDrv.ListIndex), 1))
    UpdateUsedSpace
End Sub

Private Sub cISO_Progress(ByVal lngMax As Long, ByVal lngValue As Long)
    If Not frmSimplePrg.prg.Max = lngMax Then
        frmSimplePrg.prg.Max = lngMax
    End If
    frmSimplePrg.prg.Value = lngValue
End Sub

Private Sub cmdBack_Click()
    frmSelectProject.Show
    Unload Me
End Sub

Private Sub cmdDrvNfo_Click()
    frmDriveInfo.Show vbModal, Me
End Sub

Private Sub cmdWrite_Click()

    On Error GoTo ErrorHandler
    Dim intFiles    As Integer
    intFiles = UBound(cISO.GetLocalFiles)
    On Error GoTo 0

    If UpdateUsedSpace Then
        If MsgBox("The remaining space is smaller then the amount of data to write." & vbCrLf & _
                  "Continue?", vbYesNo Or vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    frmDataCDSettings.Show vbModal, Me
    Exit Sub

ErrorHandler:
    MsgBox "No files in the project.", vbExclamation

End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmSelectProject.Show
    Unload Me
End Sub

Private Sub mnuLoadPrj_Click()

    On Error GoTo ErrorHandler

    dlgFLD.FileName = vbNullString
    dlgFLD.ShowOpen

    If Not cISO.LoadProject(dlgFLD.FileName) Then
        MsgBox "Failed to load project.", vbExclamation, "Error"
    End If

    ' show directories
    lstFiles.ListItems.Clear
    AddRootNode

    ListDir "\"

    ' make sure files in root are shown
    lstDirs_Click

    UpdateUsedSpace

ErrorHandler:

End Sub

Private Sub mnuNewDir_Click()

    Dim strDir  As String

    strDir = InputBox("New directory's name:")
    If strDir = vbNullString Then Exit Sub

    If Not cISO.NewDir(lstDirs.SelectedItem.Key, strDir) Then
        MsgBox "Couldn't create the directory"
        Exit Sub
    End If

    strDir = AddSlash(lstDirs.SelectedItem.Key) & strDir

    AddRootNode
    ListDir "\"

    ' make sure the new directory is shown
    With lstDirs.Nodes(strDir)
        .Expanded = True
        .EnsureVisible
        .Selected = True
    End With

End Sub

Private Sub mnuRemDir_Click()

    'remove the selected directory
    If Not cISO.RemDir(lstDirs.SelectedItem.Key) Then
        MsgBox "failed", vbExclamation
        Exit Sub
    End If

    'build the directory list
    lstFiles.ListItems.Clear
    AddRootNode
    ListDir "\"

    UpdateUsedSpace

End Sub

Private Sub mnuRemFiles_Click()

    Dim i   As Integer

    ' remove selected file
    With lstFiles.ListItems

        For i = .count To 1 Step -1
            If .Item(i).Selected Then

                ' selected, so remove it
                If cISO.RemFile(.Item(i).Key) Then
                    ' remove file from listview
                    .Remove i
                Else
                    MsgBox "Failed to remove " & .Item(i).Key, vbExclamation, "Error"
                End If

            End If
        Next

    End With

    UpdateUsedSpace

End Sub

Private Sub mnuSaveToISO_Click()

    On Error GoTo ErrorHandler

    dlgISO.FileName = vbNullString
    dlgISO.ShowSave

    frmSimplePrg.Show , Me
    frmSimplePrg.lblStat = "Saving ISO image..."

    If Not cISO.CreateISO(dlgISO.FileName) Then
        MsgBox "Failed to save ISO image.", vbExclamation, "Error"
    Else
        MsgBox "Finished.", vbInformation, "Ok"
    End If

    Unload frmSimplePrg

ErrorHandler:

End Sub

Private Sub mnuSavePrj_Click()

    On Error GoTo ErrorHandler

    dlgFLD.FileName = vbNullString
    dlgFLD.ShowSave

    If Not cISO.SaveProject(dlgFLD.FileName) Then
        MsgBox "Failed to save project.", vbExclamation, "Error"
    End If

ErrorHandler:

End Sub

Private Sub lstFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i   As Long

    'add dropped dirs
    For i = 1 To Data.Files.count
        If DirExists(Data.Files(i)) Then
            If Not cISO.AddDir(Data.Files(i), lstDirs.SelectedItem.Key) Then
                MsgBox "Couldn't add directory " & Data.Files(i), vbExclamation, "Error"
            End If
        ElseIf FileExists(Data.Files(i)) Then
            If Not cISO.AddFile(Data.Files(i), lstDirs.SelectedItem.Key) Then
                MsgBox "Couldn't add file " & Data.Files(i), vbExclamation, "Error"
            End If
        End If
    Next

    ' build new directory structure
    AddRootNode
    ListDir "\"

    ' show files for current dir
    lstDirs_Click

    UpdateUsedSpace

End Sub

Private Sub lstFiles_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete: mnuRemFiles_Click
    End Select
End Sub

Private Sub lstFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuRClickF
End Sub

Private Sub lstDirs_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i   As Long

    'add dropped dirs
    For i = 1 To Data.Files.count
        If DirExists(Data.Files(i)) Then
            If Not cISO.AddDir(Data.Files(i), lstDirs.SelectedItem.Key) Then
                MsgBox "Couldn't add directory " & Data.Files(i), vbExclamation, "Error"
            End If
        ElseIf FileExists(Data.Files(i)) Then
            If Not cISO.AddFile(Data.Files(i), lstDirs.SelectedItem.Key) Then
                MsgBox "Couldn't add file " & Data.Files(i), vbExclamation, "Error"
            End If
        End If
    Next

    ' build new directory structure
    AddRootNode
    ListDir "\"

    ' show files for current dir
    lstDirs_Click

    UpdateUsedSpace

End Sub

Private Sub lstDirs_Click()

    ' if a directory is empty,
    ' strFiles() will have no bounds
    On Error Resume Next

    Dim i           As Integer
    Dim strFiles()  As String

    ' get the files of the selected dir
    strFiles = cISO.GetFiles(lstDirs.SelectedItem.Key)
    ' and show them
    lstFiles.ListItems.Clear
    For i = 0 To UBound(strFiles)
        ' filename
        With lstFiles.ListItems.Add(, AddSlash(lstDirs.SelectedItem.Key) & strFiles(i), strFiles(i), , 3)
            ' filesize
            .SubItems(1) = FormatFileSize(cISO.GetFileDetailByPath(.Key, FD_FileSize))
        End With
    Next

End Sub

Private Sub lstDirs_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown: lstDirs_Click
    End Select
End Sub

Private Sub lstDirs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuRClick
End Sub

Private Sub cmdOptions_Click()
    PopupMenu mnuOptions, , cmdOptions.Left, cmdOptions.Top + cmdOptions.Height
End Sub

Private Sub AddRootNode()
    lstDirs.Nodes.Clear
    lstDirs.Nodes.Add(, , "\", "\ [" & VolumeID & "]", 2).Selected = True
End Sub

Public Function FileExists(ByVal Path As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(Path) And (vbDirectory Or vbVolume)) = 0
End Function

Public Function DirExists(ByVal Path As String) As Boolean
    On Error Resume Next
    DirExists = CBool(GetAttr(Path) And vbDirectory)
End Function

'format bytes
Private Function FormatFileSize(ByVal dblFileSize As Double, _
    Optional ByVal strFormatMask As String) As String

    Select Case dblFileSize
        Case 0 To 1023               ' Bytes
            FormatFileSize = Format(dblFileSize) & " bytes"
        Case 1024 To 1048575         ' KB
            If strFormatMask = Empty Then strFormatMask = "###0"
            FormatFileSize = Format(dblFileSize \ 1024, strFormatMask) & " KB"
        Case 1024# ^ 2 To 1073741823 ' MB
            If strFormatMask = Empty Then strFormatMask = "####0.00"
            FormatFileSize = Format(dblFileSize \ (1024 ^ 2), strFormatMask) & " MB"
        Case Is > 1073741823#        ' GB
            If strFormatMask = Empty Then strFormatMask = "####0.00"
            FormatFileSize = Format(dblFileSize \ (1024 ^ 3), strFormatMask) & " GB"
    End Select

End Function

Private Sub ListDir(ByVal strPath As String)

    Dim dirs()  As String
    Dim i       As Integer

    'if the directory has sub directories
    If cISO.DirCount(strPath) > 0 Then

        'get them all
        dirs = cISO.GetDirs(strPath)

        For i = 0 To UBound(dirs)

            'and add dem...
            lstDirs.Nodes.Add strPath, tvwChild, AddSlash(strPath) & dirs(i), dirs(i), 1

            '...and all of their sub directories
            ListDir AddSlash(strPath) & dirs(i)

        Next

    End If

End Sub

Private Sub mnuClear_Click()

    cISO.ClearDirsFiles
    lstFiles.ListItems.Clear
    AddRootNode

    UpdateUsedSpace

End Sub

Private Sub Form_Load()

    Set cISO = New FL_ISO9660Writer
    Set cISOCD = New FL_CDDataWriter

    cISOCD.TempDir = GetSetting("Flamedv4", "DataCD", "temp", cISOCD.TempDir)

    ShowDrives

    ' show root
    AddRootNode

    VolumeID = "NEW_VOLUME"
    SystemID = "WIN32"
    PublisherID = "FLAMEDv4"
    AppID = "FLAMEDv4 DATA WRITER"

End Sub

Private Sub mnuVD_Click()
    frmVD.Show vbModal, Me
End Sub

Public Sub Burn()

    Dim strMsg  As String

    Set cISOCD.ISOClass = cISO

    Me.Hide
    frmDataCDPrg.Show

    Select Case cISOCD.WriteISOtoCD(strDrvID)
        Case BURNRET_CLOSE_SESSION: strMsg = "Could not close session."
        Case BURNRET_CLOSE_TRACK: strMsg = "Could not cloe track."
        Case BURNRET_FILE_ACCESS: strMsg = "Failed to access a file."
        Case BURNRET_INVALID_MEDIA: strMsg = "Invalid medium in drive."
        Case BURNRET_ISOCREATION: strMsg = "ISO creation failed."
        Case BURNRET_NO_NEXT_WRITABLE_LBA: strMsg = "Could not get next writable LBA."
        Case BURNRET_NOT_EMPTY: strMsg = "Disk is finalized."
        Case BURNRET_OK: strMsg = "Finished."
        Case BURNRET_SYNC_CACHE: strMsg = "Could not synchronize cache."
        Case BURNRET_WPMP: strMsg = "Write Parameters Page invalid"
        Case BURNRET_WRITE: strMsg = "Write error (Buffer Underrun?)"
    End Select

    MsgBox strMsg, vbInformation

    Me.Show
    Unload frmDataCDSettings
    Unload frmDataCDPrg

End Sub

Property Let ISOClass(aval As FL_ISO9660Writer)
    Set cISO = aval
End Property

Private Sub cISOCD_CheckForFiles()
    With frmDataCDPrg.lstStatus.ListItems.Add(SmallIcon:=1)
        .SubItems(1) = "Checking for presence of files..."
    End With
End Sub

Private Sub cISOCD_ClosingSession()
    With frmDataCDPrg.lstStatus.ListItems.Add(SmallIcon:=1)
        .SubItems(1) = "Closing session..."
    End With
End Sub

Private Sub cISOCD_FilesMissing(strFiles() As String)

    With frmDataCDPrg.lstStatus.ListItems.Add(SmallIcon:=1)
        .SubItems(1) = "Files are missing."
    End With

    MsgBox "Files are missing: " & vbCrLf & Join(strFiles, ","), vbExclamation

End Sub

Private Sub cISOCD_Finished()
    With frmDataCDPrg.lstStatus.ListItems.Add(SmallIcon:=1)
        .SubItems(1) = "Finished."
    End With
End Sub

Private Sub cISOCD_ISOProgress(ByVal lngMax As Long, ByVal lngValue As Long)
    On Error Resume Next
    If Not frmDataCDPrg.prg.Max = lngMax Then frmDataCDPrg.prg.Max = lngMax
    frmDataCDPrg.prg.Value = lngValue
    DoEvents
End Sub

Private Sub cISOCD_StartWriting()
    With frmDataCDPrg.lstStatus.ListItems.Add(SmallIcon:=1)
        .SubItems(1) = "Writing data track..."
    End With
    DoEvents
End Sub

Private Sub cISO_ISOStart()
    With frmDataCDPrg.lstStatus.ListItems.Add(SmallIcon:=1)
        .SubItems(1) = "Creating file system..."
    End With
    DoEvents
End Sub

Private Sub cISOCD_WriteProgress(ByVal Percent As Integer)
    On Error Resume Next
    If Not frmDataCDPrg.prg.Max = 100 Then frmDataCDPrg.prg.Max = 100
    frmDataCDPrg.prg.Value = Percent
End Sub

Private Function UpdateUsedSpace() As Boolean

    On Error Resume Next

    Dim lngFree As Long

    cCDNfo.GetInfo strDrvID
    cTrkNfo.GetInfo strDrvID, cCDNfo.Tracks

    lngFree = (cCDNfo.Capacity - (cTrkNfo.TrackEnd.LBA * 2048&)) \ 2048&

    prgUsed.Max = lngFree
    prgUsed.Value = cISO.ISOSize / 2048&

    UpdateUsedSpace = cISO.ISOSize / 2048& > prgUsed.Max

End Function

Private Sub ShowDrives()

    Dim strDrives() As String
    Dim i           As Long

    strDrives = GetDriveList(OPT_CDWRITERS)

    For i = LBound(strDrives) To UBound(strDrives) - 1

        cDrvNfo.GetInfo cManager.DrvChr2DrvID(strDrives(i))

        With cDrvNfo
            cboDrv.AddItem strDrives(i) & ": " & _
                           .Vendor & " " & _
                           .Product & " " & _
                           .Revision & " [" & _
                           .HostAdapter & ":" & _
                           .Target & "]"
        End With

    Next

    cboDrv.ListIndex = 0

End Sub
