VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Flamed v4 Data CD Writer"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.ProgressBar prg 
      Height          =   240
      Left            =   225
      TabIndex        =   7
      Top             =   3675
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList img 
      Left            =   5025
      Top             =   525
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
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Oben ausrichten
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      ToolTips        =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
      EndProperty
      Begin VB.ComboBox cboDrvs 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   75
         Style           =   2  'Dropdown-Liste
         TabIndex        =   6
         Top             =   15
         Width           =   3915
      End
      Begin VB.ComboBox cboSpeed 
         Height          =   315
         ItemData        =   "frmMain.frx":186E
         Left            =   4050
         List            =   "frmMain.frx":1870
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Top             =   15
         Width           =   1665
      End
   End
   Begin MSComDlg.CommonDialog dlgFLA 
      Left            =   4650
      Top             =   1650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "FLA projects (*.fla)|*.fla"
      Flags           =   2
   End
   Begin MSComDlg.CommonDialog dlgISO 
      Left            =   5175
      Top             =   1650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "ISO images (*.iso)|*.iso"
      Flags           =   2
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   2865
      Left            =   2475
      TabIndex        =   2
      Top             =   450
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   5054
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
      Height          =   2865
      Left            =   75
      TabIndex        =   1
      Top             =   450
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   5054
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
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Unten ausrichten
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   5070
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   423
      Style           =   1
      SimpleText      =   "Drop files/directories"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frmPrg 
      Caption         =   "Progress"
      Height          =   690
      Left            =   75
      TabIndex        =   3
      Top             =   3375
      Width           =   5640
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu S6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadPrj 
         Caption         =   "Load project..."
      End
      Begin VB.Menu mnuSavePrj 
         Caption         =   "Save project..."
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveISO 
         Caption         =   "Save as ISO..."
      End
      Begin VB.Menu mnuBurnISO 
         Caption         =   "Burn ISO image..."
      End
      Begin VB.Menu mnuS2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOTF 
         Caption         =   "On The Fly"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEjectDisk 
         Caption         =   "Eject Disk"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFinalize 
         Caption         =   "Finalize Disk"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCloseSession 
         Caption         =   "Close session"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTestMode 
         Caption         =   "Test Mode"
      End
      Begin VB.Menu mnuS3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuISODescriptors 
         Caption         =   "ISO descriptors..."
      End
      Begin VB.Menu mnuS7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelTempDir 
         Caption         =   "Select temp dir..."
      End
      Begin VB.Menu mnuS8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseLastSession 
         Caption         =   "Finalize disc..."
      End
   End
   Begin VB.Menu mnuBurnDisc 
      Caption         =   "Burn Disc!"
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cManager            As New FL_Manager
Private cDrvInfo            As New FL_DriveInfo
Private cCDInfo             As New FL_CDInfo

Private WithEvents cISO     As FL_ISO9660Writer
Attribute cISO.VB_VarHelpID = -1
Private WithEvents cISOCD   As FL_CDISOWriter
Attribute cISOCD.VB_VarHelpID = -1
Private WithEvents cDataCD  As FL_CDDataWriter
Attribute cDataCD.VB_VarHelpID = -1

Private strDrvID            As String

' Some properties for to be accessed from
' other objects in the project
Public Property Get VolumeID() As String
    VolumeID = cISO.VolumeID
End Property

Public Property Let VolumeID(val As String)
    cISO.VolumeID = val
    Me.Caption = "Flamed v4 Data CD Writer - " & cISO.VolumeID
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

' drive selected
' get its drive ID and show its writing speeds
' >> You may refresh the writing speed
' >> after a new medium arrived in the drive,
' >> as the supported speeds may depend on the medium.
Private Sub cboDrvs_Click()
    strDrvID = cManager.DrvChr2DrvID(Left$(cboDrvs.List(cboDrvs.ListIndex), 1))
    cCDInfo.GetInfo strDrvID
    ShowStat
    ListSpeeds
End Sub

' FL_CDDataWriter first checks for the
' availability of files contained in the
' passed compilation to avoid errors
' during write process.
Private Sub cDataCD_CheckForFiles()
    sbar.SimpleText = "Checking for availability of files..."
End Sub

' Some files are missing, show them.
Private Sub cDataCD_FilesMissing(strFiles() As String)

    Dim i       As Integer
    Dim strBuf  As String

    sbar.SimpleText = "Files are missing!"

    For i = LBound(strFiles) To UBound(strFiles)
        strBuf = strBuf & strFiles(i) & vbCrLf
    Next

    MsgBox "Files are missing:" & vbCrLf & _
           strBuf, vbExclamation, "Writing canceled"

End Sub

Private Sub cDataCD_ClosingSession()
    sbar.SimpleText = "Closing session..."
End Sub

Private Sub cDataCD_Finished()
    sbar.SimpleText = "Finished"
End Sub

' building temporary ISO image...
' this will also fire for on the fly writing,
' as the ISO header first gets saved to HDD.
' But don't worry, in most cases it's
' less then half a meg big.
Private Sub cDataCD_ISOProgress(ByVal lngMax As Long, ByVal lngValue As Long)
    On Error Resume Next
    If Not prg.Max = lngMax Then prg.Max = lngMax
    prg.Value = lngValue
    sbar.SimpleText = "Creating temporary image... " & Format(CLng(lngValue / lngMax * 100), "00") & "%"
End Sub

Private Sub cDataCD_StartWriting()
    sbar.SimpleText = "Starting writing..."
End Sub

Private Sub cDataCD_WriteProgress(ByVal percent As Integer)
    ' NO TIME WAISTING STUFF HERE!!!
    ' EVENT GETS DIRECTLY FIRED FROM
    ' THE WRITING FUNCTION!
    On Error Resume Next

    If Not prg.Max = 100 Then prg.Max = 100
    prg.Value = percent
    sbar.SimpleText = "Writing... " & Format(percent, "00") & "%"
End Sub

Private Sub cISO_Progress(ByVal lngMax As Long, ByVal lngValue As Long)
    On Error Resume Next

    If Not prg.Max = lngMax Then prg.Max = lngMax
    prg.Value = lngValue
    sbar.SimpleText = "Saving ISO image... " & Format(CInt(lngValue / lngMax * 100), "00") & "%"
End Sub

Private Sub cISOCD_ClosingSession()
    sbar.SimpleText = "Closing session..."
End Sub

Private Sub cISOCD_Finished()
    sbar.SimpleText = "Finished."
End Sub

Private Sub cISOCD_Progress(percent As Integer)
    ' NO TIME WAISTING STUFF HERE!!!
    ' EVENT GETS DIRECTLY FIRED FROM
    ' THE WRITING FUNCTION!
    On Error Resume Next
    If Not prg.Max = 100 Then prg.Max = 100
    prg.Value = percent
    sbar.SimpleText = "Writing... " & Format(percent, "00") & "%"
End Sub

Private Sub cISOCD_StartWriting()
    sbar.SimpleText = "Starting writing..."
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    lstDirs.Width = Me.ScaleWidth * 1 / 3
    lstFiles.Width = Me.ScaleWidth * 2 / 3 - lstDirs.Left - 10
    lstFiles.Left = lstDirs.Left + lstDirs.Width + 5

    lstDirs.Height = Me.ScaleHeight - frmPrg.Height - sbar.Height - lstDirs.Top - 8
    lstFiles.Height = lstDirs.Height

    frmPrg.Top = lstDirs.Top + lstDirs.Height + 3
    frmPrg.Width = Me.ScaleWidth - frmPrg.Left * 2

    prg.Top = frmPrg.Top + 19
    prg.Width = Me.ScaleWidth - prg.Left * 2

    cboDrvs.Width = (Me.ScaleWidth * 6 / 8 - 8) * Screen.TwipsPerPixelX
    cboSpeed.Width = (Me.ScaleWidth * 2 / 8 - 6) * Screen.TwipsPerPixelX
    cboSpeed.Left = cboDrvs.Left + cboDrvs.Width + 4 * Screen.TwipsPerPixelX

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

Private Sub lstDirs_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i   As Long

    'add dropped dirs
    For i = 1 To Data.Files.Count
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

    ShowStat

End Sub


Private Sub lstFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i   As Long

    'add dropped dirs
    For i = 1 To Data.Files.Count
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

    ShowStat

End Sub

Private Sub lstFiles_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete: mnuRemFiles_Click
    End Select
End Sub

Private Sub lstFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuRClickF
End Sub

Private Sub PrepareWrite()

    cISOCD.EjectAfterWrite = mnuEjectDisk.Checked
    cDataCD.EjectAfterWrite = mnuEjectDisk.Checked
    cISOCD.NextSessionAllowed = Not mnuFinalize.Checked
    cDataCD.NextSessionAllowed = Not mnuFinalize.Checked
    cISOCD.TestMode = mnuTestMode.Checked
    cDataCD.TestMode = mnuTestMode.Checked
    cDataCD.OnTheFly = mnuOTF.Checked
    cDataCD.CloseSession = mnuCloseSession.Checked

    If Not cManager.SetCDRomSpeed(strDrvID, &HFFFF&, cboSpeed.ItemData(cboSpeed.ListIndex)) Then
        MsgBox "Could not set write speed.", vbExclamation, "Error"
    End If

End Sub

Private Sub mnuBurnDisc_Click()

    Dim strRes  As String

    ' prepare
    Set cDataCD.ISOClass = cISO
    PrepareWrite

    ' do the job
    Select Case cDataCD.WriteISOtoCD(strDrvID)

        Case BURNRET_CLOSE_SESSION: strRes = "Closing session failed"
        Case BURNRET_CLOSE_TRACK: strRes = "Closing track failed"
        Case BURNRET_FILE_ACCESS: strRes = "File access error"
        Case BURNRET_INVALID_MEDIA: strRes = "Invalid medium"
        Case BURNRET_ISOCREATION: strRes = "ISO creation failed"
        Case BURNRET_NOT_EMPTY: strRes = "Disk isn't empty"
        Case BURNRET_OK: strRes = "Finished"
        Case BURNRET_SYNC_CACHE: strRes = "Synchronizing cache failed"
        Case BURNRET_WPMP: strRes = "Invalid write parameters mode page"
        Case BURNRET_WRITE: strRes = "Write error (buffer underrun?)"

    End Select

    MsgBox "Result:" & vbCrLf & strRes, vbInformation, "Finished"

End Sub

Private Sub mnuBurnISO_Click()

    On Error GoTo ErrorHandler

    Dim strRes  As String

    dlgISO.FileName = vbNullString
    dlgISO.ShowOpen

    cISOCD.ISOFile = dlgISO.FileName

    PrepareWrite

    Select Case cISOCD.WriteISOtoCD(strDrvID)

        Case BURNRET_CLOSE_SESSION: strRes = "Closing session failed"
        Case BURNRET_CLOSE_TRACK: strRes = "Closing track failed"
        Case BURNRET_FILE_ACCESS: strRes = "File access error"
        Case BURNRET_INVALID_MEDIA: strRes = "Invalid medium"
        Case BURNRET_ISOCREATION: strRes = "ISO creation failed"
        Case BURNRET_NOT_EMPTY: strRes = "Disk isn't empty"
        Case BURNRET_OK: strRes = "Finished"
        Case BURNRET_SYNC_CACHE: strRes = "Synchronizing cache failed"
        Case BURNRET_WPMP: strRes = "Invalid write parameters mode page"
        Case BURNRET_WRITE: strRes = "Write error (buffer underrun?)"

    End Select

    MsgBox "Result:" & vbCrLf & strRes, vbInformation, "Finished"

ErrorHandler:

End Sub

Private Sub mnuClear_Click()

    cISO.ClearDirsFiles
    lstFiles.ListItems.Clear
    AddRootNode

    VolumeID = "NEW_VOLUME"
    SystemID = "WIN32"
    PublisherID = "FLAMEDv4"
    AppID = "FLAMEDv4 DATA WRITER"

End Sub

Private Sub AddRootNode()
    lstDirs.Nodes.Clear
    lstDirs.Nodes.Add(, , "\", "\ [" & VolumeID & "]", 2).Selected = True
End Sub

Private Sub mnuCloseLastSession_Click()

    If MsgBox("App will freeze for some seconds." & vbCrLf & _
           "Continue?", vbQuestion Or vbYesNo, "Continue?") = vbNo Then
        Exit Sub
    End If

    If Not cDataCD.CloseLastSession(strDrvID) Then
        MsgBox "Failed.", vbExclamation
    Else
        MsgBox "Finished.", vbInformation
    End If

End Sub

Private Sub mnuCloseSession_Click()
    mnuCloseSession.Checked = Not mnuCloseSession.Checked
End Sub

Private Sub mnuOTF_Click()
    mnuOTF.Checked = Not mnuOTF.Checked
End Sub

Private Sub mnuEjectDisk_Click()
    mnuEjectDisk.Checked = Not mnuEjectDisk.Checked
End Sub

Private Sub mnuFinalize_Click()
    mnuFinalize.Checked = Not mnuFinalize.Checked
End Sub

Private Sub mnuTestMode_Click()
    mnuTestMode.Checked = Not mnuTestMode.Checked
End Sub

Private Sub mnuISODescriptors_Click()
    frmISODescriptors.Show vbModal, Me
End Sub

Private Sub mnuLoadPrj_Click()

    On Error GoTo ErrorHandler

    dlgFLA.FileName = vbNullString
    dlgFLA.ShowOpen

    If Not cISO.LoadProject(dlgFLA.FileName) Then
        MsgBox "Failed to load project.", vbExclamation, "Error"
    End If

    ' show directories
    lstFiles.ListItems.Clear
    AddRootNode

    ListDir "\"

    ' make sure files in root are shown
    lstDirs_Click

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

End Sub

Private Sub mnuRemFiles_Click()

    Dim i   As Integer

    ' remove selected file
    With lstFiles.ListItems

        For i = .Count To 1 Step -1
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

End Sub

Private Sub mnuSaveISO_Click()

    On Error GoTo ErrorHandler

    dlgISO.FileName = vbNullString
    dlgISO.ShowSave

    If Not cISO.CreateISO(dlgISO.FileName) Then
        MsgBox "Failed to save ISO image.", vbExclamation, "Error"
    Else
        MsgBox "Finished.", vbInformation, "Ok"
    End If

ErrorHandler:

End Sub

Private Sub mnuSavePrj_Click()

    On Error GoTo ErrorHandler

    dlgFLA.FileName = vbNullString
    dlgFLA.ShowSave

    If Not cISO.SaveProject(dlgFLA.FileName) Then
        MsgBox "Failed to save project.", vbExclamation, "Error"
    End If

ErrorHandler:

End Sub

Private Sub mnuSelTempDir_Click()

    Dim strNewPath  As String

    strNewPath = BrowseForFolder("Please select a new temp:", _
                                 cDataCD.TempDir, _
                                 hWnd, True, , True)

    If Not strNewPath = "" Then
        cDataCD.TempDir = strNewPath
    End If

End Sub

Private Sub ListSpeeds()

    Dim i           As Integer
    Dim intSpeeds() As Integer

    ' get write speeds
    intSpeeds = cDrvInfo.GetWriteSpeeds(strDrvID)

    ' show them
    cboSpeed.Clear
    For i = 0 To UBound(intSpeeds)
        cboSpeed.AddItem intSpeeds(i) & " KB/s (" & intSpeeds(i) \ 176 & "x)"
        cboSpeed.ItemData(i) = intSpeeds(i)
    Next

    ' add descriptor for maximum speed
    cboSpeed.AddItem "Max."
    cboSpeed.ItemData(i) = &HFFFF&

    cboSpeed.ListIndex = cboSpeed.ListCount - 1

End Sub

Private Sub ListCDRs()

    Dim i           As Integer
    Dim strDrvs()   As String

    ' get all CD/DVD drives
    strDrvs = cManager.GetCDVDROMs

    ' check if they can write to CD-R(W)
    For i = 0 To UBound(strDrvs) - 1
        If IsCDRWriter(strDrvs(i)) Then

            ' found one, add it to the list
            cboDrvs.AddItem strDrvs(i) & ": " & _
                            cDrvInfo.Vendor & " " & _
                            cDrvInfo.Product & " " & _
                            cDrvInfo.Revision

        End If
    Next

    ' found at least one drive?
    If cboDrvs.ListCount > 0 Then
        cboDrvs.ListIndex = 0
    Else
        MsgBox "No CD writers found.", vbExclamation, "Error"
    End If

End Sub

Private Function IsCDRWriter(char As String) As Boolean

    strDrvID = cManager.DrvChr2DrvID(char)

    If Not cDrvInfo.GetInfo(strDrvID) Then
        Exit Function
    End If

    ' drive has CD-R(W) write capability?
    IsCDRWriter = (cDrvInfo.WriteCapabilities And WC_CDR) Or _
                  (cDrvInfo.WriteCapabilities And WC_CDRW)

End Function

Private Sub Form_Load()

    Set cDataCD = New FL_CDDataWriter
    Set cISOCD = New FL_CDISOWriter
    Set cISO = New FL_ISO9660Writer

    If Not cManager.Init() Then
        MsgBox "No interfaces found.", vbExclamation, "Error"
        Unload Me
    End If

    ' show CD writers
    ListCDRs

    ' show root
    AddRootNode

    VolumeID = "NEW_VOLUME"
    SystemID = "WIN32"
    PublisherID = "FLAMEDv4"
    AppID = "FLAMEDv4 DATA WRITER"

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cManager.Goodbye
End Sub

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

Private Function AddSlash(ByVal sVal As String) As String
    If Not Right$(sVal, 1) = "\" Then sVal = sVal & "\"
    AddSlash = sVal
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

Public Function FileExists(ByVal Path As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(Path) And (vbDirectory Or vbVolume)) = 0
End Function

Public Function DirExists(ByVal Path As String) As Boolean
    On Error Resume Next
    DirExists = CBool(GetAttr(Path) And vbDirectory)
End Function

Private Sub ShowStat()
    sbar.SimpleText = "Used: " & FormatFileSize(cISO.ISOSize) & " / " & FormatFileSize(cCDInfo.Capacity)
End Sub
