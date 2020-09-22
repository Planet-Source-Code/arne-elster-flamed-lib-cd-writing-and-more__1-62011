VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmDriveInfo 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Drive information"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6435
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
   MinButton       =   0   'False
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox picDrvNfo 
      BorderStyle     =   0  'Kein
      Height          =   3540
      Left            =   300
      ScaleHeight     =   236
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   4
      Top             =   1575
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ListView lstReadCaps 
         Height          =   1065
         Left            =   1350
         TabIndex        =   6
         Top             =   75
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   1879
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView lstWriteCaps 
         Height          =   1065
         Left            =   1350
         TabIndex        =   8
         Top             =   1275
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   1879
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ListView lstMisc 
         Height          =   1065
         Left            =   1350
         TabIndex        =   10
         Top             =   2475
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   1879
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label lblMisc 
         AutoSize        =   -1  'True
         Caption         =   "miscellaneous:"
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   2475
         Width           =   1035
      End
      Begin VB.Label lblWriteCaps 
         AutoSize        =   -1  'True
         Caption         =   "Write capabilities:"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   1275
         Width           =   1275
      End
      Begin VB.Label lblReadCaps 
         AutoSize        =   -1  'True
         Caption         =   "Read capabilities:"
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   75
         Width           =   1260
      End
   End
   Begin VB.PictureBox picCDNfo 
      BorderStyle     =   0  'Kein
      Height          =   3450
      Left            =   300
      ScaleHeight     =   230
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   11
      Top             =   1650
      Width           =   5940
      Begin VB.CommandButton cmdLock 
         Caption         =   "Lock/Unlock"
         Height          =   315
         Left            =   4215
         TabIndex        =   26
         Top             =   3075
         Width           =   1290
      End
      Begin VB.CommandButton cmdEject 
         Caption         =   "Eject/Load"
         Height          =   315
         Left            =   2850
         TabIndex        =   25
         Top             =   3075
         Width           =   1290
      End
      Begin VB.CommandButton cmdBlank 
         Caption         =   "Blank CD-RW"
         Height          =   315
         Left            =   1350
         TabIndex        =   24
         Top             =   3075
         Width           =   1215
      End
      Begin VB.CommandButton cmdCloseSession 
         Caption         =   "Close session"
         Height          =   315
         Left            =   75
         TabIndex        =   23
         Top             =   3075
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstCDNfo 
         Height          =   3045
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   5371
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.PictureBox picSectors 
      BorderStyle     =   0  'Kein
      Height          =   3540
      Left            =   225
      ScaleHeight     =   236
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   15
      Top             =   1575
      Visible         =   0   'False
      Width           =   6015
      Begin MSComDlg.CommonDialog dlgDAT 
         Left            =   4350
         Top             =   675
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "DAT files (*.dat)|*.dat|All files (*.*)|*.*"
         Flags           =   2
      End
      Begin VB.CommandButton cmdSaveSector 
         Caption         =   "Save"
         Height          =   315
         Left            =   4800
         TabIndex        =   22
         Top             =   0
         Width           =   1140
      End
      Begin VB.ComboBox cboReadMode 
         Height          =   315
         ItemData        =   "frmDriveInfo.frx":0000
         Left            =   2505
         List            =   "frmDriveInfo.frx":000D
         Style           =   2  'Dropdown-Liste
         TabIndex        =   20
         Top             =   0
         Width           =   2145
      End
      Begin ComCtl2.UpDown udLBA 
         Height          =   315
         Left            =   1215
         TabIndex        =   19
         Top             =   0
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   327681
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtLBA"
         BuddyDispid     =   196615
         OrigLeft        =   90
         OrigRight       =   106
         OrigBottom      =   21
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtLBA 
         Height          =   285
         Left            =   450
         MaxLength       =   6
         TabIndex        =   18
         Text            =   "0"
         Top             =   15
         Width           =   765
      End
      Begin MSComctlLib.ListView lstHex 
         Height          =   3045
         Left            =   0
         TabIndex        =   16
         Top             =   375
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   5371
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Offset"
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "hexadecimal"
            Object.Width           =   5538
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ASCII"
            Object.Width           =   2716
         EndProperty
      End
      Begin VB.Label lblReadMode 
         AutoSize        =   -1  'True
         Caption         =   "Readmode:"
         Height          =   195
         Left            =   1575
         TabIndex        =   21
         Top             =   45
         Width           =   825
      End
      Begin VB.Label lblLBA 
         AutoSize        =   -1  'True
         Caption         =   "LBA:"
         Height          =   195
         Left            =   75
         TabIndex        =   17
         Top             =   75
         Width           =   330
      End
   End
   Begin MSComctlLib.ImageList imgContent 
      Left            =   5100
      Top             =   675
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
            Picture         =   "frmDriveInfo.frx":0028
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDriveInfo.frx":27DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDriveInfo.frx":2934
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picContent 
      BorderStyle     =   0  'Kein
      Height          =   3540
      Left            =   225
      ScaleHeight     =   236
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   13
      Top             =   1575
      Visible         =   0   'False
      Width           =   6015
      Begin MSComctlLib.TreeView tvwTracks 
         Height          =   3465
         Left            =   75
         TabIndex        =   14
         Top             =   75
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   6112
         _Version        =   393217
         Indentation     =   471
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgContent"
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList img 
      Left            =   5775
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDriveInfo.frx":50E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDriveInfo.frx":7898
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDriveInfo.frx":A04A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDriveInfo.frx":C7FC
            Key             =   ""
         EndProperty
      EndProperty
   End
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
      TabIndex        =   2
      Top             =   675
      Width           =   5340
   End
   Begin MSComctlLib.TabStrip tabstrip 
      Height          =   4140
      Left            =   75
      TabIndex        =   1
      Top             =   1125
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   7303
      Separators      =   -1  'True
      TabMinWidth     =   2587
      ImageList       =   "img"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Drive info"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Disk info"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Disk content"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sector Viewer"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "Drive:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   750
      Width           =   435
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drive information"
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
      Left            =   375
      TabIndex        =   0
      Top             =   120
      Width           =   2160
   End
   Begin VB.Shape shpHdr 
      BackColor       =   &H00336471&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   6690
   End
   Begin VB.Menu mnuRClick 
      Caption         =   "RClick"
      Visible         =   0   'False
      Begin VB.Menu mnuGoViewer 
         Caption         =   "GoTo sector viewer"
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpandAll 
         Caption         =   "Expand all"
      End
      Begin VB.Menu mnuCollapseAll 
         Caption         =   "Collapse all"
      End
   End
End
Attribute VB_Name = "frmDriveInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cMonitor As FL_DoorMonitor
Attribute cMonitor.VB_VarHelpID = -1

Private cDrvNfo         As New FL_DriveInfo
Private cCDNfo          As New FL_CDInfo
Private cTrkNfo         As New FL_TrackInfo
Private cSessNfo        As New FL_SessionInfo
Private cReader         As New FL_CDReader

Private strPrvDriveID   As String

Private btBuffer()      As Byte

Private Sub ShowDrives()

    Dim strDrives() As String
    Dim i           As Long

    strDrives = GetDriveList(OPT_ALL)

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

Private Sub cboDrv_Click()
    strPrvDriveID = cManager.DrvChr2DrvID(Left(cboDrv.List(cboDrv.ListIndex), 1))
    tabstrip_Click
End Sub

Private Sub cmdBlank_Click()
    frmBlankCDRW.DriveID = strPrvDriveID
    frmBlankCDRW.Show vbModal, Me
End Sub

Private Sub cmdCloseSession_Click()

    Dim cDataCD As New FL_CDDataWriter
    Dim blnFin  As Boolean

    If cCDNfo.LastSessionState <> STAT_INCOMPLETE Then
        MsgBox "Last session is either empty or already closed.", vbExclamation
        Exit Sub
    End If

    blnFin = MsgBox("Finalize?", vbQuestion Or vbYesNo) = vbYes

    MsgBox "Please wait till message appears!"

    If cDataCD.CloseLastSession(strPrvDriveID, blnFin) Then
        MsgBox "Finished", vbInformation
    Else
        MsgBox "Failed", vbExclamation
    End If

End Sub

Private Sub cmdEject_Click()

    ' can we determine the drive's state?
    If Not cDrvNfo.GetInfo(strPrvDriveID) Then

        ' no, just try to open it
        MsgBox "Could not read drive information." & vbCrLf & _
               "Will just try to open the drive.", vbExclamation, "error"

        If Not cManager.UnLoadDrive(strPrvDriveID) Then
            MsgBox "Failed to open the drive.", vbExclamation, "Error"
        End If

        Exit Sub

    End If

    ' drive closed?
    If cDrvNfo.DriveClosed Then
        ' open
        If Not cManager.UnLoadDrive(strPrvDriveID) Then
            MsgBox "Failed to open the drive.", vbExclamation, "Error"
        End If
    Else
        ' close
        If Not cManager.LoadDrive(strPrvDriveID) Then
            MsgBox "Failed to close the drive.", vbExclamation, "Error"
        End If
    End If

End Sub

Private Sub cmdLock_Click()

    ' can we determine the drive's state?
    If Not cDrvNfo.GetInfo(strPrvDriveID) Then

        ' no, just try to unlock the drive
        ' as we can't determine wether the
        ' drive is locked or not.
        ' Would be bad if we nverthless
        ' locked it and the method wouldn't
        ' allow to unlock it.
        MsgBox "Could not read drive information." & vbCrLf & _
               "Will just try to unlock the drive.", vbExclamation, "error"

        If Not cManager.UnLockDrive(strPrvDriveID) Then
            MsgBox "Failed to unlock the drive.", vbExclamation, "Error"
        End If

        Exit Sub

    End If

    If cDrvNfo.DriveLocked Then
        If Not cManager.UnLockDrive(strPrvDriveID) Then
            MsgBox "Failed to unlock the drive.", vbExclamation, "Error"
        End If
    Else
        If Not cManager.LockDrive(strPrvDriveID) Then
            MsgBox "Failed to lock the drive.", vbExclamation, "Error"
        End If
    End If

End Sub

Private Sub cmdSaveSector_Click()
    On Error GoTo ErrorHandler

    Dim FF  As Integer: FF = FreeFile

    dlgDAT.ShowSave
    Open dlgDAT.FileName For Binary As #FF
        Put #FF, , btBuffer
    Close #FF

    Exit Sub

ErrorHandler:
    MsgBox "Failed:" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub cMonitor_arrival(ByVal drive As String)
    If LCase$(drive) = LCase$(cManager.DrvID2DrvChr(strPrvDriveID)) Then
        cboDrv_Click
    End If
End Sub

Private Sub Form_Load()
    Set cMonitor = New FL_DoorMonitor
    cMonitor.InitDoorMonitor
    ShowDrives
    cboReadMode.ListIndex = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cMonitor.DeInitDoorMonitor
End Sub

Private Sub mnuCollapseAll_Click()
    Dim n   As Node

    For Each n In tvwTracks.Nodes
        n.Expanded = False
    Next
End Sub

Private Sub mnuExpandAll_Click()
    Dim n   As Node

    For Each n In tvwTracks.Nodes
        n.Expanded = True
    Next
End Sub

Private Sub mnuGoViewer_Click()
    If tvwTracks.SelectedItem.Tag = vbNullString Then
        Exit Sub
    End If
    tabstrip.Tabs(4).Selected = True
    txtLBA = tvwTracks.SelectedItem.Tag
    ShowSector txtLBA
End Sub

Private Sub tabstrip_Click()
    Select Case tabstrip.SelectedItem.index
        Case 1
            picCDNfo.Visible = False
            picContent.Visible = False
            picDrvNfo.Visible = True
            picSectors.Visible = False
            DriveInfo
        Case 2
            picCDNfo.Visible = True
            picDrvNfo.Visible = False
            picContent.Visible = False
            picSectors.Visible = False
            CDInfo
        Case 3
            picCDNfo.Visible = False
            picDrvNfo.Visible = False
            picContent.Visible = True
            picSectors.Visible = False
            CDContentInfo
        Case 4
            picCDNfo.Visible = False
            picDrvNfo.Visible = False
            picContent.Visible = False
            picSectors.Visible = True
            txtLBA.Text = 0
            ShowSector 0
    End Select
End Sub

Private Sub CDContentInfo()

    On Error GoTo ErrorHandler

    Dim i   As Integer, j   As Integer

    tvwTracks.Nodes.Clear

    If Not cCDNfo.GetInfo(strPrvDriveID) Then
        MsgBox "Couldn't read CD information", vbExclamation, "Error"
        Exit Sub
    End If

    ' show sessions
    For i = 1 To cCDNfo.Sessions

        If Not cSessNfo.GetInfo(strPrvDriveID, i) Then
            MsgBox "Couldn't read information about session " & i, vbExclamation, "Error"
            ' don't want to have complex structures here
            ' keep it as simple as possible
            GoTo SkipSession
        End If

        ' add session node
        tvwTracks.Nodes.Add(, , "s" & Format(i, "00"), "Session " & Format(i, "00"), 3).Expanded = True

        ' show tracks for the current session
        For j = cSessNfo.FirstTrack To cSessNfo.LastTrack

            ' add track node
            With tvwTracks.Nodes.Add("s" & Format(i, "00"), tvwChild, "t" & Format(j, "00"), "Track " & Format(j, "00"))

                If Not cTrkNfo.GetInfo(strPrvDriveID, j) Then
                    MsgBox "Couldn't read information about track " & j, vbExclamation, "Error"
                    ' don't want to have complex structures here
                    ' keep it as simple as possible
                    GoTo SkipTrack
                End If

                .Image = Abs(Not CBool(cTrkNfo.Mode = MODE_AUDIO)) + 1
                .Expanded = True
                .Tag = cTrkNfo.TrackStart.LBA

            End With

            ' track information
            tvwTracks.Nodes.Add("t" & Format(j, "00"), tvwChild, , "Track start: " & cTrkNfo.TrackStart.MSF & " MSF (" & cTrkNfo.TrackStart.LBA & " LBA)").Tag = cTrkNfo.TrackStart.LBA
            tvwTracks.Nodes.Add("t" & Format(j, "00"), tvwChild, , "Track length: " & cTrkNfo.TrackLength.MSF & " MSF (" & cTrkNfo.TrackLength.LBA & " LBA)").Tag = cTrkNfo.TrackLength.LBA
            tvwTracks.Nodes.Add("t" & Format(j, "00"), tvwChild, , "Track end: " & cTrkNfo.TrackEnd.MSF & " MSF (" & cTrkNfo.TrackEnd.LBA & " LBA)").Tag = cTrkNfo.TrackEnd.LBA
            tvwTracks.Nodes.Add "t" & Format(j, "00"), tvwChild, , "Mode: " & TrackMode2Str(cTrkNfo.Mode)

SkipTrack:
        Next

SkipSession:
    Next

ErrorHandler:

End Sub

Private Function TrackMode2Str(m As FL_TrackModes) As String
    Select Case m
        Case MODE_AUDIO: TrackMode2Str = "audio"
        Case MODE_MODE1: TrackMode2Str = "mode 1"
        Case MODE_MODE2: TrackMode2Str = "mode 2"
        Case MODE_MODE2_FORM1: TrackMode2Str = "mode 2 form 1"
        Case MODE_MODE2_FORM2: TrackMode2Str = "mode 2 form 2"
    End Select
End Function

Private Sub CDInfo()

    lstCDNfo.ListItems.Clear

    If Not cCDNfo.GetInfo(strPrvDriveID) Then
        MsgBox "Could not get disk info.", vbExclamation
        Exit Sub
    End If

    With cCDNfo

        AddItem lstCDNfo, "Type", CDTypeToStr(.MediaType)

        If .MediaType = ROMTYPE_CDR Or _
           .MediaType = ROMTYPE_CDRW Or _
           .MediaType = ROMTYPE_CDROM Or _
           .MediaType = ROMTYPE_CDROM_R_RW Then

            AddItem lstCDNfo, "CD-ROM/R/RW type", STypeToStr(.CDRWType)
            AddItem lstCDNfo, "Vendor", .CDRWVendor

        End If

        AddItem lstCDNfo, "Size", (.Size \ 1024 ^ 2) & " MB"
        AddItem lstCDNfo, "Capacity", (.Capacity \ 1024 ^ 2) & " MB"
        AddItem lstCDNfo, "Erasable", .Erasable
        AddItem lstCDNfo, "Last session's state", Status2Str(.LastSessionState)
        AddItem lstCDNfo, "Lead-In", .LeadInMSF.MSF & " MSF (" & .LeadInMSF.LBA & " LBA)"
        AddItem lstCDNfo, "Last possible Lead-Out start", .LeadOutMSF.MSF & " MSF (" & .LeadOutMSF.LBA & " LBA)"
        AddItem lstCDNfo, "Media status", Status2Str(.MediaStatus)
        AddItem lstCDNfo, "Sessions", .Sessions
        AddItem lstCDNfo, "Tracks", .Tracks

    End With

End Sub

Private Sub DriveInfo()

    lstReadCaps.ListItems.Clear
    lstWriteCaps.ListItems.Clear
    lstMisc.ListItems.Clear

    If Not cDrvNfo.GetInfo(strPrvDriveID) Then
        MsgBox "Could not get drive info.", vbExclamation
        Exit Sub
    End If

    With cDrvNfo
        AddItem lstReadCaps, "Max. read speed", .ReadSpeedMax & " KB/s"
        AddItem lstReadCaps, "Cur. read speed", .ReadSpeedCur & " KB/s"
        AddItem lstReadCaps, "CD-R", CBool(.ReadCapabilities And RC_CDR)
        AddItem lstReadCaps, "CD-RW", CBool(.ReadCapabilities And RC_CDRW)
        AddItem lstReadCaps, "DVD-ROM", CBool(.ReadCapabilities And RC_DVDROM)
        AddItem lstReadCaps, "DVD-RAM", CBool(.ReadCapabilities And RC_DVDRAM)
        AddItem lstReadCaps, "DVD-R", CBool(.ReadCapabilities And RC_DVDR)
        AddItem lstReadCaps, "DVD+R", CBool(.ReadCapabilities And RC_DVDPR)
        AddItem lstReadCaps, "DVD+R DL", CBool(.ReadCapabilities And RC_DVDPRDL)
        AddItem lstReadCaps, "DVD-RW", CBool(.ReadCapabilities And RC_DVDR)
        AddItem lstReadCaps, "DVD+RW", CBool(.ReadCapabilities And RC_DVDPRW)
        AddItem lstReadCaps, "C2 Errors", CBool(.ReadCapabilities And RC_C2)
        AddItem lstReadCaps, "Bar Code", CBool(.ReadCapabilities And RC_BARCODE)
        AddItem lstReadCaps, "CDDA raw", CBool(.ReadCapabilities And RC_CDDARAW)
        AddItem lstReadCaps, "CD-Text", CBool(.ReadCapabilities And RC_CDTEXT)
        AddItem lstReadCaps, "ISRC", CBool(.ReadCapabilities And RC_ISRC)
        AddItem lstReadCaps, "Mode 2 Form 1", CBool(.ReadCapabilities And RC_MODE2FORM1)
        AddItem lstReadCaps, "Mode 2 Form 2", CBool(.ReadCapabilities And RC_MODE2FORM2)
        AddItem lstReadCaps, "Mount Rainer", CBool(.ReadCapabilities And RC_MRW)
        AddItem lstReadCaps, "Multisession", CBool(.ReadCapabilities And RC_MULTISESSION)
        AddItem lstReadCaps, "Sub-Channels", CBool(.ReadCapabilities And RC_SUBCHANNELS)
        AddItem lstReadCaps, "Sub-Channels corrected", CBool(.ReadCapabilities And RC_SUBCHANNELS_CORRECTED)
        AddItem lstReadCaps, "Sub-Channels from Lead-In", CBool(.ReadCapabilities And RC_SUBCHANNELS_FROM_LEADIN)

        AddItem lstWriteCaps, "Max. write speed", .WriteSpeedMax & " KB/s"
        AddItem lstWriteCaps, "Cur. write speed", .WriteSpeedCur & " KB/s"
        AddItem lstWriteCaps, "CD-R", CBool(.WriteCapabilities And WC_CDR)
        AddItem lstWriteCaps, "CD-RW", CBool(.WriteCapabilities And WC_CDRW)
        AddItem lstWriteCaps, "DVD-R", CBool(.WriteCapabilities And WC_DVDR)
        AddItem lstWriteCaps, "DVD+R", CBool(.WriteCapabilities And WC_DVDPR)
        AddItem lstWriteCaps, "DVD+R DL", CBool(.WriteCapabilities And WC_DVDPRDL)
        AddItem lstWriteCaps, "DVD-RW", CBool(.WriteCapabilities And WC_DVDRRW)
        AddItem lstWriteCaps, "DVD+RW", CBool(.WriteCapabilities And WC_DVDPRW)
        AddItem lstWriteCaps, "DVD-RAM", CBool(.WriteCapabilities And WC_DVDRAM)
        AddItem lstWriteCaps, "Mount Rainer", CBool(.WriteCapabilities And WC_MRW)
        AddItem lstWriteCaps, "BURN-Proof", CBool(.WriteCapabilities And WC_BURNPROOF)
        AddItem lstWriteCaps, "Test-Mode", CBool(.WriteCapabilities And WC_TESTMODE)
        AddItem lstWriteCaps, "TAO", CBool(.WriteCapabilities And WC_TAO)
        AddItem lstWriteCaps, "TAO+Test", CBool(.WriteCapabilities And WC_TAO_TEST)
        AddItem lstWriteCaps, "SAO", CBool(.WriteCapabilities And WC_SAO)
        AddItem lstWriteCaps, "SAO+Test", CBool(.WriteCapabilities And WC_SAO_TEST)
        AddItem lstWriteCaps, "DAO/16", CBool(.WriteCapabilities And WC_RAW_16)
        AddItem lstWriteCaps, "DAO/16+Test", CBool(.WriteCapabilities And WC_RAW_16_TEST)
        AddItem lstWriteCaps, "DAO/96", CBool(.WriteCapabilities And WC_RAW_96)
        AddItem lstWriteCaps, "DAO/96+Test", CBool(.WriteCapabilities And WC_RAW_96_TEST)

        AddItem lstMisc, "Analog audio playback", .AnalogAudioPlayback
        AddItem lstMisc, "Buffer size", .BufferSizeKB & " KB"
        AddItem lstMisc, "Anti Jitter", .JitterEffectCorrection
        AddItem lstMisc, "Loading mechanism", LoadingMech2Str(.LoadingMechanism)
        AddItem lstMisc, "Lockable", .Lockable
        AddItem lstMisc, "Physical interface", IPh2Str(.PhysicalInterface)
        AddItem lstMisc, "Disk present", .DiscPresent
        AddItem lstMisc, "Drive closed", .DriveClosed
        AddItem lstMisc, "Drive locked", .DriveLocked
        AddItem lstMisc, "Idle timer", .IdleTimer100MS & " ms"
        AddItem lstMisc, "Spindown timer", .SpinDownTimerMS & " ms"
        AddItem lstMisc, "Standby timer", .StandbyTimer100MS & " ms"
    End With

End Sub

Private Sub AddItem(lst As ListView, text1 As String, text2 As String)
    With lst.ListItems.Add(Text:=text1)
        .SubItems(1) = text2
    End With
End Sub

Private Function IPh2Str(i As FL_PhysicalInterfaces) As String
    Select Case i
        Case IF_ATAPI: IPh2Str = "ATAPI"
        Case IF_IEEE: IPh2Str = "IEEE"
        Case IF_SCSI: IPh2Str = "SCSI"
        Case IF_UNKNWN: IPh2Str = "unknown"
        Case IF_USB: IPh2Str = "USB"
    End Select
End Function

Private Function LoadingMech2Str(mech As FL_LoadingMech) As String
    Select Case mech
        Case LOAD_CADDY: LoadingMech2Str = "Caddy"
        Case LOAD_CHANGER: LoadingMech2Str = "Changer"
        Case LOAD_POPUP: LoadingMech2Str = "Popup"
        Case LOAD_TRAY: LoadingMech2Str = "Tray"
        Case LOAD_UNKNWN: LoadingMech2Str = "Unknown"
    End Select
End Function

Private Function Status2Str(s As FL_Status) As String
    Select Case s
        Case STAT_COMPLETE: Status2Str = "complete"
        Case STAT_EMPTY: Status2Str = "empty"
        Case STAT_INCOMPLETE: Status2Str = "incomplete"
        Case STAT_UNKNWN: Status2Str = "unknown"
    End Select
End Function

Private Function STypeToStr(s As FL_CDSubType) As String
    Select Case s
        Case STYPE_CDI: STypeToStr = "CD-I"
        Case STYPE_CDROMDA: STypeToStr = "CD-ROM/CDDA"
        Case STYPE_UNKNWN: STypeToStr = "Unknown"
        Case STYPE_XA: STypeToStr = "CD-XA"
    End Select
End Function

Private Function CDTypeToStr(s As FL_CDType) As String
    Select Case s
        Case ROMTYPE_CDR: CDTypeToStr = "CD-R"
        Case ROMTYPE_CDROM: CDTypeToStr = "CD-ROM"
        Case ROMTYPE_CDROM_R_RW: CDTypeToStr = "CD-ROM/R/RW"
        Case ROMTYPE_CDRW: CDTypeToStr = "CD-RW"
        Case ROMTYPE_DVD_P_R: CDTypeToStr = "DVD+R"
        Case ROMTYPE_DVD_P_RW: CDTypeToStr = "DVD+RW"
        Case ROMTYPE_DVD_R: CDTypeToStr = "DVD-R"
        Case ROMTYPE_DVD_RAM: CDTypeToStr = "DVD-RAM"
        Case ROMTYPE_DVD_ROM: CDTypeToStr = "DVD-ROM"
        Case ROMTYPE_DVD_RW: CDTypeToStr = "DVD-RW"
    End Select
End Function

Private Sub tvwTracks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuRClick
End Sub

Private Sub txtLBA_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then ShowSector txtLBA
    KeyAscii = KeyAscii * Abs(IsNumeric(Chr$(KeyAscii)))
End Sub

Private Sub udLBA_Change()
    ShowSector txtLBA
End Sub

Private Sub cboReadMode_Click()

    Dim lngSize As Long
    Dim i       As Integer
    Dim blnRAW  As Boolean
    Dim blnCH   As Boolean

    blnRAW = cboReadMode.ListIndex > 0
    blnCH = cboReadMode.ListIndex = 2

    cCDNfo.GetInfo strPrvDriveID

    cTrkNfo.GetInfo strPrvDriveID, cCDNfo.Tracks
    udLBA.Max = cTrkNfo.TrackEnd.LBA

    For i = 1 To cCDNfo.Tracks
        cTrkNfo.GetInfo strPrvDriveID, i
        If CLng(txtLBA.Text) >= cTrkNfo.TrackStart.LBA And CLng(txtLBA.Text) <= cTrkNfo.TrackEnd.LBA Then
            lngSize = cReader.BufferSize(strPrvDriveID, 1, i, blnRAW, Abs(blnCH))
            Exit For
        End If
    Next

    If lngSize = 0 Then lngSize = 2352
    ReDim btBuffer(lngSize - 1) As Byte

End Sub

Private Sub ShowSector(LBA As Long)

    Dim i   As Integer
    Dim blnRAW  As Boolean
    Dim blnCH   As Boolean

    blnRAW = cboReadMode.ListIndex > 0
    blnCH = cboReadMode.ListIndex = 2

    If Not cReader.ReadSectorsLBA(strPrvDriveID, CLng(txtLBA), 1, btBuffer, blnRAW, Abs(blnCH)) Then
        MsgBox "Could not read sector " & txtLBA
        Exit Sub
    End If

    RenderBuffer

End Sub

Private Sub RenderBuffer()

    '     bytes per line
    Const BPL           As Integer = 13

    Dim i               As Integer
    Dim btPart(BPL - 1) As Byte
    Dim btRest()        As Byte

    lstHex.ListItems.Clear

    For i = LBound(btBuffer) To UBound(btBuffer) Step BPL

        If i + BPL > UBound(btBuffer) Then

            ReDim btRest(UBound(btBuffer) - i) As Byte
            CopyMemory btRest(0), btBuffer(i), UBound(btBuffer) - i

            With lstHex.ListItems
                With .Add
                    .Text = Format(i, "000000")
                    .SubItems(1) = ToHex(StrConv(btRest, vbUnicode))
                    .SubItems(2) = RemNulls(StrConv(btRest, vbUnicode))
                End With
            End With

            Exit For

        Else

            CopyMemory btPart(0), btBuffer(i), BPL

            With lstHex.ListItems
                With .Add
                    .Text = Format(i, "000000")
                    .SubItems(1) = ToHex(StrConv(btPart, vbUnicode))
                    .SubItems(2) = RemNulls(StrConv(btPart, vbUnicode))
                End With
            End With

        End If

    Next

End Sub

Private Function FormatEx(num As String) As String
    num = Format(num, "00")
    If Len(CStr(num)) = 1 Then FormatEx = "0" & num Else: FormatEx = num
End Function

Private Function ToHex(ByVal strText As String) As String

    Dim i       As Integer
    Dim strBuf  As String

    For i = 1 To Len(strText)
        ToHex = ToHex & FormatEx(Hex(Asc(Mid(strText, i, 1))))
    Next

End Function

Private Function RemNulls(ByVal strText As String) As String

    'supported chars
    Const ChrMap As String = " AaBbCcDdEeFfGgHhIiJjKkLlMmNn" & _
                             "OoPpQqRrSsTtUuVvWwXxYyZz0123" & _
                             "456789!""""§$%&/()=?,.-;:_<>|°"

    Dim sBuf As String, i As Long

    For i = 1 To Len(strText)
        If InStr(1, ChrMap, Mid$(strText, i, 1)) > 1 Then
            sBuf = sBuf & Mid$(strText, i, 1)
        ElseIf Mid$(strText, i, 1) = " " Then
            sBuf = sBuf & " "
        Else
            sBuf = sBuf & "."
        End If
    Next i

    RemNulls = sBuf

End Function
