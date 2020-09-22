VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSessToBIN 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Session to BIN/CUE"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
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
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   459
   StartUpPosition =   1  'Fenstermitte
   Begin MSComDlg.CommonDialog dlgCUE 
      Left            =   5482
      Top             =   1650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Cue sheets (*.cue)|*.cue"
      Flags           =   2
   End
   Begin VB.ComboBox cboSpeed 
      Height          =   315
      Left            =   5962
      Style           =   2  'Dropdown-Liste
      TabIndex        =   13
      ToolTipText     =   "Readspeed"
      Top             =   675
      Width           =   765
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
      Left            =   637
      Style           =   2  'Dropdown-Liste
      TabIndex        =   12
      Top             =   675
      Width           =   4665
   End
   Begin VB.Frame frmFile 
      Caption         =   "CUE file"
      Height          =   690
      Left            =   232
      TabIndex        =   8
      Top             =   3075
      Width           =   6540
      Begin VB.PictureBox picFile 
         BorderStyle     =   0  'Kein
         Height          =   465
         Left            =   75
         ScaleHeight     =   465
         ScaleWidth      =   6390
         TabIndex        =   9
         Top             =   180
         Width           =   6390
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   285
            Left            =   5850
            TabIndex        =   11
            Top             =   75
            Width           =   420
         End
         Begin VB.TextBox txtFile 
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   75
            Width           =   5565
         End
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   5197
      TabIndex        =   6
      Top             =   4725
      Width           =   1440
   End
   Begin VB.CommandButton cmdDrvNfo 
      Caption         =   "Drive information"
      Height          =   315
      Left            =   3622
      TabIndex        =   5
      Top             =   4725
      Width           =   1440
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "Back"
      Height          =   330
      Left            =   247
      TabIndex        =   4
      Top             =   4725
      Width           =   1365
   End
   Begin VB.Frame frmPrg 
      Caption         =   "Progress"
      Height          =   690
      Left            =   232
      TabIndex        =   1
      Top             =   3900
      Width           =   6540
      Begin VB.PictureBox picPrg 
         BorderStyle     =   0  'Kein
         Height          =   390
         Left            =   75
         ScaleHeight     =   390
         ScaleWidth      =   6390
         TabIndex        =   2
         Top             =   225
         Width           =   6390
         Begin MSComctlLib.ProgressBar prg 
            Height          =   315
            Left            =   90
            TabIndex        =   3
            Top             =   0
            Width           =   6240
            _ExtentX        =   11007
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   1
         End
      End
   End
   Begin MSComctlLib.ImageList img 
      Left            =   6082
      Top             =   1575
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSessToBIN.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSessToBIN.frx":27B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstSessions 
      Height          =   1740
      Left            =   112
      TabIndex        =   7
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3069
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "First track"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last Track"
         Object.Width           =   2672
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Size"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label lblSpeed 
      AutoSize        =   -1  'True
      Caption         =   "Speed:"
      Height          =   195
      Left            =   5362
      TabIndex        =   15
      Top             =   750
      Width           =   510
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "Drive:"
      Height          =   195
      Left            =   112
      TabIndex        =   14
      Top             =   750
      Width           =   435
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Session 2 BIN/CUE"
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
      Width           =   2340
   End
   Begin VB.Shape shpHdr 
      BackColor       =   &H00336471&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   6915
   End
End
Attribute VB_Name = "frmSessToBIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cGrab As FL_CueGrabber
Attribute cGrab.VB_VarHelpID = -1

Private cDrvNfo     As New FL_DriveInfo
Private cCDNfo      As New FL_CDInfo
Private cSessNfo    As New FL_SessionInfo

Private blnCancel   As Boolean

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

Private Sub ShowSpeeds()

    Dim i           As Integer
    Dim intSpeeds() As Integer

    cboSpeed.Clear

    intSpeeds = cDrvNfo.GetReadSpeeds(strDrvID)

    For i = LBound(intSpeeds) To UBound(intSpeeds)
        cboSpeed.AddItem (intSpeeds(i) \ 176) & "x"
        cboSpeed.ItemData(cboSpeed.ListCount - 1) = intSpeeds(i)
    Next

    cboSpeed.ListIndex = cboSpeed.ListCount - 1

End Sub

Private Sub ShowSessions()

    Dim i   As Integer

    lstSessions.ListItems.Clear

    If Not cCDNfo.GetInfo(strDrvID) Then
        MsgBox "Could not read disk.", vbExclamation
        Exit Sub
    End If

    For i = 1 To cCDNfo.Sessions

        If Not cSessNfo.GetInfo(strDrvID, i) Then
            MsgBox "Could not get info about track " & i, vbExclamation
        End If

        With lstSessions.ListItems
            With .Add(Text:=Format(i, "00"))
                .SubItems(1) = cSessNfo.FirstTrack
                .SubItems(2) = cSessNfo.LastTrack
                .SubItems(3) = ((cSessNfo.LeadOutLBA - cSessNfo.FirstTrackLBA) * 2048& \ 1024& ^ 2&) & " MB"
                .SmallIcon = 2
            End With
        End With

    Next

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

Private Sub cboDrv_Click()
    strDrvID = cManager.DrvChr2DrvID(Left$(cboDrv.List(cboDrv.ListIndex), 1))
    ShowSpeeds
    ShowSessions
End Sub

Private Sub cGrab_Progress(ByVal Percent As Integer, ByVal Track As Integer, ByVal startLBA As Long, ByVal endLBA As Long, Cancel As Boolean)
    prg.Value = Percent
    Cancel = blnCancel
    DoEvents
End Sub

Private Sub cmdBack_Click()
    Me.Hide
    frmImgTools.Show
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo ErrorHandler
    dlgCUE.ShowSave
    txtFile = dlgCUE.FileName
ErrorHandler:
End Sub

Private Sub cmdDrvNfo_Click()
    frmDriveInfo.Show vbModal, Me
End Sub

Private Sub cmdSave_Click()

    Dim strMsg  As String
    Dim strBIN  As String

    If cmdSave.Caption = "Cancel" Then
        blnCancel = True
        Exit Sub
    End If

    If txtFile = vbNullString Then
        MsgBox "No destination specified.", vbExclamation
        Exit Sub
    End If

    strBIN = FileFromPath(txtFile)
    strBIN = Left$(strBIN, Len(strBIN) - 3) & "bin"
    strBIN = PathFromPathFile(txtFile) & strBIN

    cManager.SetCDRomSpeed strDrvID, cboSpeed.ItemData(cboSpeed.ListIndex), &HFFFF&

    cmdBack.Enabled = Not cmdBack.Enabled
    cmdBrowse.Enabled = Not cmdBrowse.Enabled
    cmdDrvNfo.Enabled = Not cmdDrvNfo.Enabled
    cboDrv.Enabled = Not cboDrv.Enabled
    cboSpeed.Enabled = Not cboSpeed.Enabled
    
    blnCancel = False
    cmdSave.Caption = "Cancel"

    Select Case cGrab.SessionToBinCue(strDrvID, lstSessions.SelectedItem.index, strBIN, txtFile)
        Case ST_CANCELED: strMsg = "Canceled"
        Case ST_ENCODER_INIT: strMsg = "Could not init encoder."
        Case ST_FINISHED: strMsg = "Finished"
        Case ST_INVALID_SESSION: strMsg = "Invalid session."
        Case ST_INVALID_TRACKMODE: strMsg = "Track has invalid mode."
        Case ST_INVALID_TRACKNO: strMsg = "Invalid track number"
        Case ST_NOT_READY: strMsg = "Drive not ready."
        Case ST_READ_ERR: strMsg = "Read error."
        Case ST_UNKNOWN_ERR: strMsg = "Unknown error occured."
        Case ST_WRITE_ERR: strMsg = "Write error (HDD full?)"
    End Select

    MsgBox strMsg, vbInformation

    cmdBack.Enabled = Not cmdBack.Enabled
    cmdBrowse.Enabled = Not cmdBrowse.Enabled
    cmdDrvNfo.Enabled = Not cmdDrvNfo.Enabled
    cboDrv.Enabled = Not cboDrv.Enabled
    cboSpeed.Enabled = Not cboSpeed.Enabled

    cmdSave.Caption = "Save"

End Sub

Private Sub Form_Load()
    Set cGrab = New FL_CueGrabber
    ShowDrives
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmImgTools.Show
End Sub
