VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Audio CD Writer"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picButtons 
      BorderStyle     =   0  'Kein
      Height          =   465
      Left            =   150
      ScaleHeight     =   465
      ScaleWidth      =   5940
      TabIndex        =   10
      Top             =   3900
      Width           =   5940
      Begin VB.CommandButton cmdMoveUp 
         Caption         =   "Move up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   15
         Top             =   75
         Width           =   1065
      End
      Begin VB.CommandButton cmdMoveDown 
         Caption         =   "Move down"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3525
         TabIndex        =   14
         Top             =   75
         Width           =   1065
      End
      Begin VB.CommandButton cmdAddFile 
         Caption         =   "Add files"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   13
         Top             =   75
         Width           =   1065
      End
      Begin VB.CommandButton cmdRemFile 
         Caption         =   "Remove file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         TabIndex        =   12
         Top             =   75
         Width           =   1065
      End
      Begin VB.CommandButton cmdBurn 
         Caption         =   "Burn disk"
         Default         =   -1  'True
         Height          =   315
         Left            =   4875
         TabIndex        =   11
         Top             =   75
         Width           =   1065
      End
   End
   Begin VB.PictureBox picPrg 
      BorderStyle     =   0  'Kein
      Height          =   540
      Left            =   150
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   391
      TabIndex        =   5
      Top             =   4575
      Width           =   5865
      Begin MSComctlLib.ProgressBar prgTotal 
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   45
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar prgTrack 
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   300
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   75
         TabIndex        =   9
         Top             =   45
         Width           =   405
      End
      Begin VB.Label lblTrack 
         AutoSize        =   -1  'True
         Caption         =   "Track:"
         Height          =   195
         Left            =   75
         TabIndex        =   8
         Top             =   300
         Width           =   465
      End
   End
   Begin MSComctlLib.ImageList img 
      Left            =   4650
      Top             =   1725
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   5250
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Supported (*.wav;*.mp3)|*.wav;*.mp3|MPEG-3 audio (*.mp3)|*.mp3|PCM WAV (*.wav)|*.wav"
   End
   Begin VB.ComboBox cboSpeed 
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
      Left            =   4215
      Style           =   2  'Dropdown-Liste
      TabIndex        =   4
      Top             =   30
      Width           =   1890
   End
   Begin VB.ComboBox cboDrv 
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
      Left            =   90
      Style           =   2  'Dropdown-Liste
      TabIndex        =   3
      Top             =   30
      Width           =   4065
   End
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Unten ausrichten
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   5310
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   423
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Frame frmPrg 
      Caption         =   "Progress"
      Height          =   840
      Left            =   90
      TabIndex        =   1
      Top             =   4350
      Width           =   6015
   End
   Begin MSComctlLib.ListView lstTracks 
      Height          =   3465
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6112
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "track"
         Object.Width           =   1217
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "length"
         Object.Width           =   2116
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "local"
         Object.Width           =   6615
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cAudio   As FL_CDAudioWriter
Attribute cAudio.VB_VarHelpID = -1
Private cManager            As New FL_Manager
Private cDrvInfo            As New FL_DriveInfo
Private cCDInfo             As New FL_CDInfo

Private strDrvID            As String

Private Sub ListTracks()
    Dim i As Integer

    With lstTracks.ListItems
        .Clear

        For i = 0 To cAudio.FileCount - 1
            With .Add(Text:=" " & Format(i + 1, "00"), SmallIcon:=1)
                .SubItems(1) = FormatTime(cAudio.TrackLength(i)) & " min"
                .SubItems(2) = cAudio.file(i)
            End With
        Next

    End With
End Sub

Private Function FormatTime(ByVal sec As Long) As String
    FormatTime = Format(sec \ 60, "00") & ":" & _
                 Format(sec - (sec \ 60) * 60, "00")
End Function

Private Sub cAudio_CacheProgress(ByVal Percent As Integer, ByVal Track As Integer)
    On Error Resume Next
    prgTrack.Value = Percent
    DoEvents
End Sub

Private Sub cAudio_ClosingSession()
    sbar.SimpleText = "Closing session..."
End Sub

Private Sub cAudio_ClosingTrack(ByVal Track As Integer)
    sbar.SimpleText = "Closing track..."
End Sub

Private Sub cAudio_Finished()
    prgTotal.Value = prgTotal.Max
    sbar.SimpleText = "Finished."
End Sub

Private Sub cAudio_StartCaching()
    sbar.SimpleText = "Caching track..."
End Sub

Private Sub cAudio_StartWriting()
    sbar.SimpleText = "Writing track..."
End Sub

Private Sub cAudio_WriteProgress(ByVal Percent As Integer, ByVal Track As Integer)
    On Error Resume Next
    prgTrack.Value = Percent
    prgTotal.Value = Track
    DoEvents
End Sub

' drive selected
' get its drive ID and show its writing speeds
' >> You may refresh the writing speed
' >> after a new medium arrived in the drive,
' >> as the supported speeds may depend on the medium.
Private Sub cboDrv_Click()
    strDrvID = cManager.DrvChr2DrvID(Left$(cboDrv.List(cboDrv.ListIndex), 1))
    cCDInfo.GetInfo strDrvID
    ListSpeeds
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
            cboDrv.AddItem strDrvs(i) & ": " & _
                           cDrvInfo.Vendor & " " & _
                           cDrvInfo.Product & " " & _
                           cDrvInfo.Revision

        End If
    Next

    ' found at least one drive?
    If cboDrv.ListCount > 0 Then
        cboDrv.ListIndex = 0
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

Private Sub cmdAddFile_Click()

    Dim strFiles()  As String
    Dim strPath     As String
    Dim i           As Integer

    On Error GoTo ErrorHandler
    dlg.Flags = cdlOFNAllowMultiselect Or _
                cdlOFNExplorer Or _
                cdlOFNLongNames
    dlg.ShowOpen
    On Error GoTo 0

    ' multiple files selected?
    If InStr(dlg.FileName, Chr$(0)) > 0 Then

        strFiles = Split(dlg.FileName, Chr$(0))
        strPath = AddSlash(strFiles(0))
        For i = 1 To UBound(strFiles)
            strFiles(i - 1) = strPath & strFiles(i)
        Next
        ReDim Preserve strFiles(UBound(strFiles) - 1) As String

    ' one file selected
    Else

        ReDim strFiles(0) As String
        strFiles(0) = dlg.FileName

    End If

    ' add selected files
    For i = 0 To UBound(strFiles)

        If lstTracks.ListItems.Count = 99 Then
            MsgBox "Only 99 tracks allowed!", vbExclamation
            Exit For
        End If

        If Not cAudio.AddFile(strFiles(i)) Then
            MsgBox strFiles(i) & vbCrLf & " is not supported.", vbExclamation
        End If

    Next

    ListTracks

ErrorHandler:

End Sub

Private Sub cmdBurn_Click()

    prgTotal.Max = cAudio.FileCount

    cmdAddFile.Enabled = Not cmdAddFile.Enabled
    cmdBurn.Enabled = Not cmdBurn.Enabled
    cmdMoveDown.Enabled = Not cmdMoveDown.Enabled
    cmdMoveUp.Enabled = Not cmdMoveUp.Enabled
    cmdRemFile.Enabled = Not cmdRemFile.Enabled
    cboDrv.Enabled = Not cboDrv.Enabled
    cboSpeed.Enabled = Not cboSpeed.Enabled

    If Not cManager.SetCDRomSpeed(strDrvID, &HFFFF&, cboSpeed.ItemData(cboSpeed.ListIndex)) Then
        If MsgBox("Could not set writing speed." & vbCrLf & "Continue?", vbYesNo Or vbQuestion) = vbNo Then Exit Sub
    End If

    Select Case cAudio.WriteAudioToCD(strDrvID)
        Case BURNRET_CLOSE_SESSION: MsgBox "Error closing session."
        Case BURNRET_CLOSE_TRACK: MsgBox "Error closing track."
        Case BURNRET_FILE_ACCESS: MsgBox "Could not access a file."
        Case BURNRET_INVALID_MEDIA: MsgBox "Invalid medium."
        Case BURNRET_NO_NEXT_WRITABLE_LBA: MsgBox "Could not get next writable address."
        Case BURNRET_NOT_EMPTY: MsgBox "Disk isn't empty!"
        Case BURNRET_OK: MsgBox "Finished."
        Case BURNRET_SYNC_CACHE: MsgBox "Synchronizing cache failed."
        Case BURNRET_WPMP: MsgBox "Could not send write parameters page."
        Case BURNRET_WRITE: MsgBox "Write error (Buffer Underrun?"
    End Select

    cmdAddFile.Enabled = Not cmdAddFile.Enabled
    cmdBurn.Enabled = Not cmdBurn.Enabled
    cmdMoveDown.Enabled = Not cmdMoveDown.Enabled
    cmdMoveUp.Enabled = Not cmdMoveUp.Enabled
    cmdRemFile.Enabled = Not cmdRemFile.Enabled
    cboDrv.Enabled = Not cboDrv.Enabled
    cboSpeed.Enabled = Not cboSpeed.Enabled

End Sub

Private Sub cmdMoveDown_Click()
    cAudio.MoveIndexDown lstTracks.SelectedItem.Index - 1
    ListTracks
End Sub

Private Sub cmdMoveUp_Click()
    cAudio.MoveIndexUp lstTracks.SelectedItem.Index - 1
    ListTracks
End Sub

Private Sub cmdRemFile_Click()
    On Error Resume Next
    cAudio.RemFile lstTracks.SelectedItem.Index - 1
    ListTracks
End Sub

Private Sub Form_Load()

    Set cAudio = New FL_CDAudioWriter
    cAudio.EjectAfterWrite = True

    If Not cManager.Init() Then
        MsgBox "No interfaces found.", vbExclamation
        Unload Me
    End If

    ' show CD writers
    ListCDRs

End Sub

Private Sub Form_Resize()

    ' top
    cboSpeed.Width = Me.ScaleWidth * 3 / 8 - cboDrv.Left
    cboDrv.Width = Me.ScaleWidth * 5 / 8
    cboSpeed.Left = cboDrv.Left + cboDrv.Width + 6
    cboSpeed.Width = cboSpeed.Width - cboDrv.Left * 2
    lstTracks.Width = Me.ScaleWidth - lstTracks.Left * 2
    lstTracks.Height = Me.ScaleHeight - frmPrg.Height - cboDrv.Height - picButtons.Height - 29

    ' bottom
    frmPrg.Width = Me.ScaleWidth - frmPrg.Left * 2
    picPrg.Width = Me.ScaleWidth - picPrg.Left * 2
    frmPrg.Top = lstTracks.Top + lstTracks.Height + picButtons.Height
    picPrg.Top = frmPrg.Top + 13
    picButtons.Top = lstTracks.Top + lstTracks.Height

End Sub

Private Sub picPrg_Resize()
    prgTotal.Width = picPrg.ScaleWidth - prgTotal.Left * 1.2
    prgTrack.Width = picPrg.ScaleWidth - prgTrack.Left * 1.2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cManager.Goodbye
End Sub

Private Function AddSlash(ByVal aval As String) As String
    If Not Right$(aval, 1) = "\" Then aval = aval & "\"
    AddSlash = aval
End Function
