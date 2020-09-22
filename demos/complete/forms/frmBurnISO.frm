VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBurnISO 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Burn ISO image"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
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
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "Back"
      Height          =   330
      Left            =   150
      TabIndex        =   15
      Top             =   3150
      Width           =   1365
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Burn ISO image"
      Default         =   -1  'True
      Height          =   330
      Left            =   4500
      TabIndex        =   14
      Top             =   3150
      Width           =   1365
   End
   Begin VB.CommandButton cmdDrvNfo 
      Caption         =   "Drive information"
      Height          =   315
      Left            =   2850
      TabIndex        =   13
      Top             =   3150
      Width           =   1440
   End
   Begin VB.Frame frmOptions 
      Caption         =   "Options"
      Height          =   1515
      Left            =   97
      TabIndex        =   5
      Top             =   1500
      Width           =   5865
      Begin VB.PictureBox picSettings 
         BorderStyle     =   0  'Kein
         Height          =   1290
         Left            =   75
         ScaleHeight     =   86
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   371
         TabIndex        =   6
         Top             =   195
         Width           =   5565
         Begin VB.CheckBox chkTestmode 
            Caption         =   "Test mode"
            Height          =   195
            Left            =   600
            TabIndex        =   16
            Top             =   450
            Width           =   1740
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
            Left            =   600
            Style           =   2  'Dropdown-Liste
            TabIndex        =   11
            Top             =   0
            Width           =   4890
         End
         Begin VB.ComboBox cboSpeed 
            Height          =   315
            Left            =   3675
            Style           =   2  'Dropdown-Liste
            TabIndex        =   9
            Top             =   405
            Width           =   1815
         End
         Begin VB.CheckBox chkEjectDisk 
            Caption         =   "Eject disk after write"
            Height          =   195
            Left            =   600
            TabIndex        =   8
            Top             =   720
            Value           =   1  'Aktiviert
            Width           =   1815
         End
         Begin VB.CheckBox chkFinalize 
            Caption         =   "Finalize disk"
            Height          =   195
            Left            =   600
            TabIndex        =   7
            Top             =   1005
            Value           =   1  'Aktiviert
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Drive:"
            Height          =   195
            Left            =   75
            TabIndex        =   12
            Top             =   75
            Width           =   435
         End
         Begin VB.Label lblWriteSpeed 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            Caption         =   "Write speed:"
            Height          =   195
            Left            =   2700
            TabIndex        =   10
            Top             =   450
            Width           =   930
         End
      End
   End
   Begin VB.Frame frmFile 
      Caption         =   "File"
      Height          =   690
      Left            =   97
      TabIndex        =   1
      Top             =   675
      Width           =   5940
      Begin VB.PictureBox picFile 
         BorderStyle     =   0  'Kein
         Height          =   465
         Left            =   75
         ScaleHeight     =   465
         ScaleWidth      =   5790
         TabIndex        =   2
         Top             =   180
         Width           =   5790
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   285
            Left            =   5175
            TabIndex        =   4
            Top             =   75
            Width           =   420
         End
         Begin VB.TextBox txtFile 
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   75
            Width           =   4965
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgISO 
      Left            =   5400
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "ISO images (*.iso)|*.iso"
      Flags           =   2
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISO Image burner"
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
      Width           =   2310
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
End
Attribute VB_Name = "frmBurnISO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cISOCD   As FL_CDISOWriter
Attribute cISOCD.VB_VarHelpID = -1

Private cDrvNfo             As New FL_DriveInfo
Private cCDNfo              As New FL_CDInfo

Private Sub ShowSpeeds()

    Dim i           As Integer
    Dim intSpeeds() As Integer

    intSpeeds = cDrvNfo.GetWriteSpeeds(strDrvID)

    cboSpeed.Clear
    For i = LBound(intSpeeds) To UBound(intSpeeds)
        cboSpeed.AddItem intSpeeds(i) & " KB/s (" & (intSpeeds(i) \ 176) & "x)"
        cboSpeed.ItemData(cboSpeed.ListCount - 1) = intSpeeds(i)
    Next
    cboSpeed.AddItem "Max."
    cboSpeed.ItemData(cboSpeed.ListCount - 1) = &HFFFF&

    cboSpeed.ListIndex = cboSpeed.ListCount - 1

End Sub

Private Sub cboDrv_Click()
    strDrvID = cManager.DrvChr2DrvID(Left$(cboDrv.List(cboDrv.ListIndex), 1))
    ShowSpeeds
End Sub

Private Sub cISOCD_ClosingSession()
    With frmDataCDPrg.lstStatus.ListItems
        With .Add(SmallIcon:=1)
            .SubItems(1) = "Closing session..."
        End With
    End With
End Sub

Private Sub cISOCD_Finished()
    With frmDataCDPrg.lstStatus.ListItems
        With .Add(SmallIcon:=1)
            .SubItems(1) = "Finished"
        End With
    End With
End Sub

Private Sub cISOCD_Progress(Percent As Integer)
    On Error Resume Next
    frmDataCDPrg.prg.Value = Percent
End Sub

Private Sub cISOCD_StartWriting()
    With frmDataCDPrg.lstStatus.ListItems
        With .Add(SmallIcon:=1)
            .SubItems(1) = "Writing track..."
        End With
    End With
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo ErrorHandler
    dlgISO.ShowOpen
    txtFile = dlgISO.FileName
ErrorHandler:
End Sub

Private Sub cmdDrvNfo_Click()
    frmDriveInfo.Show vbModal, Me
End Sub

Private Sub cmdWrite_Click()

    Dim strMsg  As String

    If txtFile = vbNullString Then
        MsgBox "No ISO image selected.", vbExclamation
        Exit Sub
    End If

    cCDNfo.GetInfo strDrvID
    If FileLen(txtFile) > cCDNfo.Capacity Then
        If MsgBox("Image size exceeds disk capacity." & vbCrLf & _
                  "Continue?", vbYesNo Or vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    cISOCD.ISOFile = txtFile
    cISOCD.EjectAfterWrite = chkEjectDisk
    cISOCD.NextSessionAllowed = Not CBool(chkFinalize)
    cISOCD.TestMode = chkTestmode

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
    Unload frmDataCDPrg

End Sub

Private Sub Form_Load()
    Set cISOCD = New FL_CDISOWriter
    ShowDrives
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmImgTools.Show
End Sub

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
