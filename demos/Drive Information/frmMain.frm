VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "drive information"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   379
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox lst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2985
      Left            =   322
      TabIndex        =   1
      Top             =   750
      Width           =   5040
   End
   Begin VB.DriveListBox cboDrvs 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   322
      TabIndex        =   0
      Top             =   300
      Width           =   5040
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cManager    As New FL_Manager
Private cInfo       As New FL_DriveInfo

Private strDriveID  As String

Private Sub cboDrvs_Change()
    lst.Clear
    If cManager.IsCDVDDrive(cboDrvs.Drive) Then
        strDriveID = cManager.DrvChr2DrvID(cboDrvs.Drive)
        Info
    End If
End Sub

Private Sub Info()

    If Not cInfo.GetInfo(strDriveID) Then
        MsgBox "Failed to read information.", vbExclamation, "Error"
        Exit Sub
    End If

    With cInfo
        lst.AddItem "Disc present: " & .DiscPresent
        lst.AddItem "Drive closed: " & .DriveClosed
        lst.AddItem "Is lockable: " & .Lockable
        lst.AddItem "Drive locked: " & .DriveLocked
        lst.AddItem "Loading mechanism: " & LoadingMech2Str(.LoadingMechanism)
        lst.AddItem ""
        lst.AddItem "Idle Timer: " & .IdleTimer100MS & " MS"
        lst.AddItem "Standby Timer: " & .StandbyTimer100MS & " MS"
        lst.AddItem "Spindown Timer: " & .SpinDownTimerMS & " MS"
        lst.AddItem ""
        lst.AddItem "Supports playing analog audio: " & .AnalogAudioPlayback
        lst.AddItem "Has jitter effect correction: " & .JitterEffectCorrection
        lst.AddItem ""
        lst.AddItem "Vendor: " & .Vendor
        lst.AddItem "Product: " & .Product
        lst.AddItem "Revision: " & .Revision
        lst.AddItem "Interface: " & IPh2Str(.PhysicalInterface)
        lst.AddItem "Host Adapter: " & .HostAdapter
        lst.AddItem "Target: " & .Target
        lst.AddItem ""
        lst.AddItem "Max read speed: " & .ReadSpeedMax & " KB/s"
        lst.AddItem "Cur read speed: " & .ReadSpeedCur & " KB/s"
        lst.AddItem "Max write speed: " & .WriteSpeedMax & " KB/s"
        lst.AddItem "Cur write speed: " & .WriteSpeedCur & " KB/s"
        lst.AddItem ""
        lst.AddItem "Reads barcode: " & CBool(.ReadCapabilities And RC_BARCODE)
        lst.AddItem "Reads C2 Error Pointers: " & CBool(.ReadCapabilities And RC_C2)
        lst.AddItem "Reads raw CDDA: " & CBool(.ReadCapabilities And RC_CDDARAW)
        lst.AddItem "Reads CD-R: " & CBool(.ReadCapabilities And RC_CDR)
        lst.AddItem "Reads CD-RW: " & CBool(.ReadCapabilities And RC_CDRW)
        lst.AddItem "Reads CD-Text: " & CBool(.ReadCapabilities And RC_CDTEXT)
        lst.AddItem "Reads DVD+R: " & CBool(.ReadCapabilities And RC_DVDPR)
        lst.AddItem "Reads DVD+RW: " & CBool(.ReadCapabilities And RC_DVDPRW)
        lst.AddItem "Reads DVD+R DL: " & CBool(.ReadCapabilities And RC_DVDPRDL)
        lst.AddItem "Reads DVD-R: " & CBool(.ReadCapabilities And RC_DVDR)
        lst.AddItem "Reads DVD-RAM: " & CBool(.ReadCapabilities And RC_DVDRAM)
        lst.AddItem "Reads DVD-ROM: " & CBool(.ReadCapabilities And RC_DVDROM)
        lst.AddItem "Reads ISRC: " & CBool(.ReadCapabilities And RC_ISRC)
        lst.AddItem "Reads Mode 2 Form 1 sectors: " & CBool(.ReadCapabilities And RC_MODE2FORM1)
        lst.AddItem "Reads Mode 2 Form 2 sectors: " & CBool(.ReadCapabilities And RC_MODE2FORM2)
        lst.AddItem "Reads multi-session discs: " & CBool(.ReadCapabilities And RC_MULTISESSION)
        lst.AddItem "Reads Mount Rainer: " & CBool(.ReadCapabilities And RC_MRW)
        lst.AddItem "Reads sub-channels: " & CBool(.ReadCapabilities And RC_SUBCHANNELS)
        lst.AddItem "Reads sub-channels corrected: " & CBool(.ReadCapabilities And RC_SUBCHANNELS_CORRECTED)
        lst.AddItem "Reads sub-channel from Lead-In: " & CBool(.ReadCapabilities And RC_SUBCHANNELS_FROM_LEADIN)
        lst.AddItem ""
        lst.AddItem "Buffer size: " & .BufferSizeKB & " KB"
        lst.AddItem "Writes CD-R: " & CBool(.WriteCapabilities And WC_CDR)
        lst.AddItem "Writes CD-RW: " & CBool(.WriteCapabilities And WC_CDRW)
        lst.AddItem "Writes DVD+R: " & CBool(.WriteCapabilities And WC_DVDPR)
        lst.AddItem "Writes DVD+RW: " & CBool(.WriteCapabilities And WC_DVDPRW)
        lst.AddItem "Writes DVD+R DL: " & CBool(.WriteCapabilities And WC_DVDPRDL)
        lst.AddItem "Writes DVD-R: " & CBool(.WriteCapabilities And WC_DVDR)
        lst.AddItem "Writes DVD-RW: " & CBool(.WriteCapabilities And WC_DVDRRW)
        lst.AddItem "Writes DVD-RAM: " & CBool(.WriteCapabilities And WC_DVDRAM)
        lst.AddItem "Writes RAW+16: " & CBool(.WriteCapabilities And WC_RAW_16)
        lst.AddItem "Writes RAW+16 with Testmode: " & CBool(.WriteCapabilities And WC_RAW_16_TEST)
        lst.AddItem "Writes RAW+96: " & CBool(.WriteCapabilities And WC_RAW_96)
        lst.AddItem "Writes RAW+96 with Testmode: " & CBool(.WriteCapabilities And WC_RAW_96_TEST)
        lst.AddItem "Writes SAO: " & CBool(.WriteCapabilities And WC_SAO)
        lst.AddItem "Writes SAO with Testmode: " & CBool(.WriteCapabilities And WC_SAO_TEST)
        lst.AddItem "Writes TAO: " & CBool(.WriteCapabilities And WC_TAO)
        lst.AddItem "Writes TAO with Testmode: " & CBool(.WriteCapabilities And WC_TAO_TEST)
        lst.AddItem "Writes Mount Rainer: " & CBool(.WriteCapabilities And WC_MRW)
        lst.AddItem "Supports Testmode: " & CBool(.WriteCapabilities And WC_TESTMODE)
        lst.AddItem "Supports BURN-Proof technology: " & CBool(.WriteCapabilities And WC_BURNPROOF)
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

Private Sub Form_Load()
    If Not cManager.Init(False) Then
        MsgBox "No interfaces found.", vbExclamation, "Error"
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cManager.Goodbye
End Sub
