VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Mode-1 Track Grabber"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3285
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   80
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   219
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2700
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   2
   End
   Begin VB.DriveListBox cboDrv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   75
      TabIndex        =   4
      Top             =   75
      Width           =   3165
   End
   Begin VB.ComboBox cboTrack 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown-Liste
      TabIndex        =   3
      Top             =   480
      Width           =   765
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Default         =   -1  'True
      Height          =   315
      Left            =   2175
      TabIndex        =   2
      Top             =   450
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2175
      TabIndex        =   1
      Top             =   825
      Width           =   990
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   825
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblSession 
      Caption         =   "Track:"
      Height          =   240
      Left            =   150
      TabIndex        =   5
      Top             =   525
      Width           =   915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cReader  As FL_TrackGrabber
Attribute cReader.VB_VarHelpID = -1
Private cManager            As New FL_Manager
Private cInfo               As New FL_CDInfo
Private cTrack              As New FL_TrackInfo

Private strDrvID            As String
Private blnCancel           As Boolean

Private Sub cboDrv_Change()

    strDrvID = vbNullString
    cboTrack.Clear

    If cManager.IsCDVDDrive(cboDrv.Drive) Then

        strDrvID = cManager.DrvChr2DrvID(cboDrv.Drive)

        If Not cInfo.GetInfo(strDrvID) Then
            MsgBox "Could not read CD information.", vbExclamation, "Error"
            Exit Sub
        End If

        ShowTracks

    End If
End Sub

Private Sub ShowTracks()
    Dim i   As Integer
    For i = 1 To cInfo.Tracks
        cboTrack.AddItem Format(i, "00")
    Next
    cboTrack.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    blnCancel = True
End Sub

Private Sub cmdRead_Click()

    Dim udeRet  As FL_SAVETRACK

    If Not cTrack.GetInfo(strDrvID, cboTrack.List(cboTrack.ListIndex)) Then
        MsgBox "Couldn't read track information.", vbExclamation, "Error"
        Exit Sub
    End If

    If cTrack.mode = MODE_MODE1 Then

        ' MODE-1 ISO image
        dlg.FileName = vbNullString
        dlg.Filter = "ISO images (*.iso)|*.iso"
        dlg.ShowSave
        If dlg.FileName = vbNullString Then Exit Sub

        cmdRead.Enabled = Not cmdRead.Enabled
        cmdCancel.Enabled = Not cmdCancel.Enabled

        udeRet = cReader.DataTrackToISO(strDrvID, cboTrack.List(cboTrack.ListIndex), dlg.FileName)

    Else

        ' never seen an Mode-2 ISO image...
        MsgBox "Track mode not supported.", vbExclamation, "Error"
        Exit Sub

    End If

    ' save selected session
    Select Case udeRet
        Case ST_CANCELED: MsgBox "Canceled.", vbExclamation, "Canceled"
        Case ST_FINISHED: MsgBox "Finished.", vbInformation, "Ok"
        Case ST_INVALID_SESSION: MsgBox "Invalid session.", vbExclamation, "Error"
        Case ST_INVALID_TRACKMODE: MsgBox "Invalid Track mode.", vbExclamation, "Error"
        Case ST_INVALID_TRACKNO: MsgBox "Invalid track number.", vbExclamation, "Error"
        Case ST_NOT_READY: MsgBox "Drive not ready.", vbExclamation, "Error"
        Case ST_READ_ERR: MsgBox "Read error.", vbExclamation, "Error"
        Case ST_UNKNOWN_ERR: MsgBox "Unknown error.", vbExclamation, "Error"
        Case ST_WRITE_ERR: MsgBox "Write error. Not enough disk space?", vbExclamation, "Error"
    End Select

    cmdRead.Enabled = Not cmdRead.Enabled
    cmdCancel.Enabled = Not cmdCancel.Enabled

End Sub

Private Sub Form_Load()
    Set cReader = New FL_TrackGrabber

    If Not cManager.Init() Then
        MsgBox "No interfaces found.", vbExclamation, "Error"
        Unload Me
    End If
End Sub

Private Sub cReader_Progress(ByVal percent As Integer, ByVal Track As Integer, ByVal startLBA As Long, ByVal endLBA As Long, Cancel As Boolean)
    ' from time to time there are percent
    ' values greater than 100, so...
    On Error Resume Next

    ' show percent
    prg.Value = percent
    ' user may clicked cancel button
    Cancel = blnCancel
    '
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cManager.Goodbye
End Sub
