VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Audio Track Grabber"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
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
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4575
      Top             =   525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   2
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4125
      TabIndex        =   7
      Top             =   2550
      Width           =   915
   End
   Begin VB.CommandButton cmdGrab 
      Caption         =   "Grab"
      Default         =   -1  'True
      Height          =   315
      Left            =   3075
      TabIndex        =   6
      Top             =   2550
      Width           =   990
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   315
      Left            =   225
      TabIndex        =   5
      Top             =   2550
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ListBox lstBitrate 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   1245
      IntegralHeight  =   0   'False
      Left            =   3075
      TabIndex        =   4
      Top             =   1170
      Width           =   1965
   End
   Begin VB.OptionButton optMP3 
      Caption         =   "Grab to MP3"
      Height          =   240
      Left            =   2925
      TabIndex        =   3
      Top             =   900
      Width           =   2040
   End
   Begin VB.OptionButton optWAV 
      Caption         =   "Grab to WAV"
      Height          =   240
      Left            =   2925
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.ListBox lstTracks 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1800
      IntegralHeight  =   0   'False
      Left            =   225
      TabIndex        =   1
      Top             =   600
      Width           =   2565
   End
   Begin VB.DriveListBox cboDrv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   4890
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cGrabber As FL_TrackGrabber
Attribute cGrabber.VB_VarHelpID = -1
Private cCDInfo             As New FL_CDInfo
Private cTrackInfo          As New FL_TrackInfo
Private cManager            As New FL_Manager

Private strDrvID            As String

Private blnCancel           As Boolean

Private Sub cboDrv_Change()

    strDrvID = vbNullString
    lstTracks.Clear

    If cManager.IsCDVDDrive(cboDrv.Drive) Then
        strDrvID = cManager.DrvChr2DrvID(cboDrv.Drive)
        ShowAudioTracks
    End If

End Sub

Sub ShowAudioTracks()

    Dim i   As Integer

    ' read disk information
    If Not cCDInfo.GetInfo(strDrvID) Then
        MsgBox "Could not read CD information.", vbExclamation
        Exit Sub
    End If

    For i = 1 To cCDInfo.Tracks
        ' read track information
        If Not cTrackInfo.GetInfo(strDrvID, i) Then
            MsgBox "Could not get info about track " & i, vbAbortRetryIgnore
        Else
            ' is audio track?
            If cTrackInfo.mode = MODE_AUDIO Then
                ' add audio track to the list
                lstTracks.AddItem "Track " & Format(i, "00")
                lstTracks.ItemData(lstTracks.ListCount - 1) = i
            End If
        End If
    Next

    If lstTracks.ListCount = 0 Then
        MsgBox "No audio tracks found!", vbExclamation
    End If

End Sub

Private Sub cGrabber_Progress(ByVal Percent As Integer, ByVal Track As Integer, ByVal startLBA As Long, ByVal endLBA As Long, Cancel As Boolean)
    prg.Value = Percent
    Cancel = blnCancel
    DoEvents
End Sub

Private Sub cmdCancel_Click()
    blnCancel = True
End Sub

Private Sub cmdGrab_Click()

    Dim ret As FL_SAVETRACK

    blnCancel = False

    ' track selected?
    If lstTracks.ListIndex < 0 Then
        MsgBox "No track selected."
        Exit Sub
    End If

    ' select filter for common dialog
    If optWAV Then
        dlg.Filter = "PCM WAV (*.wav)|*.wav"
    Else
        If lstBitrate.ListIndex < 0 Then
            MsgBox "No bitrate selected.", vbExclamation
            Exit Sub
        End If
        dlg.Filter = "MPEG-3 audio (*.mp3)|*.mp3"
    End If

    On Error GoTo ErrorHandler
    dlg.ShowSave
    On Error GoTo 0

    cmdGrab.Enabled = Not cmdGrab.Enabled
    cmdCancel.Enabled = Not cmdCancel.Enabled

    ' grab to WAV or MP3
    If optWAV Then
        ret = cGrabber.AudioTrackToWAV(strDrvID, _
                                       lstTracks.ItemData(lstTracks.ListIndex), _
                                       dlg.FileName)

    Else
        ret = cGrabber.AudioTrackToMP3(strDrvID, _
                                       lstTracks.ItemData(lstTracks.ListIndex), _
                                       dlg.FileName, _
                                       lstBitrate.ItemData(lstBitrate.ListIndex))

    End If

    cmdGrab.Enabled = Not cmdGrab.Enabled
    cmdCancel.Enabled = Not cmdCancel.Enabled

    Select Case ret
        Case ST_CANCELED: MsgBox "Canceled.", vbInformation
        Case ST_ENCODER_INIT: MsgBox "Failed to initialize encoder.", vbExclamation
        Case ST_FINISHED: MsgBox "Finished.", vbInformation
        Case ST_INVALID_SESSION: MsgBox "Invalid session."
        Case ST_INVALID_TRACKMODE: MsgBox "Invalid track mode."
        Case ST_INVALID_TRACKNO: MsgBox "Invalid track number."
        Case ST_NOT_READY: MsgBox "Unit not ready.", vbExclamation
        Case ST_READ_ERR: MsgBox "Read error.", vbExclamation
        Case ST_UNKNOWN_ERR: MsgBox "Unknown error.", vbExclamation
        Case ST_WRITE_ERR: MsgBox "Write error.", vbExclamation
    End Select

ErrorHandler:

End Sub

Private Sub Form_Load()

    If Not cManager.Init() Then
        MsgBox "No interfaces found.", vbExclamation
        Unload Me
    End If

    Set cGrabber = New FL_TrackGrabber

    AddBitrates

End Sub

Private Sub Form_Unload(Cancel As Integer)
    cManager.Goodbye
End Sub

Private Sub optMP3_Click()
    lstBitrate.Enabled = optMP3.Value
End Sub

Private Sub optWAV_Click()
    lstBitrate.Enabled = optMP3.Value
End Sub

Private Sub AddBitrates()
    ' add most compatible bitrates
    With lstBitrate
        .AddItem "128 KBit"
        .ItemData(0) = [128 kBits]
        .AddItem "160 KBit"
        .ItemData(1) = [160 kBits]
        .AddItem "192 KBit"
        .ItemData(2) = [192 kBits]
        .AddItem "224 KBit"
        .ItemData(3) = [224 kBits]
        .AddItem "256 KBit"
        .ItemData(4) = [256 kBits]
        .AddItem "320 KBit"
        .ItemData(5) = [320 kBits]
    End With
End Sub
