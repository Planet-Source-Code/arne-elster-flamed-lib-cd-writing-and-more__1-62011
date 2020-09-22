VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Flamed v4 CD Player"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
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
   ScaleHeight     =   132
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdNextTrack 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4725
      TabIndex        =   10
      Top             =   1050
      Width           =   390
   End
   Begin VB.CommandButton cmdPrevTrack 
      Caption         =   "<<"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4350
      TabIndex        =   9
      Top             =   1050
      Width           =   390
   End
   Begin prjCDPlayer.MasterVolume vol 
      Height          =   120
      Left            =   3525
      TabIndex        =   6
      ToolTipText     =   "Volume:  100%"
      Top             =   1695
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   212
      Segmented       =   0   'False
      SliderIcon      =   "frmMain.frx":0000
      Value           =   100
   End
   Begin prjCDPlayer.ucSlider sld 
      Height          =   240
      Left            =   150
      Top             =   1650
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   423
      Enabled         =   0   'False
      SliderIcon      =   "frmMain.frx":0389
      Orientation     =   0
      RailPicture     =   "frmMain.frx":049B
   End
   Begin VB.CheckBox chkDigital 
      Caption         =   "Digital mode"
      Height          =   315
      Left            =   3525
      TabIndex        =   5
      Top             =   75
      Width           =   1590
   End
   Begin VB.DriveListBox drv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Top             =   75
      Width           =   3240
   End
   Begin VB.ListBox lstTracks 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      IntegralHeight  =   0   'False
      Left            =   150
      TabIndex        =   3
      Top             =   450
      Width           =   3240
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3525
      TabIndex        =   2
      Top             =   1050
      Width           =   840
   End
   Begin VB.CommandButton cmdPauseResume 
      Caption         =   "Pause/Resume"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3525
      TabIndex        =   1
      Top             =   750
      Width           =   1590
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3525
      TabIndex        =   0
      Top             =   450
      Width           =   1590
   End
   Begin VB.Label lblTrack 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Track: 01"
      Height          =   195
      Left            =   2535
      TabIndex        =   11
      Top             =   1425
      Width           =   825
   End
   Begin VB.Label lblVol 
      AutoSize        =   -1  'True
      Caption         =   "Volume:"
      Height          =   195
      Left            =   3525
      TabIndex        =   8
      Top             =   1425
      Width           =   720
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      Caption         =   "Position:"
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   1425
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cPlayer  As FL_CDPlayer
Attribute cPlayer.VB_VarHelpID = -1
Private cManager            As FL_Manager

Private strDrvID            As String

Private blnNoTimer         As Boolean

Private Sub chkDigital_Click()

    ' mode changed, reinit
    cPlayer.StopTrack
    cPlayer.DigitalMode = chkDigital.Value

    ' get the volume setting for the mode
    vol.Value = cPlayer.Volume
    vol_ValueChanged

    drv_Change

End Sub

Private Sub cmdNextTrack_Click()
    cPlayer.NextTrack
End Sub

Private Sub cmdPauseResume_Click()
    cPlayer.PauseResumeTrack
End Sub

Private Sub cmdPlay_Click()
    If lstTracks.ListCount = 0 Then Exit Sub
    cPlayer.PlayTrack lstTracks.ListIndex + 1
End Sub

Private Sub cmdPrevTrack_Click()
    cPlayer.PrevTrack
End Sub

Private Sub cmdStop_Click()
    cPlayer.StopTrack
End Sub

Private Sub cPlayer_StateChanged(ByVal State As FlamedLib.FL_PlaybackState)
    Select Case State

        Case PBS_PAUSING
            cmdPauseResume.Enabled = True
            cmdPlay.Enabled = False
            cmdStop.Enabled = True
            sld.Enabled = True
            cmdNextTrack.Enabled = True
            cmdPrevTrack.Enabled = True

        Case PBS_PLAYING
            cmdPlay.Enabled = False
            cmdPauseResume.Enabled = True
            cmdStop.Enabled = True
            sld.Enabled = True
            cmdNextTrack.Enabled = True
            cmdPrevTrack.Enabled = True

        Case PBS_STOPPED
            cmdPlay.Enabled = True
            cmdPauseResume.Enabled = False
            cmdStop.Enabled = False
            sld.Enabled = False
            cmdNextTrack.Enabled = False
            cmdPrevTrack.Enabled = False

    End Select
End Sub

Private Sub cPlayer_Timer(ByVal pos As Long)

    ' if the slider isn't getting changed atm,
    ' set the new position
    If Not blnNoTimer Then sld.Value = pos

    With cPlayer.CurrentPos
        lblPos = "Position: " & Format(.M, "00") & ":" & Format(.s, "00")
    End With

End Sub

Private Sub cPlayer_TrackChanged(ByVal Track As Integer)
    ' set the new slider maximum
    sld.Max = cPlayer.TrackLength(Track).LBA
    lblTrack = "Track: " & Format(Track, "00")
End Sub

Private Sub cPlayer_TrackEndReached(ByVal Track As Integer, stopplay As Boolean)
    ' You may set stopplay to true
    ' so cPlayer will stop playback
End Sub

Private Sub drv_Change()

    cPlayer.StopTrack

    lstTracks.Clear
    strDrvID = vbNullString

    If cManager.IsCDVDDrive(drv.Drive) Then
        strDrvID = cManager.DrvChr2DrvID(drv.Drive)
        ShowTracks
        cmdPlay.Enabled = True
    Else
        cmdPlay.Enabled = False
    End If

End Sub

Sub ShowTracks()

    Dim i   As Integer

    If Not cPlayer.OpenDrive(strDrvID) Then
        MsgBox "Failed to read tracks.", vbExclamation, "Error"
        Exit Sub
    End If

    For i = 1 To cPlayer.TrackCount
        lstTracks.AddItem "Track " & Format(i, "00") & "      " & _
                          cPlayer.TrackLength(i).MSF
    Next

    If lstTracks.ListCount > 0 Then lstTracks.ListIndex = 0

End Sub

Private Sub Form_Load()

    Set cPlayer = New FL_CDPlayer
    Set cManager = New FL_Manager

    If Not cManager.Init Then
        MsgBox "No interfaces found.", vbExclamation, "Error"
        Unload Me
    End If

    chkDigital_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
    cManager.Goodbye
End Sub

Private Sub sld_MouseDown(Shift As Integer)
    blnNoTimer = True
End Sub

Private Sub sld_MouseUp(Shift As Integer)
    ' seek to the selected LBA
    cPlayer.SeekTrack sld.Value
    blnNoTimer = False
End Sub

Private Sub vol_ValueChanged()
    ' volume slider moved
    cPlayer.Volume = vol.Value
    lblVol = "Volume: " & Format(vol.Value, "00") & "%"
End Sub
