VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAudioGrabber 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Flamed v4 - Audio Grabber"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
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
   ScaleHeight     =   484
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      Caption         =   "Player"
      Height          =   1290
      Left            =   75
      TabIndex        =   58
      Top             =   1200
      Width           =   2265
      Begin VB.PictureBox picPlayer 
         BorderStyle     =   0  'Kein
         Height          =   990
         Left            =   75
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   59
         Top             =   225
         Width           =   2115
         Begin prjFlamed.ucSlider prg 
            Height          =   240
            Left            =   150
            Top             =   675
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   423
            SliderIcon      =   "frmAudioGrabber.frx":0000
            Orientation     =   0
            RailPicture     =   "frmAudioGrabber.frx":0112
         End
         Begin VB.CommandButton cmdPlay 
            Height          =   315
            Left            =   825
            Picture         =   "frmAudioGrabber.frx":012E
            Style           =   1  'Grafisch
            TabIndex        =   62
            Top             =   0
            Width           =   390
         End
         Begin VB.CommandButton cmdPause 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1275
            Picture         =   "frmAudioGrabber.frx":04C8
            Style           =   1  'Grafisch
            TabIndex        =   61
            Top             =   0
            Width           =   390
         End
         Begin VB.CommandButton cmdStop 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1725
            Picture         =   "frmAudioGrabber.frx":086D
            Style           =   1  'Grafisch
            TabIndex        =   60
            Top             =   0
            Width           =   390
         End
         Begin VB.Label lblTrackPos 
            Alignment       =   2  'Zentriert
            AutoSize        =   -1  'True
            Caption         =   "00:00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   64
            Top             =   375
            Width           =   480
         End
         Begin VB.Label lblTrack 
            AutoSize        =   -1  'True
            Caption         =   "01"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   150
            TabIndex        =   63
            Top             =   0
            Width           =   300
         End
      End
   End
   Begin prjFlamed.isButton cmdReset 
      Height          =   300
      Left            =   4725
      TabIndex        =   57
      Top             =   2595
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      Icon            =   "frmAudioGrabber.frx":0C07
      Style           =   1
      Caption         =   "R"
      iNonThemeStyle  =   1
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
   Begin VB.Frame frmStartEnd 
      Caption         =   "Relative start/end addresses"
      Height          =   990
      Left            =   2400
      TabIndex        =   38
      Top             =   2625
      Width           =   4665
      Begin VB.PictureBox picStartEnd 
         BorderStyle     =   0  'Kein
         Height          =   735
         Left            =   75
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   301
         TabIndex        =   39
         Top             =   225
         Width           =   4515
         Begin VB.TextBox txtEndF 
            Alignment       =   2  'Zentriert
            Height          =   285
            Left            =   1650
            MaxLength       =   2
            TabIndex        =   52
            Text            =   "00"
            Top             =   375
            Width           =   315
         End
         Begin VB.TextBox txtEndS 
            Alignment       =   2  'Zentriert
            Height          =   285
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   50
            Text            =   "00"
            Top             =   375
            Width           =   315
         End
         Begin VB.TextBox txtEndM 
            Alignment       =   2  'Zentriert
            Height          =   285
            Left            =   750
            MaxLength       =   2
            TabIndex        =   48
            Text            =   "00"
            Top             =   375
            Width           =   315
         End
         Begin VB.TextBox txtStartF 
            Alignment       =   2  'Zentriert
            Height          =   285
            Left            =   1650
            MaxLength       =   2
            TabIndex        =   45
            Text            =   "00"
            Top             =   30
            Width           =   315
         End
         Begin VB.TextBox txtStartS 
            Alignment       =   2  'Zentriert
            Height          =   285
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   43
            Text            =   "00"
            Top             =   30
            Width           =   315
         End
         Begin VB.TextBox txtStartM 
            Alignment       =   2  'Zentriert
            Height          =   285
            Left            =   750
            MaxLength       =   2
            TabIndex        =   41
            Text            =   "00"
            Top             =   30
            Width           =   315
         End
         Begin VB.Label lblLengthMSF 
            AutoSize        =   -1  'True
            Caption         =   "00:00:00 MSF"
            Height          =   195
            Left            =   3375
            TabIndex        =   56
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lblLength 
            Caption         =   "Length:"
            Height          =   240
            Left            =   2745
            TabIndex        =   55
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "}"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2475
            TabIndex        =   54
            Top             =   150
            Width           =   135
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "MSF"
            Height          =   195
            Left            =   2040
            TabIndex        =   53
            Top             =   420
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   ":"
            Height          =   195
            Left            =   1560
            TabIndex        =   51
            Top             =   420
            Width           =   60
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   ":"
            Height          =   195
            Left            =   1110
            TabIndex        =   49
            Top             =   420
            Width           =   60
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "End:"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   420
            Width           =   330
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "MSF"
            Height          =   195
            Left            =   2040
            TabIndex        =   46
            Top             =   75
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   ":"
            Height          =   195
            Left            =   1560
            TabIndex        =   44
            Top             =   75
            Width           =   60
         End
         Begin VB.Label lblSeperator1 
            AutoSize        =   -1  'True
            Caption         =   ":"
            Height          =   195
            Left            =   1110
            TabIndex        =   42
            Top             =   75
            Width           =   60
         End
         Begin VB.Label lblStartMSF 
            AutoSize        =   -1  'True
            Caption         =   "Start:"
            Height          =   195
            Left            =   150
            TabIndex        =   40
            Top             =   75
            Width           =   420
         End
      End
   End
   Begin VB.CommandButton cmdDrvNfo 
      Caption         =   "Drive information"
      Height          =   315
      Left            =   3975
      TabIndex        =   37
      Top             =   6825
      Width           =   1440
   End
   Begin VB.ComboBox cboSpeed 
      Height          =   315
      Left            =   6225
      Style           =   2  'Dropdown-Liste
      TabIndex        =   34
      ToolTipText     =   "Readspeed"
      Top             =   675
      Width           =   765
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "Back"
      Height          =   330
      Left            =   150
      TabIndex        =   33
      Top             =   6825
      Width           =   1365
   End
   Begin VB.CommandButton cmdGrab 
      Caption         =   "Grab tracks"
      Height          =   315
      Left            =   5550
      TabIndex        =   32
      Top             =   6825
      Width           =   1365
   End
   Begin MSComctlLib.ImageList img 
      Left            =   3975
      Top             =   675
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
            Picture         =   "frmAudioGrabber.frx":0C23
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAudioGrabber.frx":33D5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmTags 
      Caption         =   "Tags"
      Height          =   1365
      Left            =   2400
      TabIndex        =   18
      Top             =   1200
      Width           =   4665
      Begin VB.PictureBox picTags 
         BorderStyle     =   0  'Kein
         Height          =   1065
         Left            =   75
         ScaleHeight     =   71
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   301
         TabIndex        =   19
         Top             =   225
         Width           =   4515
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   285
            Left            =   3375
            TabIndex        =   30
            Top             =   750
            Width           =   1065
         End
         Begin VB.CommandButton cmdFreeDB 
            Caption         =   "FreeDB"
            Height          =   285
            Left            =   3375
            TabIndex        =   29
            Top             =   397
            Width           =   1065
         End
         Begin VB.CommandButton cmdCDText 
            Caption         =   "CD-Text"
            Height          =   285
            Left            =   3375
            TabIndex        =   28
            Top             =   45
            Width           =   1065
         End
         Begin VB.TextBox txtTitle 
            Height          =   285
            Left            =   675
            TabIndex        =   25
            Top             =   750
            Width           =   2565
         End
         Begin VB.TextBox txtArtist 
            Height          =   285
            Left            =   675
            TabIndex        =   23
            Top             =   397
            Width           =   2565
         End
         Begin VB.TextBox txtAlbum 
            Height          =   285
            Left            =   675
            TabIndex        =   21
            Top             =   45
            Width           =   2565
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Title:"
            Height          =   195
            Left            =   210
            TabIndex        =   24
            Top             =   780
            Width           =   360
         End
         Begin VB.Label lblArtist 
            AutoSize        =   -1  'True
            Caption         =   "Artist:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   435
            Width           =   450
         End
         Begin VB.Label lblAlbum 
            AutoSize        =   -1  'True
            Caption         =   "Album:"
            Height          =   195
            Left            =   75
            TabIndex        =   20
            Top             =   75
            Width           =   495
         End
      End
   End
   Begin VB.Frame frmSave 
      Caption         =   "Save options"
      Height          =   1215
      Left            =   2400
      TabIndex        =   12
      Top             =   5475
      Width           =   4665
      Begin VB.PictureBox picSave 
         BorderStyle     =   0  'Kein
         Height          =   915
         Left            =   75
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   301
         TabIndex        =   13
         Top             =   225
         Width           =   4515
         Begin VB.CommandButton cmdBrowsePath 
            Caption         =   "..."
            Height          =   285
            Left            =   3975
            TabIndex        =   31
            Top             =   60
            Width           =   465
         End
         Begin VB.ComboBox cboFilemask 
            Height          =   315
            ItemData        =   "frmAudioGrabber.frx":352F
            Left            =   825
            List            =   "frmAudioGrabber.frx":3545
            Style           =   2  'Dropdown-Liste
            TabIndex        =   17
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox txtPath 
            Height          =   285
            Left            =   525
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   60
            Width           =   3390
         End
         Begin VB.Label lblFilemask 
            AutoSize        =   -1  'True
            Caption         =   "Filemask:"
            Height          =   195
            Left            =   75
            TabIndex        =   16
            Top             =   525
            Width           =   660
         End
         Begin VB.Label lblPath 
            Caption         =   "Path:"
            Height          =   165
            Left            =   75
            TabIndex        =   14
            Top             =   75
            Width           =   540
         End
      End
   End
   Begin VB.Frame frmEncoder 
      Caption         =   "Encoder"
      Height          =   1740
      Left            =   2400
      TabIndex        =   3
      Top             =   3675
      Width           =   4665
      Begin VB.PictureBox picEncoder 
         BorderStyle     =   0  'Kein
         Height          =   1365
         Left            =   75
         ScaleHeight     =   91
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   301
         TabIndex        =   4
         Top             =   225
         Width           =   4515
         Begin VB.ComboBox cboSampleRate 
            Height          =   315
            ItemData        =   "frmAudioGrabber.frx":35F2
            Left            =   1275
            List            =   "frmAudioGrabber.frx":35FC
            Style           =   2  'Dropdown-Liste
            TabIndex        =   27
            Top             =   690
            Width           =   915
         End
         Begin VB.CheckBox chkID3v2 
            Caption         =   "Write ID3v2 Tags"
            Enabled         =   0   'False
            Height          =   240
            Left            =   2625
            TabIndex        =   11
            Top             =   1050
            Value           =   1  'Aktiviert
            Width           =   1740
         End
         Begin VB.CheckBox chkID3v1 
            Caption         =   "Write ID3v1 Tags"
            Enabled         =   0   'False
            Height          =   240
            Left            =   2625
            TabIndex        =   10
            Top             =   750
            Value           =   1  'Aktiviert
            Width           =   1665
         End
         Begin VB.ComboBox cboBitrate 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3225
            Style           =   2  'Dropdown-Liste
            TabIndex        =   9
            Top             =   375
            Width           =   1140
         End
         Begin VB.OptionButton optMP3 
            Caption         =   "MP3 (ACM Codec)"
            Height          =   240
            Left            =   2400
            TabIndex        =   7
            Top             =   75
            Width           =   1890
         End
         Begin VB.CheckBox chkWrtHdr 
            Caption         =   "Write header"
            Height          =   195
            Left            =   225
            TabIndex        =   6
            Top             =   405
            Value           =   1  'Aktiviert
            Width           =   1365
         End
         Begin VB.OptionButton optWAV 
            Caption         =   "PCM WAV"
            Height          =   195
            Left            =   75
            TabIndex        =   5
            Top             =   75
            Value           =   -1  'True
            Width           =   1440
         End
         Begin VB.Label lblSR 
            AutoSize        =   -1  'True
            Caption         =   "Sample Rate:"
            Height          =   195
            Left            =   225
            TabIndex        =   26
            Top             =   750
            Width           =   960
         End
         Begin VB.Label lblBitrate 
            AutoSize        =   -1  'True
            Caption         =   "Bitrate:"
            Height          =   195
            Left            =   2625
            TabIndex        =   8
            Top             =   420
            Width           =   540
         End
      End
   End
   Begin MSComctlLib.ListView lstTracks 
      Height          =   4140
      Left            =   75
      TabIndex        =   2
      Top             =   2550
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   7303
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Track"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Length"
         Object.Width           =   2117
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
      TabIndex        =   0
      Top             =   675
      Width           =   4815
   End
   Begin VB.Label lblSpeed 
      AutoSize        =   -1  'True
      Caption         =   "Speed:"
      Height          =   195
      Left            =   5625
      TabIndex        =   36
      Top             =   750
      Width           =   510
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "Drive:"
      Height          =   195
      Left            =   150
      TabIndex        =   35
      Top             =   750
      Width           =   435
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CDDA Grabber"
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
      Left            =   225
      TabIndex        =   1
      Top             =   120
      Width           =   1845
   End
   Begin VB.Shape shpHdr 
      BackColor       =   &H00336471&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   -150
      Top             =   0
      Width           =   7365
   End
   Begin VB.Menu mnuRClick 
      Caption         =   "RClick"
      Visible         =   0   'False
      Begin VB.Menu mnuCheckAll 
         Caption         =   "Check all"
      End
      Begin VB.Menu mnuUncheckAll 
         Caption         =   "Uncheck all"
      End
   End
End
Attribute VB_Name = "frmAudioGrabber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cFreeDB      As FL_FreeDB
Attribute cFreeDB.VB_VarHelpID = -1
Private WithEvents cMonitor     As FL_DoorMonitor
Attribute cMonitor.VB_VarHelpID = -1
Private WithEvents cCDPlay      As FL_CDPlayer
Attribute cCDPlay.VB_VarHelpID = -1
Private WithEvents cGrab        As FL_TrackGrabber
Attribute cGrab.VB_VarHelpID = -1

Private cDrvNfo     As New FL_DriveInfo
Private cTrkNfo     As New FL_TrackInfo
Private cCDNfo      As New FL_CDInfo
Private cCDText     As New FL_CDText
Private cID3v1      As New clsID3v1
Private cID3v2      As New clsID3v2

Private udtTracks   As t_AudioTracks
Private blnCancel   As Boolean
Private blnNoTmr    As Boolean

Public Property Let Cancel(aval As Boolean)
    blnCancel = aval
End Property

Private Sub cboDrv_Click()
    strDrvID = cManager.DrvChr2DrvID(Left$(cboDrv.List(cboDrv.ListIndex), 1))
    cDrvNfo.GetInfo strDrvID
    cCDPlay.OpenDrive strDrvID
    ShowSpeeds
    ShowTracks
End Sub

Private Sub ShowSpeeds()

    Dim i           As Integer
    Dim intSpeeds() As Integer

    cboSpeed.Clear

    ' get read speeds in 4x steps
    '
    ' a drive may not support all of them,
    ' but most firmwares automatically select
    ' the nearest supported read speed.
    intSpeeds = cDrvNfo.GetReadSpeeds(strDrvID)

    For i = LBound(intSpeeds) To UBound(intSpeeds)
        cboSpeed.AddItem (intSpeeds(i) \ 176) & "x"
        cboSpeed.ItemData(cboSpeed.ListCount - 1) = intSpeeds(i)
    Next

    cboSpeed.ListIndex = cboSpeed.ListCount - 1

End Sub

Private Sub ShowTracks()

    Dim i   As Integer

    lstTracks.ListItems.Clear

    If Not cCDNfo.GetInfo(strDrvID) Then
        MsgBox "Could not read disk.", vbExclamation
        Exit Sub
    End If

    For i = 1 To cCDNfo.Tracks

        ' dummy tags
        udtTracks.Track(i - 1).Album = "New Album (" & Date & ")"
        udtTracks.Track(i - 1).Artist = "New Artist"
        udtTracks.Track(i - 1).Title = "Title " & Format(i, "00")
        udtTracks.Track(i - 1).no = i

        If Not cTrkNfo.GetInfo(strDrvID, i) Then
            MsgBox "Could not get info about track " & i, vbExclamation
            udtTracks.Track(i - 1).grab = False
        Else
            udtTracks.Track(i - 1).grab = cTrkNfo.Mode = MODE_AUDIO
        End If

        ' start/end LBAs
        ' -150 because we want FL_MSF to display 00:00:00 MSF as start
        udtTracks.Track(i - 1).startLBA = -150
        udtTracks.Track(i - 1).endLBA = cTrkNfo.TrackLength.LBA
        udtTracks.Track(i - 1).lenLBA = cTrkNfo.TrackLength.LBA

        With lstTracks.ListItems
            With .Add(Text:=Format(i, "00"), SmallIcon:=Abs(Not (cTrkNfo.Mode = MODE_AUDIO)) + 1)
                .SubItems(1) = cTrkNfo.TrackLength.MSF & " MSF"
                .Checked = udtTracks.Track(i - 1).grab
                .Tag = i
            End With
        End With

    Next

    udtTracks.count = cCDNfo.Tracks

End Sub

Private Sub cboSpeed_Click()
    cManager.SetCDRomSpeed strDrvID, cboSpeed.ItemData(cboSpeed.ListIndex), &HFFF&
End Sub

Private Sub cCDPlay_StateChanged(ByVal State As FlamedLib.FL_PlaybackState)
    Select Case State
        Case PBS_PAUSING
            cmdPlay.Enabled = False
            cmdStop.Enabled = True
            cmdPause.Enabled = True
        Case PBS_PLAYING:
            cmdPlay.Enabled = False
            cmdPause.Enabled = True
            cmdStop.Enabled = True
        Case PBS_STOPPED:
            cmdPlay.Enabled = True
            cmdPause.Enabled = False
            cmdStop.Enabled = False
    End Select
End Sub

Private Sub cCDPlay_Timer(ByVal POS As Long)

    ' slider gets currently moved?
    If Not blnNoTmr Then prg.Value = POS

    With cCDPlay.CurrentPos
        lblTrackPos = Format(.m, "00") & ":" & Format(.s, "00")
    End With

End Sub

Private Sub cCDPlay_TrackChanged(ByVal Track As Integer)
    lblTrack = Format(Track, "00")
    prg.Max = cCDPlay.TrackLength(Track).LBA
End Sub

Private Sub cFreeDB_Status(Status As FlamedLib.FL_CDDBState)
    Select Case Status
        Case CDDB_CLOSE: Debug.Print "Closing FreeDB connection"
        Case CDDB_DATA: Debug.Print "Recieving data..."
        Case CDDB_HELLO: Debug.Print "FreeDB Hello"
        Case CDDB_QUERY: Debug.Print "Querying FreeDB..."
        Case CDDB_RESULT: Debug.Print "Getting FreeDB result..."
    End Select
End Sub

Private Sub cmdBack_Click()
    frmSelectProject.Show
    Unload Me
End Sub

Private Sub cmdBrowsePath_Click()

    Dim strRet  As String

    strRet = BrowseForFolder("Please select a new directory", _
                             txtPath, hWnd, True, , True)

    If Not strRet = vbNullString Then
        txtPath.Text = AddSlash(strRet)
    End If

End Sub

Private Sub cmdCDText_Click()

    Dim i   As Integer

    ' check for CD-Text feature
    ' I think some drives could not support
    ' feature codes but support reading CD-Text.
    If Not CBool(cDrvNfo.ReadCapabilities And RC_CDTEXT) Then
        If MsgBox("Your drive reports it can't read CD-Text." & vbCrLf & _
                  "Continue?", vbYesNo Or vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    If Not cCDText.ReadCDText(strDrvID) Then
        MsgBox "Failed.", vbExclamation
        Exit Sub
    End If

    With cCDText
        For i = 1 To .TrackCount
            udtTracks.Track(i - 1).Album = .Album
            udtTracks.Track(i - 1).Artist = .Artist
            udtTracks.Track(i - 1).Title = .Track(i - 1)
        Next
    End With

    On Error Resume Next
    ShowTrackTags lstTracks.SelectedItem.Tag

    MsgBox "Finished.", vbInformation

End Sub

Private Sub cmdClear_Click()

    On Error Resume Next

    With udtTracks.Track(lstTracks.SelectedItem.Tag - 1)
        .Album = ""
        .Artist = ""
        .Title = ""
    End With

    ShowTrackTags lstTracks.SelectedItem.Tag

End Sub

Private Sub cmdDrvNfo_Click()
    frmDriveInfo.Show vbModal, Me
End Sub

Private Sub cmdFreeDB_Click()

    Dim i   As Integer

    If Not cFreeDB.IsConnectedToInternet Then
        MsgBox "You're not connected to the internet.", vbExclamation
        Exit Sub
    End If

    cFreeDB.DriveID = strDrvID
    If Not cFreeDB.Query(False) Then
        MsgBox "Failed.", vbExclamation
        Exit Sub
    End If

    With cFreeDB
        For i = 1 To .Tracks
            udtTracks.Track(i - 1).Album = .Album
            udtTracks.Track(i - 1).Artist = .Artist
            udtTracks.Track(i - 1).Title = .Track(i)
        Next
    End With

    On Error Resume Next
    ShowTrackTags lstTracks.SelectedItem.Tag

    MsgBox "Finished.", vbInformation

End Sub

Private Sub cmdGrab_Click()

    Dim i   As Integer, j   As Integer

    On Error Resume Next
    For i = 1 To lstTracks.ListItems.count
        If lstTracks.ListItems(i).Checked Then
            j = j + 1
        End If
    Next
    On Error GoTo 0

    If j = 0 Then MsgBox "No tracks selected.": Exit Sub

    For i = 0 To udtTracks.count - 1

        With udtTracks.Track(i)

            If lstTracks.ListItems(i + 1).Checked = True Then

                ' check for albumnames
                If .Album = vbNullString Then
                    MsgBox "Albumname missing for track " & (i + 1), vbExclamation
                    lstTracks.ListItems(i + 1).Selected = True
                    ShowTrackTags i + 1
                    Exit Sub
                End If

                ' check for artistnames
                If .Artist = vbNullString Then
                    MsgBox "Artistname missing for track " & (i + 1), vbExclamation
                    lstTracks.ListItems(i + 1).Selected = True
                    ShowTrackTags i + 1
                    Exit Sub
                End If

                ' check for tracktitles
                If .Title = vbNullString Then
                    MsgBox "Title missing for track " & (i + 1), vbExclamation
                    lstTracks.ListItems(i + 1).Selected = True
                    ShowTrackTags i + 1
                    Exit Sub
                End If

            End If

        End With

    Next

    ' grab selected tracks
    grab

End Sub

Private Sub cmdPause_Click()
    cCDPlay.PauseResumeTrack
End Sub

Private Sub cmdPlay_Click()
    On Error GoTo ErrorHandler
    If lstTracks.SelectedItem.Tag > cCDPlay.TrackCount Then Exit Sub
    cCDPlay.PlayTrack lstTracks.SelectedItem.Tag
ErrorHandler:
End Sub

Private Sub cmdReset_Click()
    ' reset start/end position
    With udtTracks.Track(lstTracks.SelectedItem.Tag - 1)
        .startLBA = -150
        .endLBA = .lenLBA
    End With
    ShowTrackTags lstTracks.SelectedItem.Tag
End Sub

Private Sub cmdStop_Click()
    cCDPlay.StopTrack
End Sub

Private Sub cMonitor_arrival(ByVal drive As String)
    If LCase$(drive) = LCase$(cManager.DrvID2DrvChr(strDrvID)) Then
        cboDrv_Click
    End If
End Sub

Private Sub cMonitor_removal(ByVal drive As String)
    If LCase$(drive) = LCase$(cManager.DrvID2DrvChr(strDrvID)) Then
        cboDrv_Click
    End If
End Sub

Private Sub Form_Load()

    Set cFreeDB = New FL_FreeDB
    Set cMonitor = New FL_DoorMonitor
    Set cCDPlay = New FL_CDPlayer
    Set cGrab = New FL_TrackGrabber

    Me.Show: DoEvents

    AddBitrates
    ShowDrives

    txtPath = GetSetting("Flamedv4", "Grabber", "path", AddSlash(App.Path))
    cCDPlay.DigitalMode = CBool(GetSetting("Flamedv4", "Grabber", "playmode", False))
    cFreeDB.Timeout = GetSetting("Flamedv4", "Grabber", "timeout", 8)

    cboFilemask.ListIndex = 2
    cboSampleRate.ListIndex = 0

    cMonitor.InitDoorMonitor

End Sub

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

Private Sub ShowTrackTags(Track As Integer)
    Dim cMSF    As New FL_MSF

    With udtTracks.Track(Track - 1)

        txtAlbum = .Album
        txtArtist = .Artist
        txtTitle = .Title

        cMSF.LBA = .startLBA
        txtStartM = Format(cMSF.m, "00")
        txtStartS = Format(cMSF.s, "00")
        txtStartF = Format(cMSF.F, "00")

        cMSF.LBA = .endLBA
        txtEndM = Format(cMSF.m, "00")
        txtEndS = Format(cMSF.s, "00")
        txtEndF = Format(cMSF.F, "00")

        cMSF.LBA = .endLBA - .startLBA - 150
        lblLengthMSF = cMSF.MSF & " MSF"

    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cMonitor.DeInitDoorMonitor
    frmSelectProject.Show
End Sub

Private Sub lstTracks_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    udtTracks.Track(Item.Tag - 1).grab = Item.Checked
End Sub

Private Sub lstTracks_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ShowTrackTags Item.Tag
End Sub

Private Sub lstTracks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuRClick
End Sub

Private Sub mnuCheckAll_Click()
    Dim i   As Integer
    For i = 1 To lstTracks.ListItems.count
        lstTracks.ListItems(i).Checked = True
        udtTracks.Track(i - 1).grab = True
    Next
End Sub

Private Sub mnuUncheckAll_Click()
    Dim i   As Integer
    For i = 1 To lstTracks.ListItems.count
        lstTracks.ListItems(i).Checked = False
        udtTracks.Track(i - 1).grab = False
    Next
End Sub

Private Sub optWAV_Click()
    ChangeEncoder
End Sub

Private Sub optMP3_Click()
    ChangeEncoder
End Sub

Private Sub ChangeEncoder()
    If optMP3 Then
        chkID3v1.Enabled = True
        chkID3v2.Enabled = True
        cboBitrate.Enabled = True
        chkWrtHdr.Enabled = False
        cboSampleRate.Enabled = False
    Else
        chkID3v1.Enabled = False
        chkID3v2.Enabled = False
        cboBitrate.Enabled = False
        chkWrtHdr.Enabled = True
        cboSampleRate.Enabled = True
    End If
End Sub

Private Sub AddBitrates()
    ' add most compatible bitrates
    With cboBitrate
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
        .ListIndex = 0
    End With
End Sub

Public Sub grab()

    Dim i       As Integer
    Dim strMsg  As String

    Dim udeRet  As FL_SAVETRACK

    Me.Hide
    frmAudioGrabberPrg.Show

    ' get number of tracks to grab
    With frmAudioGrabberPrg

        .prgTotal.Max = 1
        For i = 1 To udtTracks.count
            If udtTracks.Track(i - 1).grab Then
                .prgTotal.Max = .prgTotal.Max + 1
            End If
        Next
        .prgTotal.Max = .prgTotal.Max - 1

        For i = 1 To udtTracks.count

            If udtTracks.Track(i - 1).grab Then

                With .lstStatus.ListItems.Add(SmallIcon:=1)
                    .SubItems(1) = "Grabbing track " & i
                    .Selected = True
                End With

                udeRet = GrabTrack(i)
                Select Case udeRet
                    Case ST_CANCELED: strMsg = "Canceled"
                    Case ST_ENCODER_INIT: strMsg = "Could not init encoder."
                    Case ST_FINISHED: strMsg = "Finished grabbing track " & i
                    Case ST_INVALID_SESSION: strMsg = "Invalid session."
                    Case ST_INVALID_TRACKMODE: strMsg = "Track " & i & " has invalid mode."
                    Case ST_INVALID_TRACKNO: strMsg = "Invalid track number: " & i
                    Case ST_NOT_READY: strMsg = "Drive not ready."
                    Case ST_READ_ERR: strMsg = "Read error."
                    Case ST_UNKNOWN_ERR: strMsg = "Unknown error occured."
                    Case ST_WRITE_ERR: strMsg = "Write error (HDD full?)"
                End Select

                With .lstStatus.ListItems.Add(SmallIcon:=1)
                    .SubItems(1) = strMsg
                    .Selected = True
                End With

                .prgTotal.Value = .prgTotal.Value + 1
    
                If udeRet = ST_CANCELED Or _
                   udeRet = ST_NOT_READY Or _
                   udeRet = ST_INVALID_SESSION Or _
                   udeRet = ST_ENCODER_INIT Then
                    Exit For
                End If

                ' write ID3v1 and ID3v2 tags to MP3 file
                If optMP3.Value Then
                    With udtTracks.Track(i - 1)
                        If chkID3v1 Then
                            cID3v1.MP3File = txtPath & GetTrackFileName(i)
                            cID3v1.Album = .Album
                            cID3v1.Artist = .Artist
                            cID3v1.Title = .Title
                            cID3v1.Track = .no
                            cID3v1.Year = Year(Date)
                            cID3v1.Comment = "Flamed v4"
                            cID3v1.Update
                        End If
                        If chkID3v2 Then
                            cID3v2.MP3File = txtPath & GetTrackFileName(i)
                            cID3v2.Album = .Album
                            cID3v2.Artist = .Artist
                            cID3v2.Title = .Title
                            cID3v2.Track = .no
                            cID3v2.Year = Year(Date)
                            cID3v2.Comment = "Flamed v4"
                            cID3v2.Update
                        End If
                    End With
                End If

            End If

        Next

    End With

    MsgBox "Finished", vbInformation, "OK"

    Unload frmAudioGrabberPrg
    Me.Show

End Sub

Private Function GetTrackFileName(ByVal Track As Integer) As String

    Dim strFile As String

    With udtTracks.Track(Track - 1)
        strFile = Replace(cboFilemask.List(cboFilemask.ListIndex), "<artist>", .Artist)
        strFile = Replace(strFile, "<album>", .Album)
        strFile = Replace(strFile, "<title>", .Title)
        strFile = Replace(strFile, "<num>", Format(.no, "00"))
        strFile = Replace(strFile, "<ext>", IIf(optMP3.Value, "mp3", "wav"))
    End With

    GetTrackFileName = strFile

End Function

Private Function GrabTrack(ByVal Track As Integer) As FL_SAVETRACK

    Dim udeRet      As FL_SAVETRACK
    Dim strFile     As String
    Dim lngStart    As Long
    Dim lngEnd      As Long

    strFile = GetTrackFileName(Track)

    lngStart = udtTracks.Track(Track - 1).startLBA + 150
    lngEnd = udtTracks.Track(Track - 1).endLBA

    If optMP3.Value Then
        udeRet = cGrab.AudioTrackToMP3(strDrvID, Track, txtPath & strFile, cboBitrate.ItemData(cboBitrate.ListIndex), lngStart, lngEnd)
    Else
        udeRet = cGrab.AudioTrackToWAV(strDrvID, Track, txtPath & strFile, chkWrtHdr, cboSampleRate.List(cboSampleRate.ListIndex), lngStart, lngEnd)
    End If

    GrabTrack = udeRet

End Function

Private Sub cGrab_Progress(ByVal Percent As Integer, ByVal Track As Integer, ByVal startLBA As Long, ByVal endLBA As Long, Cancel As Boolean)
    frmAudioGrabberPrg.prgTrack.Value = Percent
    Cancel = blnCancel
    DoEvents
End Sub

Private Sub prg_MouseDown(Shift As Integer)
    blnNoTmr = True
End Sub

Private Sub prg_MouseUp(Shift As Integer)
    cCDPlay.SeekTrack prg.Value
    blnNoTmr = False
End Sub

Private Sub txtAlbum_Change()
    udtTracks.Track(lstTracks.SelectedItem.Tag - 1).Album = txtAlbum
End Sub

Private Sub txtArtist_Change()
    udtTracks.Track(lstTracks.SelectedItem.Tag - 1).Artist = txtArtist
End Sub

Private Sub txtEndM_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    KeyAscii = KeyAscii * Abs(IsNumeric(Chr$(KeyAscii)))
End Sub

Private Sub txtEndS_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    KeyAscii = KeyAscii * Abs(IsNumeric(Chr$(KeyAscii)))
End Sub

Private Sub txtEndS_LostFocus()
    SetStartEndLBA
End Sub

Private Sub txtEndF_LostFocus()
    SetStartEndLBA
End Sub

Private Sub txtEndM_LostFocus()
    SetStartEndLBA
End Sub

Private Sub txtEndF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    KeyAscii = KeyAscii * Abs(IsNumeric(Chr$(KeyAscii)))
End Sub

Private Sub txtStartF_LostFocus()
    SetStartEndLBA
End Sub

Private Sub txtStartF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    KeyAscii = KeyAscii * Abs(IsNumeric(Chr$(KeyAscii)))
End Sub

Private Sub txtStartM_LostFocus()
    SetStartEndLBA
End Sub

Private Sub txtStartS_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    KeyAscii = KeyAscii * Abs(IsNumeric(Chr$(KeyAscii)))
End Sub

Private Sub txtStartS_LostFocus()
    SetStartEndLBA
End Sub

Private Sub txtStartM_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    KeyAscii = KeyAscii * Abs(IsNumeric(Chr$(KeyAscii)))
End Sub

Private Sub SetStartEndLBA()

    Dim cMSFStart   As New FL_MSF
    Dim cMSFEnd     As New FL_MSF

    cMSFStart.m = txtStartM
    cMSFStart.s = txtStartS
    cMSFStart.F = txtStartF

    cMSFEnd.m = txtEndM
    cMSFEnd.s = txtEndS
    cMSFEnd.F = txtEndF

    If cMSFStart.LBA >= cMSFEnd.LBA Then
        cMSFStart.LBA = cMSFEnd.LBA - 1
    End If

    udtTracks.Track(lstTracks.SelectedItem.Tag - 1).startLBA = cMSFStart.LBA
    cMSFStart.LBA = cMSFStart.LBA
    txtStartM = Format(cMSFStart.m, "00")
    txtStartS = Format(cMSFStart.s, "00")
    txtStartF = Format(cMSFStart.F, "00")

    udtTracks.Track(lstTracks.SelectedItem.Tag - 1).endLBA = cMSFEnd.LBA
    txtEndM = Format(cMSFEnd.m, "00")
    txtEndS = Format(cMSFEnd.s, "00")
    txtEndF = Format(cMSFEnd.F, "00")

    cMSFEnd.LBA = cMSFEnd.LBA - cMSFStart.LBA - 150

    lblLengthMSF = cMSFEnd.MSF & " MSF"

End Sub

Private Sub txtTitle_Change()
    udtTracks.Track(lstTracks.SelectedItem.Tag - 1).Title = txtTitle
End Sub
