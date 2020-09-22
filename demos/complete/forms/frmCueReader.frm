VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCueReader 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Cue Sheet Reader"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
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
   ScaleHeight     =   324
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   420
   StartUpPosition =   1  'Fenstermitte
   Begin MSComDlg.CommonDialog dlgBIN 
      Left            =   5175
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "BIN images (*.bin)|*.bin"
      Flags           =   2
   End
   Begin VB.Frame frmPrg 
      Caption         =   "Progress"
      Height          =   690
      Left            =   218
      TabIndex        =   8
      Top             =   3600
      Width           =   5865
      Begin VB.PictureBox picPrg 
         BorderStyle     =   0  'Kein
         Height          =   390
         Left            =   75
         ScaleHeight     =   390
         ScaleWidth      =   5715
         TabIndex        =   9
         Top             =   225
         Width           =   5715
         Begin MSComctlLib.ProgressBar prg 
            Height          =   315
            Left            =   90
            TabIndex        =   10
            Top             =   0
            Width           =   5565
            _ExtentX        =   9816
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   1
         End
      End
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "Back"
      Height          =   330
      Left            =   300
      TabIndex        =   7
      Top             =   4425
      Width           =   1365
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract track"
      Default         =   -1  'True
      Height          =   315
      Left            =   4575
      TabIndex        =   6
      Top             =   4425
      Width           =   1440
   End
   Begin MSComctlLib.TreeView tvwTracks 
      Height          =   1965
      Left            =   225
      TabIndex        =   5
      Top             =   1575
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   3466
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   471
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dlgCUE 
      Left            =   5700
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Cue sheets (*.cue)|*.cue"
      Flags           =   2
   End
   Begin VB.Frame frmFile 
      Caption         =   "Cue Sheet"
      Height          =   690
      Left            =   180
      TabIndex        =   1
      Top             =   750
      Width           =   5940
      Begin VB.PictureBox picFile 
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   75
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   386
         TabIndex        =   2
         Top             =   180
         Width           =   5790
         Begin VB.TextBox txtFile 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   75
            Width           =   4965
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5175
            TabIndex        =   3
            Top             =   75
            Width           =   420
         End
      End
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cue Sheet Reader"
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
      Width           =   2280
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
Attribute VB_Name = "frmCueReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cCue  As FL_CueReader
Attribute cCue.VB_VarHelpID = -1

Private blnCancel        As Boolean

Private Function TrackMode2Str(m As FL_TrackModes) As String
    Select Case m
        Case MODE_AUDIO: TrackMode2Str = "audio"
        Case MODE_MODE1: TrackMode2Str = "mode 1"
        Case MODE_MODE2: TrackMode2Str = "mode 2"
        Case MODE_MODE2_FORM1: TrackMode2Str = "mode 2 form 1"
        Case MODE_MODE2_FORM2: TrackMode2Str = "mode 2 form 2"
    End Select
End Function

Private Sub cCue_ExtractProgress(ByVal Percent As Integer, Cancel As Boolean)
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

    dlgCUE.ShowOpen
    txtFile = dlgCUE.FileName

    Select Case cCue.OpenCue(txtFile)

        Case CUE_BINARYEXPECTED, CUE_BINFILEEXPECTED, _
             CUE_CUEFILEEXPECTED, CUE_INDEXEXPECTED, _
             CUE_INDEXMSFEXPECTED, CUE_INDEXNUMEXPECTED, _
             CUE_TRACKEXPECTED, CUE_TRACKNUMEXPECTED:
            MsgBox "Invalid fields in cue sheet.", vbExclamation

        Case CUE_BINMISSING:
            MsgBox "BIN image missing.", vbExclamation

        Case CUE_OK:
            ShowTracks

    End Select

ErrorHandler:
End Sub

Private Sub ShowTracks()

    Dim i   As Integer, j   As Integer

    tvwTracks.Nodes.Clear

    For i = 1 To cCue.TrackCount

        tvwTracks.Nodes.Add(, , "trk" & i, "Track " & Format(i, "00") & " - " & TrackMode2Str(cCue.TrackMode(i))).Tag = i

        For j = 0 To cCue.TrackIndexCount(i) - 1
            tvwTracks.Nodes.Add("trk" & i, tvwChild, , "Index " & Format(j + cCue.TrackIndexFirst(i), "00") & " (" & cCue.TrackIndexLBA(i, j) & " LBA)").Tag = i
        Next

    Next

End Sub

Private Sub cmdExtract_Click()
    On Error GoTo ErrorHandler

    If cmdExtract.Caption = "Cancel" Then
        blnCancel = True
        Exit Sub
    End If

    dlgBIN.ShowSave

    cmdBack.Enabled = Not cmdBack.Enabled
    cmdExtract.Caption = "Cancel"

    If cCue.ExtractTrack(tvwTracks.SelectedItem.Tag, dlgBIN.FileName) Then
        MsgBox "Finished", vbInformation
    Else
        MsgBox "Failed (HDD full?)", vbExclamation
    End If

    cmdBack.Enabled = Not cmdBack.Enabled
    cmdExtract.Caption = "Extract track"

ErrorHandler:
End Sub

Private Sub Form_Load()
    Set cCue = New FL_CueReader
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmImgTools.Show
End Sub
