VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Cue reader"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
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
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   284
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComctlLib.ProgressBar prg 
      Height          =   315
      Left            =   825
      TabIndex        =   7
      Top             =   1875
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   525
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   3600
      Top             =   975
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Cue sheets (*.cue)|*.cue"
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   1875
      Width           =   1140
   End
   Begin VB.ListBox lstTracks 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   825
      TabIndex        =   4
      Top             =   975
      Width           =   2715
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   150
      Width           =   540
   End
   Begin VB.TextBox txtCue 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   675
      TabIndex        =   1
      Top             =   150
      Width           =   2865
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   3600
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   2
   End
   Begin VB.Label lblTracks 
      Caption         =   "Tracks:"
      Height          =   165
      Left            =   225
      TabIndex        =   3
      Top             =   975
      Width           =   540
   End
   Begin VB.Label lblCue 
      AutoSize        =   -1  'True
      Caption         =   "Cue:"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   345
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cCueReader As FL_CueReader
Attribute cCueReader.VB_VarHelpID = -1

Private Sub cCueReader_ExtractProgress(ByVal Percent As Integer, Cancel As Boolean)
    prg.Value = Percent
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo ErrorHandler

    dlg.ShowOpen
    txtCue = dlg.FileName

ErrorHandler:
End Sub

Private Sub cmdExtract_Click()

    On Error GoTo ErrorHandler
    dlgSave.Filter = "RAW files (*.bin)|*.bin"
    dlgSave.ShowSave
    On Error GoTo 0

    If Not cCueReader.ExtractTrack(lstTracks.ListIndex + 1, dlgSave.FileName) Then
        MsgBox "Failed.", vbExclamation
    Else
        MsgBox "Finished.", vbInformation
    End If

ErrorHandler:

End Sub

Private Sub cmdRead_Click()

    lstTracks.Clear

    Select Case cCueReader.OpenCue(txtCue)

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

End Sub

Private Sub ShowTracks()
    Dim i   As Integer

    For i = 1 To cCueReader.TrackCount
        lstTracks.AddItem "Track " & Format(i, "00")
    Next
End Sub

Private Sub Form_Load()
    Set cCueReader = New FL_CueReader
End Sub
