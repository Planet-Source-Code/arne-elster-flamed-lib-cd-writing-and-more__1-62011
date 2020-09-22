VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Options"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6210
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
   ScaleHeight     =   196
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   414
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox picCDDAGrab 
      BorderStyle     =   0  'Kein
      Height          =   2265
      Left            =   150
      ScaleHeight     =   2265
      ScaleWidth      =   5865
      TabIndex        =   9
      Top             =   450
      Visible         =   0   'False
      Width           =   5865
      Begin VB.TextBox txtTimeout 
         Height          =   315
         Left            =   1500
         TabIndex        =   15
         Text            =   "8"
         Top             =   840
         Width           =   540
      End
      Begin VB.CheckBox chkDigitalPlayback 
         Caption         =   "Digital CD playback"
         Height          =   240
         Left            =   75
         TabIndex        =   13
         Top             =   525
         Width           =   2115
      End
      Begin VB.CommandButton cmdBrowseGrabPath 
         Caption         =   "..."
         Height          =   285
         Left            =   5325
         TabIndex        =   11
         Top             =   105
         Width           =   390
      End
      Begin VB.TextBox txtGrabPath 
         Height          =   285
         Left            =   1425
         TabIndex        =   10
         Top             =   105
         Width           =   3840
      End
      Begin VB.Label lblS 
         Caption         =   "s"
         Height          =   240
         Left            =   2175
         TabIndex        =   16
         Top             =   900
         Width           =   240
      End
      Begin VB.Label lblFreeDBTimeout 
         AutoSize        =   -1  'True
         Caption         =   "FreeDB Timeout:"
         Height          =   195
         Left            =   75
         TabIndex        =   14
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Default directory:"
         Height          =   195
         Left            =   75
         TabIndex        =   12
         Top             =   135
         Width           =   1275
      End
   End
   Begin VB.PictureBox picAudioCD 
      BorderStyle     =   0  'Kein
      Height          =   2265
      Left            =   150
      ScaleHeight     =   2265
      ScaleWidth      =   5865
      TabIndex        =   5
      Top             =   450
      Visible         =   0   'False
      Width           =   5865
      Begin VB.TextBox txtAudioCDTemp 
         Height          =   285
         Left            =   750
         TabIndex        =   7
         Top             =   120
         Width           =   4365
      End
      Begin VB.CommandButton cmdBrowseAudioCDTemp 
         Caption         =   "..."
         Height          =   285
         Left            =   5175
         TabIndex        =   6
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Temp:"
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   150
         Width           =   450
      End
   End
   Begin VB.PictureBox picDataCD 
      BorderStyle     =   0  'Kein
      Height          =   2265
      Left            =   150
      ScaleHeight     =   2265
      ScaleWidth      =   5865
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   5865
      Begin VB.CommandButton cmdBrowseDataCDTemp 
         Caption         =   "..."
         Height          =   285
         Left            =   5175
         TabIndex        =   4
         Top             =   120
         Width           =   390
      End
      Begin VB.TextBox txtDataCDTemp 
         Height          =   285
         Left            =   750
         TabIndex        =   3
         Top             =   120
         Width           =   4365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Temp:"
         Height          =   195
         Left            =   225
         TabIndex        =   2
         Top             =   150
         Width           =   450
      End
   End
   Begin MSComctlLib.TabStrip tabstrip 
      Height          =   2790
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   4921
      TabMinWidth     =   2587
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Data CD Writer"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Audio CD Writer"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CDDA Grabber"
            ImageVarType    =   2
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
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cDataCD     As New FL_CDDataWriter
Private cAudioCD    As New FL_CDAudioWriter

Private Sub chkDigitalPlayback_Click()
    SaveSetting "Flamedv4", "Grabber", "playmode", CBool(chkDigitalPlayback)
End Sub

Private Sub cmdBrowseAudioCDTemp_Click()

    Dim strText As String

    strText = BrowseForFolder("Please select a new temp dir", txtAudioCDTemp, hWnd, True, , True)

    If strText <> vbNullString Then
        txtAudioCDTemp = AddSlash(strText)
        SaveSetting "Flamedv4", "AudioCD", "temp", txtAudioCDTemp
    End If

End Sub

Private Sub cmdBrowseDataCDTemp_Click()

    Dim strText As String

    strText = BrowseForFolder("Please select a new temp dir", txtDataCDTemp, hWnd, True, , True)

    If strText <> vbNullString Then
        txtDataCDTemp = AddSlash(strText)
        SaveSetting "Flamedv4", "DataCD", "temp", txtDataCDTemp
    End If

End Sub

Private Sub cmdBrowseGrabPath_Click()

    Dim strText As String

    strText = BrowseForFolder("Please select a new default dir", txtGrabPath, hWnd, True, , True)

    If strText <> vbNullString Then
        txtGrabPath = AddSlash(strText)
        SaveSetting "Flamedv4", "Grabber", "path", txtGrabPath
    End If

End Sub

Private Sub Form_Load()

    txtDataCDTemp = GetSetting("Flamedv4", "DataCD", "temp", cDataCD.TempDir)
    txtAudioCDTemp = GetSetting("Flamedv4", "AudioCD", "temp", cAudioCD.TempDir)
    txtGrabPath = GetSetting("Flamedv4", "Grabber", "path", AddSlash(App.Path))
    chkDigitalPlayback = Abs(CBool(GetSetting("Flamedv4", "Grabber", "playmode", 0)))
    txtTimeout = GetSetting("Flamedv4", "Grabber", "timeout", 8)

    tabstrip.TabIndex = 1
    tabstrip_Click

End Sub

Private Sub tabstrip_Click()
    Select Case tabstrip.SelectedItem.index
        Case 1
            picDataCD.Visible = True
            picAudioCD.Visible = False
            picCDDAGrab.Visible = False
        Case 2
            picDataCD.Visible = False
            picAudioCD.Visible = True
            picCDDAGrab.Visible = False
        Case 3
            picDataCD.Visible = False
            picAudioCD.Visible = False
            picCDDAGrab.Visible = True
    End Select
End Sub

Private Sub txtTimeout_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SaveSetting "Flamedv4", "Grabber", "timeout", txtTimeout
        KeyAscii = 0
    End If
End Sub

Private Sub txtTimeout_LostFocus()
    SaveSetting "Flamedv4", "Grabber", "timeout", txtTimeout
End Sub
