VERSION 5.00
Begin VB.Form frmAudioCDSettings 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Audio CD Writer Settings"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   88
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame frmSettings 
      Caption         =   "Settings"
      Height          =   690
      Left            =   75
      TabIndex        =   2
      Top             =   105
      Width           =   5115
      Begin VB.PictureBox picSettings 
         BorderStyle     =   0  'Kein
         Height          =   390
         Left            =   75
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   331
         TabIndex        =   3
         Top             =   225
         Width           =   4965
         Begin VB.CheckBox chkEjectDisk 
            Caption         =   "Eject disk after write"
            Height          =   195
            Left            =   75
            TabIndex        =   6
            Top             =   75
            Value           =   1  'Aktiviert
            Width           =   1815
         End
         Begin VB.ComboBox cboSpeed 
            Height          =   315
            Left            =   3075
            Style           =   2  'Dropdown-Liste
            TabIndex        =   4
            Top             =   30
            Width           =   1815
         End
         Begin VB.Label lblWriteSpeed 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            Caption         =   "Write speed:"
            Height          =   195
            Left            =   2100
            TabIndex        =   5
            Top             =   75
            Width           =   930
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   3750
      TabIndex        =   1
      Top             =   930
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2325
      TabIndex        =   0
      Top             =   930
      Width           =   1290
   End
End
Attribute VB_Name = "frmAudioCDSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cDrvNfo As New FL_DriveInfo

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
    ' &HFFFF& = max. speed
    cboSpeed.ItemData(cboSpeed.ListCount - 1) = &HFFFF&

    cboSpeed.ListIndex = cboSpeed.ListCount - 1

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Me.Hide

    cManager.SetCDRomSpeed strDrvID, &HFFFF&, cboSpeed.ItemData(cboSpeed.ListIndex)
    frmAudioCD.EjectDisk = chkEjectDisk
    frmAudioCD.Burn

End Sub

Private Sub Form_Load()
    ShowSpeeds
End Sub
