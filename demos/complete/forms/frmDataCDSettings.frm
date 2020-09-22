VERSION 5.00
Begin VB.Form frmDataCDSettings 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Data CD Writer Settings"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   ControlBox      =   0   'False
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
   ScaleHeight     =   143
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2400
      TabIndex        =   8
      Top             =   1755
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   3825
      TabIndex        =   7
      Top             =   1755
      Width           =   1440
   End
   Begin VB.Frame frmSettings 
      Caption         =   "Settings"
      Height          =   1515
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5115
      Begin VB.PictureBox picSettings 
         BorderStyle     =   0  'Kein
         Height          =   1215
         Left            =   75
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   331
         TabIndex        =   1
         Top             =   225
         Width           =   4965
         Begin VB.CheckBox chkOnTheFly 
            Caption         =   "On The Fly"
            Height          =   195
            Left            =   75
            TabIndex        =   9
            Top             =   90
            Value           =   1  'Aktiviert
            Width           =   1815
         End
         Begin VB.CheckBox chkFinalize 
            Caption         =   "Finalize disk"
            Height          =   195
            Left            =   75
            TabIndex        =   6
            Top             =   900
            Value           =   1  'Aktiviert
            Width           =   1815
         End
         Begin VB.CheckBox chkEjectDisk 
            Caption         =   "Eject disk after write"
            Height          =   195
            Left            =   75
            TabIndex        =   5
            Top             =   356
            Value           =   1  'Aktiviert
            Width           =   1815
         End
         Begin VB.CheckBox chkTestmode 
            Caption         =   "Test mode"
            Height          =   195
            Left            =   75
            TabIndex        =   4
            Top             =   630
            Width           =   1740
         End
         Begin VB.ComboBox cboSpeed 
            Height          =   315
            Left            =   3075
            Style           =   2  'Dropdown-Liste
            TabIndex        =   3
            Top             =   30
            Width           =   1815
         End
         Begin VB.Label lblWriteSpeed 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            Caption         =   "Write speed:"
            Height          =   195
            Left            =   2100
            TabIndex        =   2
            Top             =   75
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "frmDataCDSettings"
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
    cboSpeed.ItemData(cboSpeed.ListCount - 1) = &HFFFF&

    cboSpeed.ListIndex = cboSpeed.ListCount - 1

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Me.Hide

    cManager.SetCDRomSpeed strDrvID, &HFFFF&, cboSpeed.ItemData(cboSpeed.ListIndex)

    frmDataCD.OnTheFly = chkOnTheFly
    frmDataCD.Finalize = chkFinalize
    frmDataCD.TestMode = chkTestmode
    frmDataCD.EjectDisk = chkEjectDisk

    frmDataCD.Burn

End Sub

Private Sub Form_Load()
    ShowSpeeds
End Sub
