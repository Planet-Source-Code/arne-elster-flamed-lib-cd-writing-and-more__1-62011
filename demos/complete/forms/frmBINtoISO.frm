VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBINtoISO 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "BIN to ISO converter"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
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
   ScaleHeight     =   255
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Default         =   -1  'True
      Height          =   315
      Left            =   4500
      TabIndex        =   13
      Top             =   3375
      Width           =   1440
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "Back"
      Height          =   330
      Left            =   225
      TabIndex        =   12
      Top             =   3375
      Width           =   1365
   End
   Begin VB.Frame frmPrg 
      Caption         =   "Progress"
      Height          =   690
      Left            =   127
      TabIndex        =   9
      Top             =   2550
      Width           =   5940
      Begin VB.PictureBox picPrg 
         BorderStyle     =   0  'Kein
         Height          =   390
         Left            =   75
         ScaleHeight     =   390
         ScaleWidth      =   5790
         TabIndex        =   10
         Top             =   225
         Width           =   5790
         Begin MSComctlLib.ProgressBar prg 
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   0
            Width           =   5640
            _ExtentX        =   9948
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   1
         End
      End
   End
   Begin VB.Frame frmOutput 
      Caption         =   "Output ISO"
      Height          =   690
      Left            =   127
      TabIndex        =   5
      Top             =   1650
      Width           =   5940
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Kein
         Height          =   465
         Left            =   75
         ScaleHeight     =   465
         ScaleWidth      =   5790
         TabIndex        =   6
         Top             =   195
         Width           =   5790
         Begin VB.TextBox txtISO 
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   75
            Width           =   4965
         End
         Begin VB.CommandButton cmdBrowseISO 
            Caption         =   "..."
            Height          =   285
            Left            =   5175
            TabIndex        =   7
            Top             =   75
            Width           =   420
         End
      End
   End
   Begin VB.Frame frmInput 
      Caption         =   "Input BIN"
      Height          =   690
      Left            =   127
      TabIndex        =   1
      Top             =   825
      Width           =   5940
      Begin VB.PictureBox picFile 
         BorderStyle     =   0  'Kein
         Height          =   465
         Left            =   75
         ScaleHeight     =   465
         ScaleWidth      =   5790
         TabIndex        =   2
         Top             =   195
         Width           =   5790
         Begin VB.TextBox txtBIN 
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   75
            Width           =   4965
         End
         Begin VB.CommandButton cmdBrowseBIN 
            Caption         =   "..."
            Height          =   285
            Left            =   5175
            TabIndex        =   3
            Top             =   75
            Width           =   420
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgISO 
      Left            =   5625
      Top             =   300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "ISO images (*.iso)|*.iso"
      Flags           =   2
   End
   Begin MSComDlg.CommonDialog dlgBIN 
      Left            =   4950
      Top             =   300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "BIN images (*.bin)|*.bin"
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BIN 2 ISO"
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
      Width           =   1260
   End
   Begin VB.Shape shpHdr 
      BackColor       =   &H00336471&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   7365
   End
End
Attribute VB_Name = "frmBINtoISO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cConv    As FL_ImageConverter
Attribute cConv.VB_VarHelpID = -1

Private blnCancel   As Boolean

Private Sub cConv_Progress(ByVal Percent As Integer, Cancel As Boolean)
    prg.Value = Percent
    Cancel = blnCancel
    DoEvents
End Sub

Private Sub cmdBack_Click()
    Me.Hide
    frmImgTools.Show
End Sub

Private Sub cmdBrowseBIN_Click()
    On Error GoTo ErrorHandler
    dlgBIN.ShowOpen
    txtBIN = dlgBIN.FileName
ErrorHandler:
End Sub

Private Sub cmdBrowseISO_Click()
    On Error GoTo ErrorHandler
    dlgISO.ShowSave
    txtISO = dlgISO.FileName
ErrorHandler:
End Sub

Private Sub cmdConvert_Click()

    Dim strMsg  As String

    If txtBIN = vbNullString Then
        MsgBox "No BIN file selected.", vbExclamation
        Exit Sub
    End If

    If txtISO = vbNullString Then
        MsgBox "No ISO file selected.", vbExclamation
        Exit Sub
    End If

    If cmdConvert.Caption = "Cancel" Then
        blnCancel = True
        Exit Sub
    End If

    cmdConvert.Caption = "Cancel"
    cmdBack.Enabled = Not cmdBack.Enabled
    
    blnCancel = False
    Select Case cConv.ConvertBINtoISO(txtBIN, txtISO)
        Case BIN2ISO_CANCELED: strMsg = "Canceled"
        Case BIN2ISO_INVALID_MODE: strMsg = "Only Mode-1 BIN supported"
        Case BIN2ISO_NOT_RAW: strMsg = "BIN is not raw or no image"
        Case BIN2ISO_OK: strMsg = "Finished"
        Case BIN2ISO_UNKNOWN: strMsg = "Unknown error"
    End Select

    MsgBox strMsg, vbInformation

    cmdConvert.Caption = "Convert"
    cmdBack.Enabled = Not cmdBack.Enabled

End Sub

Private Sub Form_Load()
    Set cConv = New FL_ImageConverter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmImgTools.Show
End Sub
