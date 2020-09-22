VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "BIN2ISO Converter"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
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
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   379
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Default         =   -1  'True
      Height          =   315
      Left            =   4650
      TabIndex        =   7
      Top             =   1350
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   315
      Left            =   225
      TabIndex        =   6
      Top             =   1350
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4950
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   2
   End
   Begin VB.CommandButton cmdBrowseISO 
      Caption         =   "..."
      Height          =   315
      Left            =   4650
      TabIndex        =   5
      Top             =   750
      Width           =   390
   End
   Begin VB.TextBox txtISO 
      Height          =   315
      Left            =   975
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   750
      Width           =   3615
   End
   Begin VB.CommandButton cmdBrowseBIN 
      Caption         =   "..."
      Height          =   315
      Left            =   4650
      TabIndex        =   2
      Top             =   255
      Width           =   390
   End
   Begin VB.TextBox txtBIN 
      Height          =   315
      Left            =   975
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   255
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "ISO file:"
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   795
      Width           =   720
   End
   Begin VB.Label lblBIN 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "BIN file:"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   300
      Width           =   705
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cConvert As FL_ImageConverter
Attribute cConvert.VB_VarHelpID = -1

Private blnCancel           As Boolean

Private Sub cConvert_Progress(ByVal Percent As Integer, Cancel As Boolean)
    prg.Value = Percent
    Cancel = blnCancel
End Sub

Private Sub cmdBrowseBIN_Click()

    On Error GoTo ErrorHandler

    dlg.Filter = "BIN images (*.bin)|*.bin"
    dlg.ShowOpen
    txtBIN = dlg.FileName

ErrorHandler:

End Sub

Private Sub cmdBrowseISO_Click()

    On Error GoTo ErrorHandler

    dlg.Filter = "ISO images (*.iso)|*.iso"
    dlg.ShowSave
    txtISO = dlg.FileName

ErrorHandler:

End Sub

Private Sub cmdConvert_Click()

    If cmdConvert.Caption = "Cancel" Then
        blnCancel = True
        Exit Sub
    End If

    cmdConvert.Caption = "Cancel"

    cmdBrowseBIN.Enabled = Not cmdBrowseBIN.Enabled
    cmdBrowseISO.Enabled = Not cmdBrowseISO.Enabled

    Select Case cConvert.ConvertBINtoISO(txtBIN, txtISO)
        Case BIN2ISO_INVALID_MODE: MsgBox "Invalid track mode."
        Case BIN2ISO_NOT_RAW: MsgBox "Either Audio/Mode-0 or BIN is not raw."
        Case BIN2ISO_OK: MsgBox "Finished."
        Case BIN2ISO_UNKNOWN: MsgBox "Unknown error."
        Case BIN2ISO_CANCELED: MsgBox "Canceled."
    End Select

    cmdBrowseBIN.Enabled = Not cmdBrowseBIN.Enabled
    cmdBrowseISO.Enabled = Not cmdBrowseISO.Enabled

    cmdConvert.Caption = "Convert"

End Sub
 
Private Sub Form_Load()
    Set cConvert = New FL_ImageConverter
End Sub
