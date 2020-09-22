VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAudioGrabberPrg 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Flamed v4 - Audio CD Grabber Progress"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
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
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2610
      TabIndex        =   8
      Top             =   4290
      Width           =   1365
   End
   Begin MSComctlLib.ProgressBar prgTrack 
      Height          =   315
      Left            =   210
      TabIndex        =   7
      Top             =   3825
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar prgTotal 
      Height          =   315
      Left            =   210
      TabIndex        =   6
      Top             =   3225
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList img 
      Left            =   5775
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAudioGrabberPrg.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmStatus 
      Caption         =   "Status"
      Height          =   2040
      Left            =   135
      TabIndex        =   1
      Top             =   750
      Width           =   6315
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'Kein
         Height          =   1740
         Left            =   75
         ScaleHeight     =   116
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   411
         TabIndex        =   2
         Top             =   225
         Width           =   6165
         Begin MSComctlLib.ListView lstStatus 
            Height          =   1740
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   6165
            _ExtentX        =   10874
            _ExtentY        =   3069
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            _Version        =   393217
            SmallIcons      =   "img"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   8996
            EndProperty
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Track:"
      Height          =   195
      Left            =   225
      TabIndex        =   5
      Top             =   3600
      Width           =   450
   End
   Begin VB.Label lblStatTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   3000
      Width           =   420
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CDDA Grab Progress"
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
      Width           =   2580
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
Attribute VB_Name = "frmAudioGrabberPrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmAudioGrabber.Cancel = True
End Sub
