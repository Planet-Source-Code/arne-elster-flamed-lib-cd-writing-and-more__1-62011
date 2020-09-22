VERSION 5.00
Begin VB.Form frmISODescriptors 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "ISO Descriptors"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4350
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
   ScaleHeight     =   2445
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Top             =   2025
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2925
      TabIndex        =   8
      Top             =   2025
      Width           =   1290
   End
   Begin VB.TextBox txtAppID 
      Height          =   315
      Left            =   1410
      TabIndex        =   7
      Top             =   1500
      Width           =   2790
   End
   Begin VB.TextBox txtPubID 
      Height          =   315
      Left            =   1425
      TabIndex        =   5
      Top             =   1050
      Width           =   2790
   End
   Begin VB.TextBox txtSysID 
      Height          =   315
      Left            =   1425
      TabIndex        =   3
      Top             =   600
      Width           =   2790
   End
   Begin VB.TextBox txtVolID 
      Height          =   315
      Left            =   1425
      TabIndex        =   1
      Top             =   150
      Width           =   2790
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Application ID:"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   1575
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Publisher ID:"
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   1125
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "System ID:"
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   675
      Width           =   990
   End
   Begin VB.Label lblVolumeID 
      AutoSize        =   -1  'True
      Caption         =   "Volume ID:"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   225
      Width           =   990
   End
End
Attribute VB_Name = "frmISODescriptors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    frmMain.AppID = txtAppID
    frmMain.VolumeID = txtVolID
    frmMain.PublisherID = txtPubID
    frmMain.SystemID = txtSysID

    Unload Me

End Sub

Private Sub cmdReset_Click()
    Form_Load
End Sub

Private Sub Form_Load()
    txtAppID = frmMain.AppID
    txtVolID = frmMain.VolumeID
    txtPubID = frmMain.PublisherID
    txtSysID = frmMain.SystemID
End Sub
