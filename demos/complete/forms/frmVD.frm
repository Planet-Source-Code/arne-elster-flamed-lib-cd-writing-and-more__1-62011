VERSION 5.00
Begin VB.Form frmVD 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Volume Descriptors"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox txtVolID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1425
      TabIndex        =   5
      Top             =   225
      Width           =   2790
   End
   Begin VB.TextBox txtSysID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1425
      TabIndex        =   4
      Top             =   675
      Width           =   2790
   End
   Begin VB.TextBox txtPubID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1425
      TabIndex        =   3
      Top             =   1125
      Width           =   2790
   End
   Begin VB.TextBox txtAppID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1410
      TabIndex        =   2
      Top             =   1575
      Width           =   2790
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2850
      TabIndex        =   1
      Top             =   2100
      Width           =   1290
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1425
      TabIndex        =   0
      Top             =   2100
      Width           =   1290
   End
   Begin VB.Label lblVolumeID 
      AutoSize        =   -1  'True
      Caption         =   "Volume ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   510
      TabIndex        =   9
      Top             =   300
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "System ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   495
      TabIndex        =   8
      Top             =   750
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Publisher ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   7
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Application ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   1650
      Width           =   1050
   End
End
Attribute VB_Name = "frmVD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    frmDataCD.AppID = txtAppID
    frmDataCD.VolumeID = txtVolID
    frmDataCD.PublisherID = txtPubID
    frmDataCD.SystemID = txtSysID

    Unload Me

End Sub

Private Sub cmdReset_Click()
    Form_Load
End Sub

Private Sub Form_Load()
    txtAppID = frmDataCD.AppID
    txtVolID = frmDataCD.VolumeID
    txtPubID = frmDataCD.PublisherID
    txtSysID = frmDataCD.SystemID
End Sub
