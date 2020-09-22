VERSION 5.00
Begin VB.Form frmImgTools 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Image Tools"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4965
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   3600
      TabIndex        =   3
      Top             =   3975
      Width           =   1290
   End
   Begin prjFlamed.ucCoolList lstPrjs 
      Height          =   3060
      Left            =   75
      TabIndex        =   1
      Top             =   825
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5398
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackSelected    =   5616590
      BackSelectedG1  =   5616590
      BoxRadius       =   5
      Focus           =   0   'False
      ItemHeight      =   40
      ItemHeightAuto  =   0   'False
      ItemOffset      =   6
      SelectModeStyle =   4
   End
   Begin VB.Label lblPrj 
      AutoSize        =   -1  'True
      Caption         =   "Please select a project from the list:"
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
      Left            =   75
      TabIndex        =   2
      Top             =   600
      Width           =   2580
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image tools"
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
      Left            =   225
      TabIndex        =   0
      Top             =   120
      Width           =   1470
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
Attribute VB_Name = "frmImgTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Select Case lstPrjs.ListIndex
        Case 0: Me.Hide: frmBurnISO.Show vbModal, frmSelectProject
        Case 1: Me.Hide: frmDataToISO.Show vbModal, frmSelectProject
        Case 2: Me.Hide: frmSessToBIN.Show vbModal, frmSelectProject
        Case 3: Me.Hide: frmCueReader.Show vbModal, frmSelectProject
        Case 4: Me.Hide: frmBINtoISO.Show vbModal, frmSelectProject
    End Select
End Sub

Private Sub Form_Load()
    With lstPrjs
        .AddItem "Burn ISO image"
        .AddItem "Data track to ISO"
        .AddItem "Session to BIN/CUE"
        .AddItem "Extract tracks from BIN/CUE"
        .AddItem "Convert BIN to ISO"
    End With
End Sub

Private Sub lstPrjs_DblClick()
    cmdOK_Click
End Sub
