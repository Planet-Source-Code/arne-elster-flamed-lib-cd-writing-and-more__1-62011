VERSION 5.00
Begin VB.Form frmSelectProject 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Flamed v4"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
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
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   315
      Left            =   150
      TabIndex        =   5
      Top             =   4425
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4950
      TabIndex        =   4
      Top             =   4425
      Width           =   1140
   End
   Begin prjFlamed.ucCoolList lstPrjs 
      Height          =   3060
      Left            =   0
      TabIndex        =   2
      Top             =   1275
      Width           =   6240
      _ExtentX        =   11007
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
      Height          =   195
      Left            =   75
      TabIndex        =   3
      Top             =   1050
      Width           =   2580
   End
   Begin VB.Label lblOpenSource 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Source CD Writing Software"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   2925
   End
   Begin VB.Label lblHdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flamed v4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   2355
   End
   Begin VB.Shape shpHdr 
      BackColor       =   &H00336471&
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   990
      Left            =   0
      Top             =   0
      Width           =   6315
   End
End
Attribute VB_Name = "frmSelectProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Select Case lstPrjs.ListIndex
        Case 0: frmDataCD.Show: Me.Hide
        Case 1: frmAudioCD.Show: Me.Hide
        Case 2: frmAudioGrabber.Show: Me.Hide
        Case 3: frmImgTools.Show vbModal, Me
        Case 4: frmOptions.Show vbModal, Me
    End Select

End Sub

Private Sub Form_Load()
    With lstPrjs
        .AddItem "Data CD project"
        .AddItem "Audio CD project"
        .AddItem "Audio CD Grabber"
        .AddItem "Image tools"
        .AddItem "Options"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If MsgBox("Quit Flamed v4?", vbYesNo Or vbQuestion, "Quit?") = vbNo Then
        Cancel = 1
        Exit Sub
    End If

    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next frm

End Sub

Private Sub lstPrjs_DblClick()
    cmdOK_Click
End Sub
