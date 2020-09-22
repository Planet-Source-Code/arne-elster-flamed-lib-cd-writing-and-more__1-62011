VERSION 5.00
Begin VB.Form frmInit 
   BorderStyle     =   0  'Kein
   Caption         =   "Flamed v4"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
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
   Picture         =   "frmInit.frx":0000
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Searching for interfaces..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0074CDCD&
      Height          =   210
      Left            =   2025
      TabIndex        =   0
      Top             =   3795
      Width           =   2115
   End
End
Attribute VB_Name = "frmInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim cDrvNfo         As New FL_DriveInfo
    Dim blnForceASPI    As Boolean
    Dim strDrives()     As String
    Dim i               As Long

    Me.Show

    ShowMsg "Searching for usable interface...", 500

    blnForceASPI = Command = "-aspi"
    If Not cManager.Init(blnForceASPI) Then
        MsgBox "No interfaces found!" & vbCrLf & _
               "Please install an ASPI driver!" & vbCrLf & _
               "App will exit.", vbExclamation, "Error"
        Unload Me
    End If

    ShowMsg "Used interface: " & Choose(cManager.CurrentInterface, "ASPI", "SPTI"), 500
    ShowMsg "Searching for drives...", 700

    strDrives = cManager.GetCDVDROMs()
    For i = LBound(strDrives) To UBound(strDrives) - 1
        With cDrvNfo
            .GetInfo cManager.DrvChr2DrvID(strDrives(i))
            If (.ReadCapabilities And RC_DVDROM) Then
                ShowMsg strDrives(i) & ": " & .Vendor & " " & .Product & " " & .Revision & " (DVD)", 400
            Else
                ShowMsg strDrives(i) & ": " & .Vendor & " " & .Product & " " & .Revision & " (CD)", 400
            End If
        End With
    Next

    ShowMsg "Ready.", 1000
    frmSelectProject.Show
    Unload Me

End Sub

Private Sub ShowMsg(msg As String, ms As Long)
    lblStatus = msg
    DoEvents
    Sleep ms
End Sub
