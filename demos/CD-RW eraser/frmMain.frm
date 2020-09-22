VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "CD-RW eraser"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
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
   ScaleHeight     =   187
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   193
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   315
      Left            =   1500
      TabIndex        =   4
      Top             =   1800
      Width           =   1290
   End
   Begin VB.ComboBox cboSpeed 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   750
      Style           =   2  'Dropdown-Liste
      TabIndex        =   3
      Top             =   525
      Width           =   2040
   End
   Begin VB.OptionButton optFullErase 
      Caption         =   "Full erase"
      Height          =   240
      Left            =   975
      TabIndex        =   2
      Top             =   1350
      Width           =   1815
   End
   Begin VB.OptionButton optQuickErase 
      Caption         =   "Quick erase"
      Height          =   240
      Left            =   975
      TabIndex        =   1
      Top             =   1050
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.DriveListBox cboDrv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   675
      TabIndex        =   0
      Top             =   150
      Width           =   2115
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "00:00"
      Height          =   195
      Left            =   75
      TabIndex        =   10
      Top             =   1965
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "expected time:"
      Height          =   195
      Left            =   75
      TabIndex        =   9
      Top             =   1770
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Method:"
      Height          =   195
      Left            =   75
      TabIndex        =   8
      Top             =   1050
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Speed:"
      Height          =   195
      Left            =   75
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Drive:"
      Height          =   165
      Left            =   75
      TabIndex        =   6
      Top             =   225
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Wait until you get notified of the end of the process!"
      Height          =   465
      Left            =   120
      TabIndex        =   5
      Top             =   2325
      Width           =   2715
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cManager        As New FL_Manager
Private cBlanker        As New FL_CDBlanker
Private cDriveInfo      As New FL_DriveInfo
Private cCDInfo         As New FL_CDInfo

Private strDrvID        As String

Private Sub cboDrv_Change()
    ' get new drive ID
    strDrvID = vbNullString
    cboSpeed.Clear
    If cManager.IsCDVDDrive(cboDrv.Drive) Then
        strDrvID = cManager.DrvChr2DrvID(cboDrv.Drive)
        ' show writing speeds for drive
        ShowSpeeds
    End If
End Sub

Private Sub ShowSpeeds()

    Dim i           As Integer
    Dim speeds()    As Integer

    If Not cDriveInfo.GetInfo(strDrvID) Then
        MsgBox "Could not get drive information.", vbExclamation
        Exit Sub
    End If

    ' writing supported?
    If cDriveInfo.WriteSpeedMax = 0 Then
        Exit Sub
    End If

    speeds = cDriveInfo.GetWriteSpeeds(strDrvID)

    For i = LBound(speeds) To UBound(speeds)
        cboSpeed.AddItem speeds(i) & " KB/s (" & speeds(i) \ 176 & "x)"
        cboSpeed.ItemData(cboSpeed.ListCount - 1) = speeds(i)
    Next

    cboSpeed.AddItem "max."
    cboSpeed.ItemData(cboSpeed.ListCount - 1) = &HFFFF&

    cboSpeed.ListIndex = cboSpeed.ListCount - 1

End Sub

Private Sub cmdStart_Click()

    Dim lngSpeed    As Long
    Dim intM        As Integer
    Dim intS        As Integer

    ' get CD info
    If Not cCDInfo.GetInfo(strDrvID) Then
        MsgBox "Could not get CD info.", vbExclamation
        Exit Sub
    End If

    ' check if medium is CD-RW
    If cCDInfo.MediaType <> ROMTYPE_CDRW Then
        MsgBox "No CD-RW inserted!", vbExclamation
        Exit Sub
    End If

    ' set speed (read = max)
    If Not cManager.SetCDRomSpeed(strDrvID, &HFFFF&, cboSpeed.ItemData(cboSpeed.ListIndex)) Then
        If MsgBox("Could not set write speed." & vbCrLf & "Continue?", vbExclamation Or vbYesNo, "Error") = vbNo Then
            Exit Sub
        End If
    End If

    cDriveInfo.GetInfo strDrvID
    lngSpeed = cDriveInfo.WriteSpeedCur

    If optFullErase Then
        intS = (cCDInfo.Capacity \ 1024) / lngSpeed
        intM = intS \ 60
        intS = intM * 60
    Else
        intM = 2
        intS = 0
    End If

    lblTime = Format(intM, "00") & ":" & Format(intS, "00")

    cmdStart.Enabled = Not cmdStart.Enabled
    optFullErase.Enabled = Not optFullErase.Enabled
    optQuickErase.Enabled = Not optQuickErase.Enabled
    cboDrv.Enabled = Not cboDrv.Enabled
    cboSpeed.Enabled = Not cboSpeed.Enabled

    If Not cBlanker.BlankCDRW(strDrvID, IIf(optFullErase.Value, BLANK_FULL, BLANK_QUICK), False) Then
        MsgBox "Failed.", vbExclamation
    Else
        MsgBox "Finished.", vbInformation
    End If

    cmdStart.Enabled = Not cmdStart.Enabled
    optFullErase.Enabled = Not optFullErase.Enabled
    optQuickErase.Enabled = Not optQuickErase.Enabled
    cboDrv.Enabled = Not cboDrv.Enabled
    cboSpeed.Enabled = Not cboSpeed.Enabled

End Sub

Private Sub Form_Load()
    If Not cManager.Init() Then
        MsgBox "No interfaces found.", vbExclamation
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cManager.Goodbye
End Sub
