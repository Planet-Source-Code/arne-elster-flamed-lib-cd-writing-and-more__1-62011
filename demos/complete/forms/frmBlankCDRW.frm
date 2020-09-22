VERSION 5.00
Begin VB.Form frmBlankCDRW 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "CD-RW eraser"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   124
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.OptionButton optQuickErase 
      Caption         =   "Quick erase"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1125
      TabIndex        =   3
      Top             =   600
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton optFullErase 
      Caption         =   "Full erase"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1125
      TabIndex        =   2
      Top             =   900
      Width           =   1815
   End
   Begin VB.ComboBox cboSpeed 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   75
      Width           =   2040
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
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
      Left            =   1650
      TabIndex        =   0
      Top             =   1350
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Speed:"
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
      TabIndex        =   7
      Top             =   150
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Method:"
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
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "expected time:"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "00:00"
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
      TabIndex        =   4
      Top             =   1515
      Width           =   420
   End
End
Attribute VB_Name = "frmBlankCDRW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cBlanker        As New FL_CDBlanker
Private cDriveInfo      As New FL_DriveInfo
Private cCDInfo         As New FL_CDInfo

Private strPrvDrvID     As String

Public Property Let DriveID(aval As String)
    strPrvDrvID = aval
    ShowSpeeds
End Property

Private Sub ShowSpeeds()

    Dim i           As Integer
    Dim speeds()    As Integer

    If Not cDriveInfo.GetInfo(strPrvDrvID) Then
        MsgBox "Could not get drive information.", vbExclamation
        Exit Sub
    End If

    ' writing supported?
    If cDriveInfo.WriteSpeedMax = 0 Then
        Exit Sub
    End If

    speeds = cDriveInfo.GetWriteSpeeds(strPrvDrvID)

    For i = LBound(speeds) To UBound(speeds)
        cboSpeed.AddItem speeds(i) & " KB/s (" & speeds(i) \ 176 & "x)"
        cboSpeed.ItemData(cboSpeed.ListCount - 1) = speeds(i)
    Next

    cboSpeed.AddItem "max."
    cboSpeed.ItemData(cboSpeed.ListCount - 1) = &HFFFF&

    cboSpeed.ListIndex = cboSpeed.ListCount - 1

End Sub

Private Sub cboSpeed_Click()
    UpdateTime
End Sub

Private Sub cmdStart_Click()

    Dim lngSpeed    As Long
    Dim intM        As Integer
    Dim intS        As Integer

    ' get CD info
    If Not cCDInfo.GetInfo(strPrvDrvID) Then
        MsgBox "Could not get CD info.", vbExclamation
        Exit Sub
    End If

    ' check if medium is CD-RW
    If cCDInfo.MediaType <> ROMTYPE_CDRW Then
        MsgBox "No CD-RW inserted!", vbExclamation
        Exit Sub
    End If

    cmdStart.Enabled = Not cmdStart.Enabled
    optFullErase.Enabled = Not optFullErase.Enabled
    optQuickErase.Enabled = Not optQuickErase.Enabled
    cboSpeed.Enabled = Not cboSpeed.Enabled

    If Not cBlanker.BlankCDRW(strPrvDrvID, IIf(optFullErase.Value, BLANK_FULL, BLANK_QUICK), False) Then
        MsgBox "Failed.", vbExclamation
    Else
        MsgBox "Finished.", vbInformation
    End If

    cmdStart.Enabled = Not cmdStart.Enabled
    optFullErase.Enabled = Not optFullErase.Enabled
    optQuickErase.Enabled = Not optQuickErase.Enabled
    cboSpeed.Enabled = Not cboSpeed.Enabled

End Sub

Private Sub UpdateTime()

    Dim lngSpeed    As Long
    Dim intS        As Integer
    Dim intM        As Integer

    ' get CD info
    If Not cCDInfo.GetInfo(strPrvDrvID) Then
        MsgBox "Could not get CD info.", vbExclamation
        Exit Sub
    End If

    ' set speed (read = max)
    cManager.SetCDRomSpeed strPrvDrvID, &HFFFF&, cboSpeed.ItemData(cboSpeed.ListIndex)

    cDriveInfo.GetInfo strPrvDrvID
    lngSpeed = cDriveInfo.WriteSpeedCur

    If optFullErase Then
        intS = ((cCDInfo.Capacity \ 1024) / lngSpeed) * 2
        intM = intS \ 60
        intS = intS - (intM * 60)
    Else
        intM = 2
        intS = 0
    End If

    lblTime = Format(intM, "00") & ":" & Format(intS, "00")

End Sub

Private Sub optFullErase_Click()
    UpdateTime
End Sub

Private Sub optQuickErase_Click()
    UpdateTime
End Sub
