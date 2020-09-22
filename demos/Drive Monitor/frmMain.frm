VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Drive Monitor"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdLoUnlock 
      Caption         =   "Lock/Unlock"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3300
      TabIndex        =   3
      Top             =   3225
      Width           =   1215
   End
   Begin VB.CommandButton cmdEjLo 
      Caption         =   "Eject/Load"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   2
      Top             =   2850
      Width           =   1215
   End
   Begin VB.DriveListBox cboDrv 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   2850
      Width           =   3015
   End
   Begin VB.ListBox lst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4365
   End
   Begin VB.Label lblStat 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Top             =   3225
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cMonitor As FL_DoorMonitor
Attribute cMonitor.VB_VarHelpID = -1

Private cManager            As New FL_Manager
Private cDriveInfo          As New FL_DriveInfo

Private strDriveID          As String

Private Sub cboDrv_Change()
    strDriveID = vbNullString
    ' CD/DVD-ROM drive?
    If cManager.IsCDVDDrive(cboDrv.drive) Then
        ' get drive ID
        strDriveID = cManager.DrvChr2DrvID(cboDrv.drive)
        DrvNfo
    Else
        lblStat = vbNullString
    End If
End Sub

Private Sub cmdEjLo_Click()

    ' can we determine the drive's state?
    If Not cDriveInfo.GetInfo(strDriveID) Then

        ' no, just try to open it
        MsgBox "Could not read drive information." & vbCrLf & _
               "Will just try to open the drive.", vbExclamation, "error"

        If Not cManager.UnLoadDrive(strDriveID) Then
            MsgBox "Failed to open the drive.", vbExclamation, "Error"
        End If

        Exit Sub

    End If

    ' drive closed?
    If cDriveInfo.DriveClosed Then
        ' open
        If Not cManager.UnLoadDrive(strDriveID) Then
            MsgBox "Failed to open the drive.", vbExclamation, "Error"
        End If
    Else
        ' close
        If Not cManager.LoadDrive(strDriveID) Then
            MsgBox "Failed to close the drive.", vbExclamation, "Error"
        End If
    End If

    ' show drive's state
    DrvNfo

End Sub

Private Sub cmdLoUnlock_Click()

    ' can we determine the drive's state?
    If Not cDriveInfo.GetInfo(strDriveID) Then

        ' no, just try to unlock the drive
        ' as we can't determine wether the
        ' drive is locked or not.
        ' Would be bad if we nverthless
        ' locked it and the method wouldn't
        ' allow to unlock it.
        MsgBox "Could not read drive information." & vbCrLf & _
               "Will just try to unlock the drive.", vbExclamation, "error"

        If Not cManager.UnLockDrive(strDriveID) Then
            MsgBox "Failed to unlock the drive.", vbExclamation, "Error"
        End If

        Exit Sub

    End If

    If cDriveInfo.DriveLocked Then
        If Not cManager.UnLockDrive(strDriveID) Then
            MsgBox "Failed to unlock the drive.", vbExclamation, "Error"
        End If
    Else
        If Not cManager.LockDrive(strDriveID) Then
            MsgBox "Failed to lock the drive.", vbExclamation, "Error"
        End If
    End If

    DrvNfo

End Sub

Private Sub cMonitor_arrival(ByVal drive As String)
    lst.AddItem "New Media in " & drive & ":"
End Sub

Private Sub cMonitor_removal(ByVal drive As String)
    lst.AddItem "Media removed from " & drive & ":"
End Sub

Private Sub DrvNfo()
    If cDriveInfo.GetInfo(strDriveID) Then
        lblStat = "Closed: " & cDriveInfo.DriveClosed
        lblStat = lblStat & "   " & "Locked: " & cDriveInfo.DriveLocked
    End If
End Sub

Private Sub Form_Load()
    Set cMonitor = New FL_DoorMonitor

    If Not cManager.Init(False) Then
        MsgBox "No interfaces found.", vbExclamation, "Error"
        Unload Me
    End If

    cMonitor.InitDoorMonitor

End Sub

Private Sub Form_Unload(Cancel As Integer)
    cMonitor.DeInitDoorMonitor
    cManager.Goodbye
End Sub
