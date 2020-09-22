VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "CD Information"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
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
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox lst 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3180
      Left            =   173
      TabIndex        =   1
      Top             =   600
      Width           =   5115
   End
   Begin VB.DriveListBox cboDrv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   173
      TabIndex        =   0
      Top             =   225
      Width           =   5115
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cManager        As New FL_Manager
Private cInfo           As New FL_CDInfo

Private strDrvID        As String

Private Sub cboDrv_Change()
    lst.Clear
    If cManager.IsCDVDDrive(cboDrv.Drive) Then
        strDrvID = cManager.DrvChr2DrvID(cboDrv.Drive)
        Info
    End If
End Sub

Sub Info()
    
    If Not cInfo.GetInfo(strDrvID) Then
        MsgBox "Failed to read CD information.", vbExclamation, "Error"
        Exit Sub
    End If

    With cInfo
        lst.AddItem "Capacity: " & (.Capacity \ 1024 ^ 2) & " MB"
        lst.AddItem "Type: " & CDTypeToStr(.MediaType)
        lst.AddItem "CD-(R(W)) Type: " & STypeToStr(.CDRWType)
        lst.AddItem "Vendor: " & .CDRWVendor
        lst.AddItem "Erasable: " & .Erasable
        lst.AddItem "Last session's state: " & Status2Str(.LastSessionState)
        lst.AddItem "Lead-In: " & .LeadInMSF.MSF & " MSF (" & .LeadInMSF.LBA & " LBA)"
        lst.AddItem "Last possible Lead-Out start: " & .LeadOutMSF.MSF & " MSF (" & .LeadOutMSF.LBA & " LBA)"
        lst.AddItem "Media status:" & Status2Str(.MediaStatus)
        lst.AddItem "Sessions: " & .Sessions
        lst.AddItem "Tracks: " & .Tracks
        lst.AddItem "Size: " & (.Size \ 1024 ^ 2) & " MB"
    End With

End Sub

Private Function Status2Str(s As FL_Status) As String
    Select Case s
        Case STAT_COMPLETE: Status2Str = "complete"
        Case STAT_EMPTY: Status2Str = "empty"
        Case STAT_INCOMPLETE: Status2Str = "incomplete"
        Case STAT_UNKNWN: Status2Str = "unknown"
    End Select
End Function

Private Function STypeToStr(s As FL_CDSubType) As String
    Select Case s
        Case STYPE_CDI: STypeToStr = "CD-I"
        Case STYPE_CDROMDA: STypeToStr = "CD-ROM/CDDA"
        Case STYPE_UNKNWN: STypeToStr = "Unknown"
        Case STYPE_XA: STypeToStr = "CD-XA"
    End Select
End Function

Private Function CDTypeToStr(s As FL_CDType) As String
    Select Case s
        Case ROMTYPE_CDR: CDTypeToStr = "CD-R"
        Case ROMTYPE_CDROM: CDTypeToStr = "CD-ROM"
        Case ROMTYPE_CDROM_R_RW: CDTypeToStr = "CD-ROM/R/RW"
        Case ROMTYPE_CDRW: CDTypeToStr = "CD-RW"
        Case ROMTYPE_DVD_P_R: CDTypeToStr = "DVD+R"
        Case ROMTYPE_DVD_P_RW: CDTypeToStr = "DVD+RW"
        Case ROMTYPE_DVD_R: CDTypeToStr = "DVD-R"
        Case ROMTYPE_DVD_RAM: CDTypeToStr = "DVD-RAM"
        Case ROMTYPE_DVD_ROM: CDTypeToStr = "DVD-ROM"
        Case ROMTYPE_DVD_RW: CDTypeToStr = "DVD-RW"
    End Select
End Function

Private Sub Form_Load()
    If Not cManager.Init(False) Then
        MsgBox "No interface found.", vbExclamation, "Error"
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cManager.Goodbye
End Sub
