VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "CDB gate"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   148
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.DriveListBox cboDrv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   383
      TabIndex        =   2
      Top             =   150
      Width           =   3915
   End
   Begin VB.ListBox lstRes 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   683
      TabIndex        =   1
      Top             =   1200
      Width           =   3315
   End
   Begin VB.CommandButton cmdInq 
      Caption         =   "Do Inquiry (6)"
      Default         =   -1  'True
      Height          =   390
      Left            =   683
      TabIndex        =   0
      Top             =   675
      Width           =   3315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' A demonstration on how to use the
' Flamed CDB gate to send and recieve
' own packets.

Private cManager        As New FL_Manager
Private cExec           As New FL_ExecCDB

Private Type t_InqDat
    PDT               As Byte     ' drive type
    PDQ               As Byte     ' removable drive
    VER               As Byte     ' MMC Version (zero for ATAPI)
    RDF               As Byte     ' interface depending field
    DLEN              As Byte     ' additional len
    rsv1(1)           As Byte     ' reserved
    Feat              As Byte     ' ?
    VID(7)            As Byte     ' vendor
    PID(15)           As Byte     ' Product
    PVER(3)           As Byte     ' revision (= Firmware Version)
    FWVER(20)         As Byte     ' ?
End Type

Private strDrvID      As String

Private Sub cboDrv_Change()

    strDrvID = vbNullString
    lstRes.Clear

    If cManager.IsCDVDDrive(cboDrv.Drive) Then
        strDrvID = cManager.DrvChr2DrvID(cboDrv.Drive)
    End If

End Sub

Private Sub cmdInq_Click()

    Dim cmd(5)  As Byte      ' 6-byte CDB
    Dim udtInq  As t_InqDat  ' inquiry data structure

    If strDrvID = vbNullString Then
        MsgBox "Please select a CD/DVD-ROM device", vbExclamation
        Exit Sub
    End If

    cmd(0) = &H12               ' inquiry opcode
    cmd(4) = Len(udtInq) - 1    ' allocation length

    If Not cExec.ExecCMD(strDrvID, cmd, 6, DIR_IN, VarPtr(udtInq), Len(udtInq) - 1) Then
        MsgBox "Inquiry failed." & vbCrLf & _
               "Sense key: " & cExec.LastSK & vbCrLf & _
               "Add. Sense Code: " & cExec.LastASC & vbCrLf & _
               "Add. Sense Code Qualifier: " & cExec.LastASCQ, vbExclamation
    Else
        lstRes.AddItem "Vendor: " & StrConv(udtInq.VID, vbUnicode)
        lstRes.AddItem "Product: " & StrConv(udtInq.PID, vbUnicode)
        lstRes.AddItem "Revision: " & StrConv(udtInq.PVER, vbUnicode)
    End If

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
