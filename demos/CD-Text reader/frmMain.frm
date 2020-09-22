VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "CD-Text reader"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
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
   ScaleHeight     =   181
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read CD-Text"
      Default         =   -1  'True
      Height          =   315
      Left            =   2550
      TabIndex        =   7
      Top             =   2295
      Width           =   1515
   End
   Begin VB.ListBox lstTracks 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   900
      TabIndex        =   6
      Top             =   1350
      Width           =   3315
   End
   Begin VB.TextBox txtAlbum 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   975
      Width           =   3315
   End
   Begin VB.TextBox txtArtist 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   3315
   End
   Begin VB.DriveListBox cboDrv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tracks:"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   1350
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Album:"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   975
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Artist:"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   600
      Width           =   525
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cManager        As New FL_Manager
Private cDrvInfo        As New FL_DriveInfo
Private cCDText         As New FL_CDText

Private strDrvID        As String

Private Sub cboDrv_Change()

    strDrvID = vbNullString
    txtAlbum = vbNullString
    txtArtist = vbNullString
    lstTracks.Clear

    ' get new drive ID
    If cManager.IsCDVDDrive(cboDrv.Drive) Then
        strDrvID = cManager.DrvChr2DrvID(cboDrv.Drive)
    End If

End Sub

Private Sub cmdRead_Click()

    Dim i   As Integer

    If strDrvID = vbNullString Then
        MsgBox "No CD/DVD-ROM drive selected.", vbExclamation
        Exit Sub
    End If

    cDrvInfo.GetInfo strDrvID

    ' check if drive supports CD-Text
    If Not CBool(cDrvInfo.ReadCapabilities And RC_CDTEXT) Then
        If MsgBox("Drive reports it can't read CD-Text." & vbCrLf & "Try either?", vbQuestion Or vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    ' read CD-Text
    If Not cCDText.ReadCDText(strDrvID) Then
        MsgBox "Reading failed.", vbExclamation, "Error"
        Exit Sub
    End If

    ' show em
    txtAlbum.Text = cCDText.Album
    txtArtist.Text = cCDText.Artist

    For i = 0 To cCDText.TrackCount - 1
        lstTracks.AddItem "Track " & (i + 1) & ": " & cCDText.Track(i)
    Next

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
