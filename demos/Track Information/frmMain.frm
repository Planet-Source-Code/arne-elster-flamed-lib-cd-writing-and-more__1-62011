VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Track information"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
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
   ScaleHeight     =   2370
   ScaleWidth      =   6030
   StartUpPosition =   2  'Bildschirmmitte
   Begin ComctlLib.TreeView trks 
      Height          =   1665
      Left            =   105
      TabIndex        =   1
      Top             =   525
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   2937
      _Version        =   327682
      Indentation     =   294
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.DriveListBox cboDrv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   5790
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cManager        As New FL_Manager
Private cCInfo          As New FL_CDInfo
Private cTInfo          As New FL_TrackInfo
Private cSInfo          As New FL_SessionInfo

Private strDrvID        As String

Private Sub cboDrv_Change()
    trks.Nodes.Clear
    If cManager.IsCDVDDrive(cboDrv.Drive) Then
        strDrvID = cManager.DrvChr2DrvID(cboDrv.Drive)
        Info
    End If
End Sub

Private Sub Info()

    Dim i   As Integer
    Dim j   As Integer

    If Not cCInfo.GetInfo(strDrvID) Then
        MsgBox "Couldn't read CD information", vbExclamation, "Error"
        Exit Sub
    End If

    ' show sessions
    For i = 1 To cCInfo.Sessions

        If Not cSInfo.GetInfo(strDrvID, i) Then
            MsgBox "Couldn't read information about session " & i, vbExclamation, "Error"
            ' don't want to have complex structures here
            ' keep it as simple as possible
            GoTo SkipSession
        End If

        ' add session node
        trks.Nodes.Add(, , "s" & Format(i, "00"), "Session " & Format(i, "00")).Expanded = True

        ' show tracks for the current session
        For j = cSInfo.FirstTrack To cSInfo.LastTrack

            ' add track node
            trks.Nodes.Add("s" & Format(i, "00"), tvwChild, "t" & Format(j, "00"), "Track " & Format(j, "00")).Expanded = True

            If Not cTInfo.GetInfo(strDrvID, j) Then
                MsgBox "Couldn't read information about track " & j, vbExclamation, "Error"
                ' don't want to have complex structures here
                ' keep it as simple as possible
                GoTo SkipTrack
            End If

            ' track information
            trks.Nodes.Add "t" & Format(j, "00"), tvwChild, , "Track start: " & cTInfo.TrackStart.MSF & " MSF"
            trks.Nodes.Add "t" & Format(j, "00"), tvwChild, , "Track length: " & cTInfo.TrackLength.MSF & " MSF"
            trks.Nodes.Add "t" & Format(j, "00"), tvwChild, , "Track end: " & cTInfo.TrackEnd.MSF & " MSF"
            trks.Nodes.Add "t" & Format(j, "00"), tvwChild, , "Mode: " & TrackMode2Str(cTInfo.mode)

SkipTrack:
        Next

SkipSession:
    Next

End Sub

Private Function TrackMode2Str(m As FL_TrackModes) As String
    Select Case m
        Case MODE_AUDIO: TrackMode2Str = "audio"
        Case MODE_MODE1: TrackMode2Str = "mode 1"
        Case MODE_MODE2: TrackMode2Str = "mode 2"
        Case MODE_MODE2_FORM1: TrackMode2Str = "mode 2 form 1"
        Case MODE_MODE2_FORM2: TrackMode2Str = "mode 2 form 2"
    End Select
End Function

Private Sub Form_Load()
    If Not cManager.Init(False) Then
        MsgBox "No interfaces found.", vbExclamation, "Error"
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cManager.Goodbye
End Sub
