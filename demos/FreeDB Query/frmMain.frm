VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "FreeDB Query"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
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
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame frmQuery 
      Caption         =   "Query"
      Height          =   1365
      Left            =   2400
      TabIndex        =   10
      Top             =   600
      Width           =   1815
      Begin VB.PictureBox picQuery 
         BorderStyle     =   0  'Kein
         Height          =   1065
         Left            =   75
         ScaleHeight     =   1065
         ScaleWidth      =   1665
         TabIndex        =   11
         Top             =   225
         Width           =   1665
         Begin VB.ListBox lstCDDB 
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   600
            IntegralHeight  =   0   'False
            Left            =   75
            TabIndex        =   13
            Top             =   375
            Width           =   1515
         End
         Begin VB.CommandButton cmdQuery 
            Caption         =   "query"
            Default         =   -1  'True
            Height          =   315
            Left            =   75
            TabIndex        =   12
            Top             =   0
            Width           =   1515
         End
      End
   End
   Begin VB.Frame frmListing 
      Caption         =   "Track listing"
      Height          =   1365
      Left            =   150
      TabIndex        =   8
      Top             =   2025
      Width           =   4065
      Begin VB.ListBox lstListing 
         BackColor       =   &H00000040&
         ForeColor       =   &H00FFFFFF&
         Height          =   1035
         Left            =   100
         TabIndex        =   9
         Top             =   225
         Width           =   3840
      End
   End
   Begin VB.Frame frmCDDB 
      Caption         =   "CDDB"
      Height          =   1365
      Left            =   150
      TabIndex        =   1
      Top             =   600
      Width           =   2115
      Begin VB.PictureBox picCDDB 
         BorderStyle     =   0  'Kein
         Height          =   765
         Left            =   75
         ScaleHeight     =   765
         ScaleWidth      =   1965
         TabIndex        =   2
         Top             =   225
         Width           =   1965
         Begin VB.TextBox txtTimeout 
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   4
            Text            =   "8"
            Top             =   75
            Width           =   540
         End
         Begin VB.Label lblID 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   1025
            TabIndex        =   7
            Top             =   480
            Width           =   60
         End
         Begin VB.Label lblCDDBID 
            AutoSize        =   -1  'True
            Caption         =   "CDDB ID:"
            Height          =   195
            Left            =   75
            TabIndex        =   6
            Top             =   480
            Width           =   870
         End
         Begin VB.Label lblSeconds 
            AutoSize        =   -1  'True
            Caption         =   "s"
            Height          =   195
            Left            =   1680
            TabIndex        =   5
            Top             =   105
            Width           =   90
         End
         Begin VB.Label lblTimeout 
            Caption         =   "Timeout:"
            Height          =   240
            Left            =   180
            TabIndex        =   3
            Top             =   105
            Width           =   765
         End
      End
   End
   Begin VB.DriveListBox cboDrv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4140
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cCDDB    As FL_FreeDB
Attribute cCDDB.VB_VarHelpID = -1
Private cManager            As New FL_Manager

Private strDrvID            As String
Private blnCancel           As Boolean

Private Sub cboDrv_Change()
    strDrvID = vbNullString
    If cManager.IsCDVDDrive(cboDrv.Drive) Then
        strDrvID = cManager.DrvChr2DrvID(cboDrv.Drive)
        cCDDB.DriveID = strDrvID
        lblID = cCDDB.CDDBID
    End If
End Sub

Private Sub cCDDB_Status(Status As FlamedLib.FL_CDDBState)
    Select Case Status
        Case CDDB_CLOSE: lstCDDB.AddItem "Closing..."
        Case CDDB_DATA: lstCDDB.AddItem "Recieving..."
        Case CDDB_HELLO: lstCDDB.AddItem "Handshake..."
        Case CDDB_QUERY: lstCDDB.AddItem "Querying..."
        Case CDDB_RESULT: lstCDDB.AddItem "parsing..."
    End Select
End Sub

Private Sub cmdQuery_Click()
    ' caption of button is "cancel"?
    If cmdQuery.Caption = "cancel" Then
        ' yes, set cancel bool
        blnCancel = True
    ' no, we have to start the query
    Else
        ' clear protocol
        lstCDDB.Clear
        ' set button caption to "cancel"
        cmdQuery.Caption = "cancel"
        ' query FreeDB
        If cCDDB.Query(blnCancel) Then
            MsgBox "Finished", vbInformation, "Ok"
            ' show recieved data
            ShowData
        Else
            MsgBox "Failed", vbExclamation, "Bah"
        End If
        ' set button caption to "query" again
        cmdQuery.Caption = "query"
    End If
End Sub

Private Sub ShowData()
    Dim i   As Integer

    With lstListing

        .Clear

        ' show album/artist
        .AddItem "Artist: " & cCDDB.Artist
        .AddItem "Album: " & cCDDB.Album
        .AddItem ""

        ' show tracks
        For i = 1 To cCDDB.Tracks
            .AddItem "Track " & Format(i, "00") & ": " & cCDDB.Track(i)
        Next

    End With
End Sub

Private Sub Form_Load()
    Set cCDDB = New FL_FreeDB

    If Not cManager.Init(False) Then
        MsgBox "No interfaces found.", vbExclamation, "Error"
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cManager.Goodbye
End Sub

Private Sub txtTimeout_KeyPress(KeyAscii As Integer)
    ' only allow numbers
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0
    cCDDB.TimeOut = CLng(txtTimeout)
End Sub
