VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "DVD info"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
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
   ScaleHeight     =   139
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   319
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox lstNfo 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1425
      Left            =   172
      TabIndex        =   1
      Top             =   525
      Width           =   4440
   End
   Begin VB.DriveListBox cboDrv 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   172
      TabIndex        =   0
      Top             =   150
      Width           =   4440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cManager        As New FL_Manager
Private cDVDInfo        As New FL_DVDInfo

Private strDrvID        As String

Private Sub cboDrv_Change()

    strDrvID = vbNullString
    lstNfo.Clear

    If cManager.IsCDVDDrive(cboDrv.Drive) Then
        strDrvID = cManager.DrvChr2DrvID(cboDrv.Drive)
        ShowInfo
    End If

End Sub

Sub ShowInfo()

    Dim intLayers   As Integer
    Dim strBuf      As String

    If Not cDVDInfo.GetInfo(strDrvID, 1) Then
        MsgBox "Could not get info for layer 1.", vbExclamation
        Exit Sub
    End If

    If Not cDVDInfo.GetInfo(strDrvID, 1) Then
        MsgBox "Could not get info for layer 1.", vbExclamation
        Exit Sub
    End If

    With cDVDInfo

        lstNfo.AddItem "Layers: " & intLayers

        Select Case .BookType
            Case FL_DVD_BOOKTYPES.DVD_ROM: strBuf = "DVD-ROM"
            Case FL_DVD_BOOKTYPES.DVD_RAM: strBuf = "DVD-RAM"
            Case FL_DVD_BOOKTYPES.DVD_R: strBuf = "DVD-R"
            Case FL_DVD_BOOKTYPES.DVD_RW: strBuf = "DVD-RW"
            Case FL_DVD_BOOKTYPES.DVD_PLUS_R: strBuf = "DVD+R"
            Case FL_DVD_BOOKTYPES.DVD_PLUS_RW: strBuf = "DVD+RW"
        End Select
        lstNfo.AddItem "Booktype: " & strBuf
        lstNfo.AddItem "Part version: " & .PartVersion

        Select Case .DiskSize
            Case DVD_120mm: strBuf = "120 mm"
            Case DVD_80mm: strBuf = "80 mm"
        End Select
        lstNfo.AddItem "Disk size: " & strBuf

        If CBool(.LayerType And DVD_DATA_EMBOSSED) Then
            strBuf = "data embossed "
        End If
        If CBool(.LayerType And DVD_DATA_RECORDED) Then
            strBuf = strBuf & "recordable "
        End If
        If CBool(.LayerType And DVD_DATA_REWRITABLE) Then
            strBuf = strBuf & "rewritable"
        End If
        lstNfo.AddItem "Layertype: " & strBuf

        Select Case .LinearDensity
            Case [0.267 um/bit]: strBuf = "0.267 um/bit"
            Case [0.280 to 0.291 um/bit]: strBuf = "0.280 to 0.291 um/bit"
            Case [0.293 um/bit]: strBuf = "0.293 um/bit"
            Case [0.353 um/bit]: strBuf = "0.353 um/bit"
            Case [0.409 to 0.435 um/bit]: strBuf = "0.409 to 0.435 um/bit"
        End Select
        lstNfo.AddItem "Linear density: " & strBuf

        Select Case .TrackDensity
            Case [0.615 um/track]: strBuf = "0.615 um/track"
            Case [0.74 um/track]: strBuf = "0.74 um/track"
            Case [0.80 um/track]: strBuf = "0.80 um/track"
        End Select
        lstNfo.AddItem "Track density: " & strBuf

        Select Case .MaximumRate
            Case [10.08 Mbps]: strBuf = "10.08 Mbps"
            Case [5.04 Mbps]: strBuf = "5.04 Mbps"
            Case [2.52 Mbps]: strBuf = "2.52 Mbps"
        End Select
        lstNfo.AddItem "Maximum Rate: " & strBuf

        Select Case .TrackPath
            Case DVD_PARALLEL_TRACK_PATH: strBuf = "parallel track path"
            Case DVD_OPPOSITE_TRACK_PATH: strBuf = "opposite track path"
        End Select
        lstNfo.AddItem "Track path: " & strBuf

        lstNfo.AddItem "Physical start sector data area: " & .PhysicalStartSectorDataArea
        lstNfo.AddItem "Physical end sector data area: " & .PhysicalEndSectorDataArea
        lstNfo.AddItem "Physical end sector layer 0: " & .PhysicalEndSectorLayer0

    End With

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
