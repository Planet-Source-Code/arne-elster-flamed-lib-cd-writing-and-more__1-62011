VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "ISO9660 Image Reader"
   ClientHeight    =   4185
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   474
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox lstInfo 
      Height          =   645
      Left            =   75
      TabIndex        =   0
      Top             =   3450
      Width           =   6930
   End
   Begin MSComctlLib.ImageList img 
      Left            =   5175
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgISO 
      Left            =   5805
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "ISO9660 images (*.iso)|*.iso"
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   6345
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   2
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   3375
      Left            =   2850
      TabIndex        =   1
      Top             =   0
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "filename"
         Object.Width           =   5116
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "size"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDirs 
      Height          =   3375
      Left            =   75
      TabIndex        =   2
      Top             =   0
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   5953
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   471
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenISO 
         Caption         =   "Open ISO"
      End
      Begin VB.Menu mnuExtractFile 
         Caption         =   "Extract selected file"
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents iso As FL_ISO9660Reader
Attribute iso.VB_VarHelpID = -1

Private Sub Form_Load()
    Set iso = New FL_ISO9660Reader
End Sub

Private Sub iso_ReadProgress(ByVal Percent As Integer)
    frmPrg.prg.Value = Percent
    DoEvents
End Sub

Private Sub tvwDirs_Click()

    Dim files() As String
    Dim i As Integer

    If iso.CurrentISO = vbNullString Then Exit Sub

    lvwFiles.ListItems.Clear

    ' are there any files in the dir?
    If iso.HasSubFiles(tvwDirs.SelectedItem.Key) Then

        ' get array with files for this dir
        files = iso.GetSubFiles(tvwDirs.SelectedItem.Key)

        ' add files and their sizes
        For i = 0 To UBound(files)
            With lvwFiles.ListItems.Add(, , files(i), , 3)
                .SubItems(1) = FormatFileSize(iso.GetFilesize(tvwDirs.SelectedItem.Key & files(i)))
            End With
        Next

    End If
End Sub

Private Sub Form_Resize()

    ' left/width
    tvwDirs.Width = Me.ScaleWidth * 3 / 8
    lvwFiles.Width = Me.ScaleWidth * 5 / 8 - 12 - tvwDirs.Left
    lvwFiles.Left = tvwDirs.Left + tvwDirs.Width + 6
    lstInfo.Width = Me.ScaleWidth - lstInfo.Left * 2

    ' top/height
    tvwDirs.Height = Me.ScaleHeight - lstInfo.Height - 12
    lvwFiles.Height = tvwDirs.Height
    lstInfo.Top = tvwDirs.Top + tvwDirs.Height + 6

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Function GetDirname(ByVal sDir As String) As String
    Dim parts() As String
    parts = Split(AddSlash(sDir), "\")
    GetDirname = parts(UBound(parts) - 1)
End Function

Private Function AddSlash(ByVal sVal As String) As String
    If Not Right$(sVal, 1) = "\" Then sVal = sVal & "\"
    AddSlash = sVal
End Function

Private Function GetUpperDir(ByVal sDir As String) As String
    Dim parts() As String, i As Integer

    sDir = AddSlash(sDir)
    parts = Split(sDir, "\")

    For i = 0 To UBound(parts) - 2
        GetUpperDir = GetUpperDir & parts(i) & "\"
    Next

    GetUpperDir = AddSlash(GetUpperDir)
End Function

Public Function FormatFileSize(ByVal dblFileSize As Double, Optional ByVal strFormatMask As String) As String
    Select Case dblFileSize

        Case 0 To 1023               ' Bytes
            FormatFileSize = Format(dblFileSize) & " bytes"

        Case 1024 To 1048575         ' KB
            If strFormatMask = Empty Then strFormatMask = "###0"
            FormatFileSize = Format(dblFileSize \ 1024#, strFormatMask) & " KB"

        Case 1024# ^ 2 To 1073741823 ' MB
            If strFormatMask = Empty Then strFormatMask = "####0.00"
            FormatFileSize = Format(dblFileSize \ (1024# * 1024#), strFormatMask) & " MB"

        Case Is > 1073741823#        ' GB
            If strFormatMask = Empty Then strFormatMask = "####0.00"
            FormatFileSize = Format(dblFileSize \ (1024# * 1024# * 1024#), strFormatMask) & " GB"

    End Select
End Function

Private Sub ListDir(ByVal sPath As String)
    Dim dirs() As String
    Dim i As Integer

    ' get all sub dirs of sPath
    If iso.HasSubDirs(sPath) Then

        ' get sub dirs for sPath
        dirs = iso.GetSubDirs(sPath)

        For i = 0 To UBound(dirs)

            ' add dirs...
            tvwDirs.Nodes.Add GetUpperDir(dirs(i)), tvwChild, dirs(i), GetDirname(dirs(i)), 1

            ' ... and sub dirs
            ListDir dirs(i)

        Next
    End If
End Sub

Private Sub mnuExtractFile_Click()
    On Error GoTo ErrH

    dlgSave.Filter = lvwFiles.SelectedItem.Text & "|" & lvwFiles.SelectedItem.Text
    dlgSave.FileName = lvwFiles.SelectedItem.Text
    dlgSave.ShowSave

    ' extract selected file
    frmPrg.Show , Me
    iso.ReadFileToFile tvwDirs.SelectedItem.Key & lvwFiles.SelectedItem.Text, dlgSave.FileName
    Unload frmPrg

ErrH:
End Sub

Private Sub mnuOpenISO_Click()

    dlgISO.FileName = vbNullString
    dlgISO.ShowOpen
    If dlgISO.FileName = vbNullString Then Exit Sub

    ' parse file system
    If Not iso.ReadISO(dlgISO.FileName) Then
        MsgBox "Konnte " & dlgISO.FileName & " nicht lesen.", vbExclamation, "Error"
        Exit Sub
    End If

    tvwDirs.Nodes.Clear
    lvwFiles.ListItems.Clear

    ' add root node
    tvwDirs.Nodes.Add , , "\", "\", 2

    ' show info about file system
    With lstInfo
        .Clear
        .AddItem "Volume ID: " & iso.VolumeID
        .AddItem "System ID: " & iso.SystemID
        .AddItem "Volume Size: " & FormatFileSize(iso.VolumeSize * 2048)
        .AddItem "Application: " & iso.ApplicationID
        .AddItem "Data Preparer: " & iso.DataPreparerID
        .AddItem "Publisher: " & iso.PublisherID
        .AddItem "Abstract file: " & iso.AbstractFile
        .AddItem "Bibliographic file: " & iso.BibliographicFile
        .AddItem "Copyright file: " & iso.CopyrightFile
        .AddItem "Creation Date: " & iso.VolumeCreationDate
        .AddItem "Effective Date: " & iso.VolumeEffectiveDate
        .AddItem "Expiration Date: " & iso.VolumeExpirationDate
        .AddItem "Modification Date: " & iso.VolumeModificationDate
    End With

    ' add dirs recursive
    ListDir "\"

End Sub
