Attribute VB_Name = "modMP3Info"
'Dieser Source stammt von http://www.vb-fun.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Option Explicit

Public Type MP3Info
  Bitrate       As Integer
  Frequency     As Long
  Channels      As Integer
  MpegVersion   As Integer
  MpegLayer     As Integer
  Duration      As Long
  VBR           As Boolean
  Frames        As Integer
End Type

Private GetMP3Info  As MP3Info

Private Function BinToDec(BinValue As String) As Long

    Dim i As Integer
    BinToDec = 0

    For i = 1 To Len(BinValue)
        If Mid(BinValue, i, 1) = 1 Then
            BinToDec = BinToDec + 2 ^ (Len(BinValue) - i)
        End If
    Next i

End Function

Private Function ByteToBit(ByteArray) As String

    Dim i   As Integer, z     As Integer

    ByteToBit = ""

    For z = 1 To 4
        For i = 7 To 0 Step -1
            If Int(ByteArray(z) / (2 ^ i)) = 1 Then
                ByteToBit = ByteToBit & "1"
                ByteArray(z) = ByteArray(z) - (2 ^ i)
            Else
                If ByteToBit <> "" Then
                    ByteToBit = ByteToBit & "0"
                End If
            End If
        Next i
    Next z

End Function

Private Function BinaryHeader(FileName As String) As String

    Dim ByteArray(4)    As Byte
    Dim XingH           As String * 4
    Dim FIO             As Integer
    Dim n               As Long
    Dim i               As Integer
    Dim X               As Byte
    Dim z               As Integer
    Dim headstart       As Long

    If FileName = vbNullString Then Exit Function

    FIO = FreeFile

    Open FileName For Binary Access Read As #FIO
    n = LOF(FIO): If n < 256 Then Close #FIO: Return

    For i = 1 To 5000

        Get #FIO, i, X

        If X = 255 Then

            Get #FIO, i + 1, X

            If X > 249 And X < 252 Then
                headstart = i
                Exit For
            End If

        End If

    Next i

    Get #1, headstart + 36, XingH

    If XingH = "Xing" Then

        GetMP3Info.VBR = True
        For z = 1 To 4 '
            Get #1, headstart + 43 + z, ByteArray(z)
        Next z

        GetMP3Info.Frames = BinToDec(ByteToBit(ByteArray))

    Else

        GetMP3Info.VBR = False

    End If
  
    For z = 1 To 4
        Get #1, headstart + z - 1, ByteArray(z)
    Next z

    Close #FIO

    BinaryHeader = ByteToBit(ByteArray)

End Function

Public Function ReadMP3(FileName As String) As MP3Info

    Dim LayerVersion As String
    Dim Version()    As Variant
    Dim layer()      As Variant
    Dim sMode()      As Variant
    Dim MpegVersion  As Integer
    Dim MpegLayer    As Integer
    Dim Freq()       As Variant
    Dim Frequency    As Long
    Dim Temp()       As Variant
    Dim Bitrate      As Long
    Dim BRate()      As Variant
    Dim bin          As String

    If FileName = "" Then Exit Function

    bin = BinaryHeader(FileName)

    Version = Array(25, 0, 2, 1)
    layer = Array(0, 3, 2, 1)
    sMode = Array(2, 2, 2, 1)

    MpegVersion = Version(BinToDec(Mid(bin, 12, 2)))
    MpegLayer = layer(BinToDec(Mid(bin, 14, 2)))

    GetMP3Info.Channels = sMode(BinToDec(Mid(bin, 25, 2)))

    Select Case MpegVersion

        Case 1
            Freq = Array(44100, 48000, 32000)

        Case 2 Or 25
            Freq = Array(22050, 24000, 16000)

        Case Else
            Frequency = 0
            Exit Function

    End Select

    Frequency = Freq(BinToDec(Mid(bin, 21, 2)))

    If GetMP3Info.VBR = True Then

        Temp = Array(, 12, 144, 144)
        Bitrate = (CDbl(FileLen(FileName)) * CDbl(Frequency)) / (Int(GetMP3Info.Frames)) / 1000& / Temp(MpegLayer)

    Else

        LayerVersion = MpegVersion & MpegLayer
  
        Select Case Val(LayerVersion)
            Case 11
                BRate = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, _
                              288, 320, 352, 384, 416, 448)

            Case 12
                BRate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, _
                              192, 224, 256, 320, 384)

            Case 13
                BRate = Array(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, _
                              160, 192, 224, 256, 320)

            Case 21 Or 251
                BRate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 144, _
                              160, 176, 192, 224, 256)

            Case 22 Or 252 Or 23 Or 253
                BRate = Array(0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, _
                              112, 128, 144, 160)

            Case Else
                Bitrate = 1
                Exit Function

        End Select
  
        Bitrate = BRate(BinToDec(Mid(bin, 17, 4)))

    End If

    With GetMP3Info
        .Bitrate = Bitrate
        .Frequency = Frequency
        .MpegLayer = MpegLayer
        .MpegVersion = MpegVersion
        .Duration = ((FileLen(FileName) * 8&) / Bitrate) / 1000&
    End With

    ReadMP3 = GetMP3Info

End Function
