Attribute VB_Name = "modFncs"
Option Explicit

Public Enum e_DriveListOpts
    OPT_CDWRITERS
    OPT_DVD
    OPT_ALL
End Enum

Public Function PathFromPathFile(ByVal strText As String) As String
    On Error Resume Next
    PathFromPathFile = Left$(strText, InStrRev(strText, "\"))
End Function

Public Function FileFromPath(ByVal strText As String) As String
    On Error Resume Next
    FileFromPath = Mid$(strText, InStrRev(strText, "\") + 1)
End Function

Public Function AddSlash(ByVal strText As String) As String

    If InStr(strText, "/") > 0 Then
        If Not Right$(strText, 1) = "/" Then strText = strText & "/"
    Else
        If Not Right$(strText, 1) = "\" Then strText = strText & "\"
    End If

    AddSlash = strText

End Function

Public Function GetDriveList(options As e_DriveListOpts) As String()

    Dim cDrvNfo     As New FL_DriveInfo

    Dim strDrvs()   As String
    Dim strRet()    As String
    ReDim strRet(0) As String

    Dim i           As Integer

    strDrvs = cManager.GetCDVDROMs

    For i = LBound(strDrvs) To UBound(strDrvs) - 1

        cDrvNfo.GetInfo cManager.DrvChr2DrvID(strDrvs(i))

        Select Case options

            Case OPT_ALL:
                strRet(UBound(strRet)) = strDrvs(i)
                ReDim Preserve strRet(UBound(strRet) + 1) As String

            Case OPT_DVD:
                If (cDrvNfo.ReadCapabilities And RC_DVDROM) Then
                    strRet(UBound(strRet)) = strDrvs(i)
                    ReDim Preserve strRet(UBound(strRet) + 1) As String
                End If

            Case OPT_CDWRITERS:
                If (cDrvNfo.WriteCapabilities And WC_CDR) Then
                    strRet(UBound(strRet)) = strDrvs(i)
                    ReDim Preserve strRet(UBound(strRet) + 1) As String
                End If

        End Select

    Next

    GetDriveList = strRet

End Function
