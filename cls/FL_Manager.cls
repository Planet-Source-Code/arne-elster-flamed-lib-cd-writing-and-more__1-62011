VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ToDo: functions to set drive timeouts

'****************************************
'*** Flamed Manager
'***
'*** for Drive (ID) management and
'*** SCSI interfaces information
'****************************************

Private cDrvInfo    As New FL_DriveInfo

Public Function Init(Optional blnForceASPI As Boolean) As Boolean
    ' initialize the library's heart
    Init = cd.ChooseInterface(blnForceASPI)
End Function

Public Sub Goodbye()
    ' close all handles and unload ASPI lib
    cd.Goodbye
End Sub

Public Function GetCDVDROMs() As String()

    Dim i           As Integer
    Dim drive       As String
    Dim drives()    As String
    ReDim drives(0) As String

    For i = 1 To 26
        drive = Chr$(i + 64)
        If IsCDVDDrive(drive) Then
            drives(UBound(drives)) = drive
            ReDim Preserve drives(UBound(drives) + 1) As String
        End If
    Next

    GetCDVDROMs = drives

End Function

' &HFFFF& = Max
Public Function SetCDRomSpeed(DriveID As String, _
                              ReadSpeed As Long, _
                              WriteSpeed As Long) As Boolean

    Dim speeds()    As Integer
    Dim i           As Integer
    Dim CAV         As Boolean

    CDRomTestUnitReady DriveID

    CDRomGetWriteSpeeds DriveID, speeds

    For i = LBound(speeds) To UBound(speeds)
        If (speeds(i) And &H7FFF) = WriteSpeed Then
            CAV = CBool(speeds(i) And &H8000)
            Exit For
        End If
    Next

    SetCDRomSpeed = CDRomSetCDSpeed(DriveID, ReadSpeed, WriteSpeed, CAV)

End Function

Public Property Get ASPISupport() As Boolean
    ASPISupport = cd.InterfaceAvailable(ASPI)
End Property

Public Property Get SPTISupport() As Boolean
    SPTISupport = cd.InterfaceAvailable(SPTI)
End Property

Public Property Get CurrentInterface() As Integer
    CurrentInterface = cd.CurrentInterface
End Property

Public Function IsCDVDDrive(ByVal char As String) As Boolean
    IsCDVDDrive = cd.IsCDROM(Left$(char, 1) & ":")
End Function

Public Function DrvChr2DrvID(ByVal char As String) As String
    DrvChr2DrvID = cd.DrvIDFromDrvChr(Left$(char, 1) & ":")
End Function

Public Function DrvID2DrvChr(ByVal DrvID As String) As String
    DrvID2DrvChr = cd.DrvChrFromDrvID(DrvID)
End Function

Public Function TestUnitReady(ByVal DrvID As String) As Boolean
    TestUnitReady = CDRomTestUnitReady(DrvID)
End Function

Public Function LoadDrive(ByVal DrvID As String) As Boolean
    LoadDrive = CDRomLoadTray(DrvID)
End Function

Public Function UnLoadDrive(ByVal DrvID As String) As Boolean
    UnLoadDrive = CDRomUnloadTray(DrvID)
End Function

Public Function LockDrive(ByVal DrvID As String) As Boolean
    LockDrive = CDRomLockMedia(DrvID)
End Function

Public Function UnLockDrive(ByVal DrvID As String) As Boolean
    UnLockDrive = CDRomUnlockMedia(DrvID)
End Function
