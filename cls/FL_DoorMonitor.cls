VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_DoorMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************
'*** Flamed drive monitor
'***
'*** detects media arrival/removal
'****************************************

Private WithEvents subcls               As lvSubclasser
Attribute subcls.VB_VarHelpID = -1

Private Const WM_DEVICECHANGE           As Long = &H219&
Private Const DBT_DEVICEARRIVAL         As Long = &H8000&
Private Const DBT_DEVICEREMOVECOMPLETE  As Long = &H8004&
Private Const DBTF_MEDIA                As Long = &H1&
Private Const DBT_DEVTYP_VOLUME         As Long = &H2&

Private Type DEV_BROADCAST_HDR
    dbch_size           As Long
    dbch_devicetype     As Long
    dbch_reserved       As Long
End Type

Private Type DEV_BROADCAST_VOLUME
    dbch_size           As Long
    dbch_devicetype     As Long
    dbch_reserved       As Long
    dbcv_unitmask       As Long
    dbcv_flags          As Integer
End Type

Private lngHWND         As Long

Public Event arrival(ByVal drive As String)
Public Event removal(ByVal drive As String)

Public Sub DeInitDoorMonitor()
    subcls.UnSubclassMe lngHWND
    DestroyWindow lngHWND
End Sub

Public Function InitDoorMonitor() As Boolean

    lngHWND = CreateWindowEx( _
                0&, "STATIC", _
                "fl_doormonitor", _
                0&, 0&, 0&, _
                0&, 0&, 0&, _
                0&, App.hInstance, _
                0&)

    If lngHWND <> 0 Then
        subcls.SubclassMe lngHWND
        InitDoorMonitor = True
    End If

End Function

Private Sub ProcessDeviceChange(wParam As Long, lParam As Long)

    Dim DBHdr   As DEV_BROADCAST_HDR
    Dim DBVol   As DEV_BROADCAST_VOLUME

    ' check for a valid pointer to read from
    If IsBadReadPtr(lParam, 4) Then Exit Sub
    CopyMemory DBHdr, ByVal lParam, LenB(DBHdr)

    Select Case wParam

        Case DBT_DEVICEARRIVAL

            If DBHdr.dbch_devicetype = DBT_DEVTYP_VOLUME Then

                CopyMemory DBVol, ByVal lParam, LenB(DBVol)
                If (DBVol.dbcv_flags And DBTF_MEDIA) = DBTF_MEDIA Then
                    RaiseEvent arrival(Chr$(FirstDriveFromMask(DBVol.dbcv_unitmask)))
                End If

            End If

        Case DBT_DEVICEREMOVECOMPLETE

            If DBHdr.dbch_devicetype = DBT_DEVTYP_VOLUME Then

                CopyMemory DBVol, ByVal lParam, LenB(DBVol)
                If (DBVol.dbcv_flags And DBTF_MEDIA) = DBTF_MEDIA Then
                    RaiseEvent removal(Chr$(FirstDriveFromMask(DBVol.dbcv_unitmask)))
                End If

            End If

    End Select

End Sub

Private Function FirstDriveFromMask(ByVal unitmask As Long) As Byte

    Dim b As Integer

    For b = 0 To 25
        If (unitmask And 2 ^ b) Then
            FirstDriveFromMask = Asc("A") + b
            Exit For
        End If
    Next b

End Function

Private Sub Class_Initialize()
    Set subcls = New lvSubclasser
End Sub

Private Sub Class_Terminate()
    subcls.UnSubclassMe lngHWND, True
End Sub

Private Sub subcls_ProcessMessage(ByVal hwnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Long, bOverriden As Boolean, lOverrideReturn As Long)
    If wMsg = WM_DEVICECHANGE Then
        ProcessDeviceChange wParam, lParam
    End If
End Sub
