VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_ExecCDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************
'*** command brigde
'***
'*** for sending/recieving own packets
'*******************************************

Public Enum FL_SRB_DIRECTION
    DIR_IN = &H8              ' retrieve data from the drive
    DIR_OUT = &H10            ' send data to the drive
End Enum

Public Property Get LastASCQ() As Integer
    LastASCQ = cd.LastASCQ
End Property

Public Property Get LastASC() As Integer
    LastASC = cd.LastASC
End Property

Public Property Get LastSK() As Integer
    LastSK = cd.LastSK
End Property

Public Function ExecCMD(ByVal DrvID As String, _
            ByRef cmd() As Byte, _
            ByVal CDBLen As Byte, _
            ByVal direction As FL_SRB_DIRECTION, _
            ByVal DataPointer As Long, _
            ByVal DataLen As Long, _
            Optional ByVal WaitForEver As Boolean = False, _
            Optional ByVal Timeout As Integer = 5) As Boolean

    ExecCMD = cd.ExecCMD(DrvID, cmd, CDBLen, _
                         WaitForEver, direction, _
                         DataPointer, DataLen, Timeout)

End Function

