VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_CDBlanker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************
'*** Flamed CD-RW blanker
'***
'*** blanks CD-RWs either fastly or fully
'*******************************************

Public Enum FL_BlankMode
    BLANK_FULL                       ' full blank
    BLANK_QUICK                      ' quick blank
End Enum

Public Function BlankCDRW(DrvID As String, _
                mode As FL_BlankMode, _
                sync As Boolean) As Boolean

    BlankCDRW = CDRomEraseCDRW(DrvID, mode, sync)

    If BlankCDRW And Not sync Then
        Do While Not CDRomTestUnitReady(DrvID)
            Sleep 800
            DoEvents
        Loop
    End If

End Function
