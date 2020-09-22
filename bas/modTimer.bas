Attribute VB_Name = "modTimer"
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

'Version 1.00, 15.01.03
'(c) by Goetz Reinecke 01/2003
'    reinecke@activevb.de

Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As _
        Long, ByVal nIDEvent As Long, ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As _
        Long, ByVal nIDEvent As Long) As Long

Private colTimer As New Collection

Private Function Init(Interval As Long) As Long
    Init = SetTimer(0&, 0&, Interval, AddressOf TimerProc)
End Function

Private Sub Terminate(ByRef hTimer As Long)
    Call KillTimer(0, hTimer)
End Sub

Private Sub TimerProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal hTimer As Long, ByVal dwTime As Long)
    On Error Resume Next
    
    Static Flag As Boolean
    Dim cTimer As clsTimer
    
        If Not Flag Then
            Flag = True
            Set cTimer = colTimer(CStr(hTimer))
            cTimer.TimerEvent
            Set cTimer = Nothing
            Flag = False
        End If
End Sub

Public Function AddObject(ByRef cTimer As clsTimer, ByRef Interval As Long) As Long
    On Error Resume Next
    
    Dim hTimer As Long
    
        hTimer = Init(Interval)
        colTimer.Add cTimer, CStr(hTimer)
        AddObject = hTimer
End Function

Public Function RemoveObject(ByRef hTimer As Long) As Boolean
    On Error Resume Next
    
    Call Terminate(hTimer)
    colTimer.Remove CStr(hTimer)
End Function
