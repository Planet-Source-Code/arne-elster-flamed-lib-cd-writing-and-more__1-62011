VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FL_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************
'*** Flamed Info (About) Class
'****************************************

Private Const strAuthor As String = "[rm_code]"
Private Const intYear   As Integer = "2005"

Public Property Get VersionMajor() As Integer
    VersionMajor = App.Major
End Property

Public Property Get VersionMinor() As Integer
    VersionMinor = App.Minor
End Property

Public Property Get Author() As String
    Author = strAuthor
End Property

Public Property Get Year() As Integer
    Year = intYear
End Property
