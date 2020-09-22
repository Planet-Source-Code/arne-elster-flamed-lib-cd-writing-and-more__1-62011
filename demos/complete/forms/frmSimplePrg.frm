VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSimplePrg 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Progress"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
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
   ScaleHeight     =   84
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   270
      Left            =   1695
      TabIndex        =   2
      Top             =   900
      Width           =   1290
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Top             =   450
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2318
      TabIndex        =   0
      Top             =   150
      Width           =   45
   End
End
Attribute VB_Name = "frmSimplePrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

