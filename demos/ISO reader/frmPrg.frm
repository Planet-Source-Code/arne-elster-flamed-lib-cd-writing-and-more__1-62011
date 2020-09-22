VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrg 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Progress"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin MSComctlLib.ProgressBar prg 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmPrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

