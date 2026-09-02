VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tempo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aguarde..."
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   ClipControls    =   0   'False
   Icon            =   "Tempo.frx":0000
   LinkTopic       =   "Form32"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "Tempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
