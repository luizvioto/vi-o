VERSION 5.00
Begin VB.Form Form25 
   Appearance      =   0  'Flat
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Form25"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form25"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   1200
      MouseIcon       =   "Form25.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Próxima página"
         Height          =   315
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   8040
         Width           =   2175
      End
      Begin VB.Image Image2 
         Height          =   7395
         Left            =   240
         MouseIcon       =   "Form25.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Form25.frx":0614
         Top             =   360
         Width           =   8835
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   135
      Left            =   1200
      MouseIcon       =   "Form25.frx":3B00
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sair"
         Height          =   315
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   8040
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Página anterior"
         Height          =   315
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   8040
         Width           =   2175
      End
      Begin VB.Image Image3 
         Height          =   7590
         Left            =   240
         MouseIcon       =   "Form25.frx":3E0A
         Picture         =   "Form25.frx":4114
         Top             =   360
         Width           =   8835
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Aircraft Fucking Optmization  - TAFO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   555
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   8910
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      FillColor       =   &H0000FFFF&
      Height          =   5175
      Left            =   2160
      Top             =   1680
      Width           =   7935
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   11280
      Picture         =   "Form25.frx":756E
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MDO   PROJETO CONCEITUAL DE AERONAVES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1560
      TabIndex        =   5
      Top             =   7440
      Width           =   9075
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   8775
      Left            =   120
      Top             =   120
      Width           =   11775
   End
   Begin VB.Image Image1 
      Height          =   5175
      Left            =   2160
      Picture         =   "Form25.frx":79B0
      Top             =   1680
      Width           =   7920
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame2.Visible = True
Frame1.Visible = False
End Sub

Private Sub Command2_Click()
Frame1.Visible = True
Frame2.Visible = False
End Sub

Private Sub Command3_Click()
Frame1.Visible = False
Frame2.Visible = False

End Sub
Private Sub Image1_Click()
Form1.Visible = True
Form1.fileinicio1.Visible = True
Form25.Visible = False
Form1.Refresh
End Sub

Private Sub Label1_Click()
Form1.Visible = True
Form1.fileinicio1.Visible = True
Form25.Visible = False
Form1.Refresh
End Sub


