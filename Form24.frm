VERSION 5.00
Begin VB.Form Form24 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Posição da asa"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form24.frx":0000
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox helpasa 
      BackColor       =   &H00E0E0E0&
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Asa s/ suportes = cantilever - asa c/ suportes = semi cantilever"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4080
         Width           =   4695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asa Alta "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "A asa é fixada sobre a fuselagem do avião"
         Height          =   495
         Left            =   2760
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asa Média"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   4
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "A asa é fixada no centro da fuselagem do avião"
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Asa Baixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2760
         TabIndex        =   2
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "A asa é fixada sob a fuselagem do avião"
         Height          =   435
         Left            =   2760
         TabIndex        =   1
         Top             =   3240
         Width           =   1545
      End
      Begin VB.Image Image4 
         Height          =   1185
         Left            =   120
         MouseIcon       =   "Form24.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Form24.frx":0614
         Top             =   1440
         Width           =   2355
      End
      Begin VB.Image Image3 
         Height          =   1185
         Left            =   120
         MouseIcon       =   "Form24.frx":97FE
         MousePointer    =   99  'Custom
         Picture         =   "Form24.frx":9B08
         Top             =   2760
         Width           =   2355
      End
      Begin VB.Image Image6 
         Height          =   1185
         Left            =   120
         MouseIcon       =   "Form24.frx":12CF2
         MousePointer    =   99  'Custom
         Picture         =   "Form24.frx":12FFC
         Top             =   120
         Width           =   2355
      End
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image3_Click()
Form1.posibaixa.Value = True
Form24.Visible = False
End Sub

Private Sub Image4_Click()
Form1.posimedia.Value = True
Form24.Visible = False
End Sub

Private Sub Image6_Click()
Form1.posialta.Value = True
Form24.Visible = False
End Sub
