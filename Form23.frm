VERSION 5.00
Begin VB.Form Form23 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trem de pouso"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form23.frx":0000
   LinkTopic       =   "Form23"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox helptp 
      BackColor       =   &H00E0E0E0&
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4515
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "H/D é a deslocamento da roda no sentido vertical em função do diametro da roda. Veja detalhe no manual."
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Entre com o valor de L para a largura ou espessura da roda."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   3360
         Width           =   4320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entre com o valor de D para o diametro da roda."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   3600
         Width           =   3405
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Bequilha é a roda que dirige o avião no solo, geralmente é a roda menor."
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   4335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Convencional"
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
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Triciclo"
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
         Left            =   2640
         TabIndex        =   3
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Trem de pouso principal frontal e bequilha traseira."
         Height          =   435
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Trem de pouso principal traseiro e bequilha frontal."
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   1185
         Left            =   120
         MouseIcon       =   "Form23.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Form23.frx":0614
         Top             =   240
         Width           =   2355
      End
      Begin VB.Image Image2 
         Height          =   1185
         Left            =   120
         MouseIcon       =   "Form23.frx":97FE
         MousePointer    =   99  'Custom
         Picture         =   "Form23.frx":9B08
         Top             =   1560
         Width           =   2355
      End
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Form1.tremconv.Value = True
Form23.Visible = False
End Sub

Private Sub Image2_Click()
Form1.tremtric.Value = True
Form23.Visible = False
End Sub

