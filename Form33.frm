VERSION 5.00
Begin VB.Form Form33 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sôbre o programa"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form33.frx":0000
   LinkTopic       =   "Form33"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Height          =   435
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   4875
      TabIndex        =   23
      Top             =   4680
      Width           =   4935
      Begin VB.CommandButton Command3 
         Caption         =   "Sair"
         Height          =   375
         Left            =   4320
         MouseIcon       =   "Form33.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   0
         Width           =   555
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Autor"
         Height          =   375
         Left            =   3360
         MouseIcon       =   "Form33.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Agradecimentos"
         Height          =   375
         Left            =   2040
         MouseIcon       =   "Form33.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Homenagem"
         Height          =   375
         Left            =   960
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Contato"
         Height          =   375
         Left            =   0
         MouseIcon       =   "Form33.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   4935
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AEROSOFT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1710
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone : (0xx) 17 3321 6028 - Barretos SP"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   3105
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "email : aerosoft@ibt-barretos.org.br"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   2460
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horário : 9:00 hs às 17:00 hs. (2ª a 6ª feira)"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   3045
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Correspondencia : Avenida 13 - Nº 60  Barretos SP  CEP 14.780-280"
         Height          =   435
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   3675
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2115
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   2280
      Width           =   4935
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   0
         ScaleHeight     =   2175
         ScaleWidth      =   4935
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   4935
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Santos Dumont"
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
            Left            =   3525
            TabIndex        =   25
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Image Image2 
            Height          =   1470
            Left            =   3555
            Picture         =   "Form33.frx":0F32
            Top             =   240
            Width           =   1245
         End
         Begin VB.Image Image1 
            Height          =   2025
            Left            =   60
            Picture         =   "Form33.frx":6FEC
            Top             =   60
            Width           =   3405
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   0
         ScaleHeight     =   2175
         ScaleWidth      =   4935
         TabIndex        =   14
         Top             =   0
         Width           =   4935
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ferdinando Faria - Projetista e construtor de aeromodelos."
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   4065
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Engº Joseph Kovács (T25 - Tucano - P51 Peregrino)."
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   3780
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aos amigos e colaboradores,"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   0
            Width           =   2040
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Eng. James R. Waterhouse - Waterhouse aeronáutica."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1200
            Width           =   3885
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "I.B.T - Instituto Barretos de Tecnologia."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1920
            Width           =   2775
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prof.Dr. Bolis Petrussanis - CTA."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   2265
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Profs. Dr.Michael G.Maunsell e Dr.Fernando M.Catalano USP."
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   4455
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carlos Alberto de Diniz Junqueira FAC Informática./Barretos SP."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   4500
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Eng. Michel Cury - 1º eng. de ensaio em vôo brasileiro (i.m)."
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   4185
         End
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Model Estimar foi criado inicialmente como ferramenta de apoio ao projeto de uma aeronave experimental construída pelo autor. "
         Height          =   435
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   4620
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form33.frx":1D8E2
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Alvaro Martins Abdalla é natural de Colina SP, formou se eng. civil em 1984 pela FEB - Barretos. "
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00000040&
         ForeColor       =   &H00000000&
         Height          =   1900
         Left            =   3240
         Picture         =   "Form33.frx":1D9FD
         ScaleHeight     =   1905
         ScaleWidth      =   1695
         TabIndex        =   21
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Produzido por AEROSOFT CNPJ 04.286.344/0001-60 "
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2010
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Reg. INPI Nº 00035911"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "software para projetos de aeromodelos"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2730
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Model Estimar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form33"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame3.Visible = False
Picture2.Visible = False
Picture1.Visible = True
Picture5.Visible = False
End Sub

Private Sub Command2_Click()
Frame3.Visible = True
Picture2.Visible = False
End Sub

Private Sub Command3_Click()
Form33.Visible = False
Frame3.Visible = True
Picture2.Visible = False
Picture5.Visible = False
End Sub

Private Sub Command4_Click()
Frame3.Visible = False
Picture2.Visible = True
Picture5.Visible = False
End Sub

Private Sub Command5_Click()
Frame3.Visible = False
Picture5.Visible = True
Picture1.Visible = True
Picture2.Visible = False
End Sub
