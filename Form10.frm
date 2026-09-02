VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Estabilidade"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.TextBox washout 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox diedro 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox LRPN 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   5040
         Width           =   615
      End
      Begin VB.TextBox htail 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         TabIndex        =   46
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox sustentaht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   5040
         Width           =   615
      End
      Begin VB.TextBox sustentasa 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox LRCG 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox LRCAht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox LRCAasa 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   5640
         Width           =   615
      End
      Begin VB.TextBox MARGEM 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox PNCG 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox CG 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox PN 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox LT 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox CAht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox CAasa 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox DEDA 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox eficiht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox incidaht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox incidasa 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   5640
         Width           =   615
      End
      Begin VB.TextBox volvt 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox volht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox yht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox yasa 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Torção (geométrica) - washout em graus na ponta da asa :                     º"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6000
         TabIndex        =   51
         Top             =   7440
         Width           =   5145
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diedro da asa (em graus):                                                                      º"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   49
         Top             =   7440
         Width           =   5025
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distancia do eixo da hélice ao Ponto Neutro :                                         m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6000
         TabIndex        =   47
         Top             =   5040
         Width           =   5175
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localização do Estabilizador horizontal / Linha da Asa ( ht ) :                m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   45
         Top             =   6840
         Width           =   5115
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "( Variação downwash estab. horiz. / alfa da asa )"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   44
         Top             =   3840
         Width           =   3495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         X1              =   5760
         X2              =   5760
         Y1              =   480
         Y2              =   7800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lt ......."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   6000
         TabIndex        =   43
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Força de Sustentação na Empenagem horizontal :                                kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   40
         Top             =   5040
         Width           =   5205
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Força de Sustentação na Asa :                                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   39
         Top             =   4440
         Width           =   5220
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distancia do eixo da hélice ao Centro de Gravidade (CG) :                      m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6000
         TabIndex        =   20
         Top             =   6840
         Width           =   5205
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localização da Corda média aerodinamica da empenagem :                 m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   1440
         Width           =   5115
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localização da Corda média aerodinamica da Asa :                              m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   840
         Width           =   5100
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "distancia entre CAasa e CAht em metros :                                m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6720
         TabIndex        =   17
         Top             =   2040
         Width           =   4485
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distancia do eixo da hélice ao Centro aerod. da Asa (Nariz) :                  m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6000
         TabIndex        =   16
         Top             =   5640
         Width           =   5190
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro aerodinâmico da empenagem - corda média  (CAht) :                   m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6000
         TabIndex        =   15
         Top             =   1440
         Width           =   5190
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro aerodinâmico da Asa  - corda média (CAasa) :                              m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6000
         TabIndex        =   14
         Top             =   840
         Width           =   5205
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Caractetisticas geometricas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   5040
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Centro de gravidade e Ponto neutro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   6000
         TabIndex        =   12
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distancia do Ponto Neutro ao Centro de Gravidade em cm. :                  m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6000
         TabIndex        =   11
         Top             =   4440
         Width           =   5190
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentagem de Margem Estática :                                                          %"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6000
         TabIndex        =   10
         Top             =   3840
         Width           =   5220
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localização do Centro de Gravidade CG :                                               %"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6000
         TabIndex        =   9
         Top             =   3240
         Width           =   5190
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localização do Ponto Neutro :                                                                 %"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6000
         TabIndex        =   8
         Top             =   2640
         Width           =   5205
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "de / da "
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   3840
         Width           =   585
      End
      Begin VB.Label label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fator de Eficiencia do Estabilizador horizontal :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   3240
         Width           =   3285
      End
      Begin VB.Label label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angulo de incidencia do Estabilizador horizontal :                                  º"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   6240
         Width           =   5040
      End
      Begin VB.Label label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angulo de incidencia da Asa (em graus) :                                              º"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   5640
         Width           =   5025
      End
      Begin VB.Label label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volume do Estabilizador vertical :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   2640
         Width           =   2325
      End
      Begin VB.Label label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volume do Estabilizador horizontal :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   2040
         Width           =   2490
      End
      Begin VB.Label label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distancia do eixo da hélice ao Centro aerod. do Estab. horiz. :                m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6000
         TabIndex        =   1
         Top             =   6240
         Width           =   5205
      End
   End
   Begin VB.Menu fileir10 
      Caption         =   "&Ir para ..."
      Begin VB.Menu fileestrut 
         Caption         =   "Estruturas"
      End
      Begin VB.Menu fileperform10 
         Caption         =   "Performance"
      End
      Begin VB.Menu fileaerodinam10 
         Caption         =   "Aerodinâmica"
      End
      Begin VB.Menu filedimension10 
         Caption         =   "Dimensionamento"
      End
      Begin VB.Menu fileorincipia10 
         Caption         =   "Página inicial"
      End
      Begin VB.Menu filedesenho10 
         Caption         =   "Retornar ao desenho"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu visualiza17 
      Caption         =   "&Visualizar"
   End
   Begin VB.Menu filefer10 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "&Calculadora && Conversor"
      End
   End
   Begin VB.Menu fileimprimi10 
      Caption         =   "Im&primir"
   End
   Begin VB.Menu sai10 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form11.Visible = True
End Sub

Private Sub fileaerodinam10_Click()
Form8.Visible = False
Form7.Visible = True
Form7.Refresh
End Sub

Private Sub filecalcon_Click()
Form29.Show
End Sub

Private Sub filedesenho10_Click()
Form18.Visible = True
Form10.Visible = False
Form18.Refresh
End Sub

Private Sub filedimension10_Click()
Form8.Visible = False
Form2.Visible = True
Form2.Refresh
End Sub

Private Sub fileestrut_Click()
Form11.Visible = True
Form10.Visible = False

'Calculo das cargas com LOOP
cl = Form7.clk4

'CONVERSÃO P/ POLEGADA
swcalc = Form2.Sw * 1550
bwcalc = (Form2.bw / 2) * 39.37
crcalc = Form2.crw * 39.37
cpcalc = Form2.cpw * 39.37

dar = Form1.roar * 0.00194
q = ((dar / 2) * ((Form8.vmax * 0.911) ^ 2))

For estaca = 0 To bwcalc Step 0.01
'1
Y = estaca
Y = Format(Y, "0.000")

'2
yd = Y / (bwcalc)
yd = Format(yd, "0.000")

'3
c = crcalc * (1 + ((Form1.Afilasa - 1) * yd))
c = Format(c, "0.000")

'4
cla = ((((1 + (4 * swcalc)) / (3.1415 * (bwcalc) * c) * (1 - (yd ^ 2)) ^ 0.5)) / 2) * cl
cla = Format(cla, "0.000")


'VALORES DE CL

If yd = 0# Then
cl0 = cla
cl0 = Format(cl0, "0.000")
Form11.cl0 = cl0
Form11.q0 = c * 0.0833 * cl0 * q * 1.488
Form11.q0 = Format(Form11.q0, "0.000")
Form11.e0 = 0
Form11.e0 = Format(Form11.e0, "0.00")
End If

If yd = 0.1 Then
cl1 = cla
cl1 = Format(cl1, "0.000")
Form11.cl1 = cl1
Form11.q1 = c * 0.0833 * cl1 * q * 1.488
Form11.q1 = Format(Form11.q1, "0.000")
Form11.e1 = 0.1 * bwcalc / 39.37
Form11.e1 = Format(Form11.e1, "0.00")
End If

If yd = 0.2 Then
cl2 = cla
cl2 = Format(cl2, "0.000")
Form11.cl2 = cl2
Form11.q2 = c * 0.0833 * cl2 * q * 1.488
Form11.q2 = Format(Form11.q2, "0.000")
Form11.e2 = 0.2 * bwcalc / 39.37
Form11.e2 = Format(Form11.e2, "0.00")
End If

If yd = 0.3 Then
cl3 = cla
cl3 = Format(cl3, "0.000")
Form11.cl3 = cl3
Form11.q3 = c * 0.0833 * cl3 * q * 1.488
Form11.q3 = Format(Form11.q3, "0.000")
Form11.e3 = 0.3 * bwcalc / 39.37
Form11.e3 = Format(Form11.e3, "0.00")
End If

If yd = 0.4 Then
cl4 = cla
cl4 = Format(cl4, "0.000")
Form11.cl4 = cl4
Form11.q4 = c * 0.0833 * cl4 * q * 1.488
Form11.q4 = Format(Form11.q4, "0.000")
Form11.e4 = 0.4 * bwcalc / 39.37
Form11.e4 = Format(Form11.e4, "0.00")
End If

If yd = 0.5 Then
cl5 = cla
cl5 = Format(cl5, "0.000")
Form11.cl5 = cl5
Form11.q5 = c * 0.0833 * cl5 * q * 1.488
Form11.q5 = Format(Form11.q5, "0.000")
Form11.e5 = 0.5 * bwcalc / 39.37
Form11.e5 = Format(Form11.e5, "0.00")
End If

If yd = 0.6 Then
cl6 = cla
cl6 = Format(cl6, "0.000")
Form11.cl6 = cl6
Form11.q6 = c * 0.0833 * cl6 * q * 1.488
Form11.q6 = Format(Form11.q6, "0.000")
Form11.e6 = 0.6 * bwcalc / 39.37
Form11.e6 = Format(Form11.e6, "0.00")
End If

If yd = 0.7 Then
cl7 = cla
cl7 = Format(cl7, "0.000")
Form11.cl7 = cl7
Form11.q7 = c * 0.0833 * cl7 * q * 1.488
Form11.q7 = Format(Form11.q7, "0.000")
Form11.e7 = 0.7 * bwcalc / 39.37
Form11.e7 = Format(Form11.e7, "0.00")
End If

If yd = 0.8 Then
cl8 = cla
cl8 = Format(cl8, "0.000")
Form11.cl8 = cl8
Form11.q8 = c * 0.0833 * cl8 * q * 1.488
Form11.q8 = Format(Form11.q8, "0.000")
Form11.e8 = 0.8 * bwcalc / 39.37
Form11.e8 = Format(Form11.e8, "0.00")
End If

If yd = 0.9 Then
cl9 = cla
cl9 = Format(cl9, "0.000")
Form11.cl9 = cl9
Form11.q9 = c * 0.0833 * cl9 * q * 1.488
Form11.q9 = Format(Form11.q9, "0.000")
Form11.e9 = 0.9 * bwcalc / 39.37
Form11.e9 = Format(Form11.e9, "0.00")
End If

If yd = 1 Then
cl10 = cla
cl10 = Format(cl10, "0.000")
Form11.cl10 = cl10
Form11.q10 = c * 0.0833 * cl10 * q * 1.488
Form11.q10 = Format(Form11.q10, "0.000")
Form11.e10 = 1 * bwcalc / 39.37
Form11.e10 = Format(Form11.e10, "0.00")
End If
Next

'Cortante
If Form11.q10 = Empty Then Form11.q10 = 0

av10 = 0
av9 = ((Abs(Form11.q10) + Abs(Form11.q9)) / 2) * (0.1 * Form2.bw / 2)
av8 = ((Abs(Form11.q9) + Abs(Form11.q8)) / 2) * (0.1 * Form2.bw / 2)
av7 = ((Abs(Form11.q8) + Abs(Form11.q7)) / 2) * (0.1 * Form2.bw / 2)
av6 = ((Abs(Form11.q7) + Abs(Form11.q6)) / 2) * (0.1 * Form2.bw / 2)
av5 = ((Abs(Form11.q6) + Abs(Form11.q5)) / 2) * (0.1 * Form2.bw / 2)
av4 = ((Abs(Form11.q5) + Abs(Form11.q4)) / 2) * (0.1 * Form2.bw / 2)
av3 = ((Abs(Form11.q4) + Abs(Form11.q3)) / 2) * (0.1 * Form2.bw / 2)
av2 = ((Abs(Form11.q3) + Abs(Form11.q2)) / 2) * (0.1 * Form2.bw / 2)
av1 = ((Abs(Form11.q2) + Abs(Form11.q1)) / 2) * (0.1 * Form2.bw / 2)
av0 = ((Abs(Form11.q1) + Abs(Form11.q0)) / 2) * (0.1 * Form2.bw / 2)

Form11.v10 = "0,000"

Form11.v9 = Abs(v10) + Abs(av9)
Form11.v9 = Format(Form11.v9, "0.000")

Form11.v8 = Abs(Form11.v9) + Abs(av8)
Form11.v8 = Format(Form11.v8, "0.000")

Form11.v7 = Abs(Form11.v8) + Abs(av7)
Form11.v7 = Format(Form11.v7, "0.000")

Form11.v6 = Abs(Form11.v7) + Abs(av6)
Form11.v6 = Format(Form11.v6, "0.000")

Form11.v5 = Abs(Form11.v6) + Abs(av5)
Form11.v5 = Format(Form11.v5, "0.000")

Form11.v4 = Abs(Form11.v5) + Abs(av4)
Form11.v4 = Format(Form11.v4, "0.000")

Form11.v3 = Abs(Form11.v4) + Abs(av3)
Form11.v3 = Format(Form11.v3, "0.000")

Form11.v2 = Abs(Form11.v3) + Abs(av2)
Form11.v2 = Format(Form11.v2, "0.000")

Form11.v1 = Abs(Form11.v2) + Abs(av1)
Form11.v1 = Format(Form11.v1, "0.000")

Form11.v0 = Abs(Form11.v1) + Abs(av0)
Form11.v0 = Format(Form11.v0, "0.000")

'Fletor

am10 = 0
am9 = ((Abs(Form11.v10) + Abs(Form11.v9)) / 2) * (0.1 * Form2.bw / 2)
am8 = ((Abs(Form11.v9) + Abs(Form11.v8)) / 2) * (0.1 * Form2.bw / 2)
am7 = ((Abs(Form11.v8) + Abs(Form11.v7)) / 2) * (0.1 * Form2.bw / 2)
am6 = ((Abs(Form11.v7) + Abs(Form11.v6)) / 2) * (0.1 * Form2.bw / 2)
am5 = ((Abs(Form11.v6) + Abs(Form11.v5)) / 2) * (0.1 * Form2.bw / 2)
am4 = ((Abs(Form11.v5) + Abs(Form11.v4)) / 2) * (0.1 * Form2.bw / 2)
am3 = ((Abs(Form11.v4) + Abs(Form11.v3)) / 2) * (0.1 * Form2.bw / 2)
am2 = ((Abs(Form11.v3) + Abs(Form11.v2)) / 2) * (0.1 * Form2.bw / 2)
am1 = ((Abs(Form11.v2) + Abs(Form11.v1)) / 2) * (0.1 * Form2.bw / 2)
am0 = ((Abs(Form11.v1) + Abs(Form11.v0)) / 2) * (0.1 * Form2.bw / 2)

Form11.m10 = "0,000"

Form11.m9 = Abs(m10) + Abs(am9)
Form11.m9 = Format(Form11.m9, "0.000")

Form11.m8 = Abs(Form11.m9) + Abs(am8)
Form11.m8 = Format(Form11.m8, "0.000")

Form11.m7 = Abs(Form11.m8) + Abs(am7)
Form11.m7 = Format(Form11.m7, "0.000")

Form11.m6 = Abs(Form11.m7) + Abs(am6)
Form11.m6 = Format(Form11.m6, "0.000")

Form11.m5 = Abs(Form11.m6) + Abs(am5)
Form11.m5 = Format(Form11.m5, "0.000")

Form11.m4 = Abs(Form11.m5) + Abs(am4)
Form11.m4 = Format(Form11.m4, "0.000")

Form11.m3 = Abs(Form11.m4) + Abs(am3)
Form11.m3 = Format(Form11.m3, "0.000")

Form11.m2 = Abs(Form11.m3) + Abs(am2)
Form11.m2 = Format(Form11.m2, "0.000")

Form11.m1 = Abs(Form11.m2) + Abs(am1)
Form11.m1 = Format(Form11.m1, "0.000")

Form11.m0 = Abs(Form11.m1) + Abs(am0)
Form11.m0 = Format(Form11.m0, "0.000")

'Abrindo o salvado do form13
If Form21.Data1.Recordset.RecordCount < 1 Then GoTo 10

If Form21.Data1.Recordset("Projeto") = Form20.p1 Then
    If Form21.Data1.Recordset("checkCG") = 1 And Form13.checkg.Value = False Then
    Form13.Checkcgorig.Value = True
    Form13.porclong = Form21.Data1.Recordset("Hlong")
    Form13.gsatual = Form21.Data1.Recordset("CGatual")
    Form13.gsnegatual = Form21.Data1.Recordset("CGnega")
    End If
    
    If Form21.Data1.Recordset("checkCG") = 1 And Form13.checkg.Value = True Then
    Form13.checkg.Value = True
    Form13.gsatual = Form13.gsatual
    Form13.gsatual = Format(Form13.gsatual, "0.00")
    Form13.gsnegatual = Form13.gsnegatual
    Form13.gsnegatual = Format(Form13.gsnegatual, "0.00")
    End If
    
    If Form21.Data1.Recordset("checkCG") = 0 Then
    Form13.checkg.Value = True
    Form13.porclong = Form21.Data1.Recordset("Hlong")
    Form13.gsatual = Form21.Data1.Recordset("CGatual")
    Form13.gsnegatual = Form21.Data1.Recordset("CGnega")
    End If
    
End If

'Calculo do Nº de Gs no form13
If Form21.Data1.Recordset("Projeto") <> Form20.p1 Then
10:
    If Form13.Checkcgorig.Value = True Then
        Form13.gsatual = Form7.cl3dasa * ((dar / 2) * (Form2.Sw * 10.76) * ((Form9.velomaxima * 0.9113) ^ 2)) / (Form1.wo * 2.2)
        Form13.gsatual = Format(Form13.gsatual, "0.00")
        Form13.gsnegatual = 0.8 * Form7.cl3dasa * ((dar / 2) * (Form2.Sw * 10.76) * ((Form9.velomaxima * 0.9113) ^ 2)) / (Form1.wo * 2.2)
        Form13.gsnegatual = Format(Form13.gsnegatual, "0.00")
    End If
    
    If Form13.checkg.Value = True Then
        Form13.gsatual = Form13.gsatual
        Form13.gsatual = Format(Form13.gsatual, "0.00")
        Form13.gsnegatual = Form13.gsnegatual
        Form13.gsnegatual = Format(Form13.gsnegatual, "0.00")
    End If
    
    If Form13.Checkcgorig.Value = False And Form13.checkg.Value = False Then
        Form13.Checkcgorig.Value = True
        Form13.gsatual = Form7.cl3dasa * ((dar / 2) * (Form2.Sw * 10.76) * ((Form9.velomaxima * 0.9113) ^ 2)) / (Form1.wo * 2.2)
        Form13.gsatual = Format(Form13.gsatual, "0.00")
        Form13.gsnegatual = 0.8 * Form7.cl3dasa * ((dar / 2) * (Form2.Sw * 10.76) * ((Form9.velomaxima * 0.9113) ^ 2)) / (Form1.wo * 2.2)
        Form13.gsnegatual = Format(Form13.gsnegatual, "0.00")
    End If
    
End If


'Calculo da altura da longarina para o form13
    Form13.altlong = (Form13.porclong / 100) * ((Form2.espessura / 100) * Form2.crw * 100)
    Form13.altlong = Format(Form13.altlong, "0.00")
    

Form13.Show
Form13.Refresh

If Form13.alma = Empty Then MsgBox "Agora faça as opções para o calculo da longarina"

End Sub

Private Sub fileimprimi10_Click()
If LIBER <> "KFAR0852" Then
    MsgBox "A rotina de impressão sómente estará disponível após a liberação permanente", 16, "Cancelamento de impressão": Exit Sub
End If

Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "IV - ESTABILIDADE"
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 10
Printer.Print Tab(20); "ABREVIATURAS"
Printer.Print Tab(20); "Etb.Horz. = Estabilizador horizontal"
Printer.Print Tab(20); "C.A-asa = Centro aerodinâmico da asa"
Printer.Print Tab(20); "C.A-ht = Centro aerodinâmico do estabilizador horizontal"
Printer.Print Tab(20); "C.G = Centro de gravidade"
Printer.Print Tab(20); "L.R = Linha de refrencia = eixo da hélice"
Printer.Print Tab(20); "L.A = Linha de corda da asa"
Printer.Print Tab(20); "P.N = Ponto neutro"
Printer.Print ""
Printer.Print ""
Printer.Print Tab(20); "Localização da corda média da asa:"; Tab(70); yasa; " m"
Printer.Print Tab(20); "Localização da corda média do Etb.Horz.:"; Tab(70); yht; " m"
Printer.Print Tab(20); "Volume da empenagem horizontal:"; Tab(70); volht;
Printer.Print Tab(20); "Volume da empenagem vertical:"; Tab(70); volvt;
Printer.Print Tab(20); "Fator de eficiencia do Etb.Horz.:"; Tab(70); eficiht;
Printer.Print Tab(20); "Variação do angulo de downwash da asa:"; Tab(70); DEDA;
Printer.Print Tab(20); "Força de sustentação da asa:"; Tab(70); sustentasa; " kgf"
Printer.Print Tab(20); "Força de sustentação do Etb.Horz.:"; Tab(70); sustentaht; " kgf"
Printer.Print Tab(20); "Angulo de incidencia da asa:"; Tab(70); incidasa; " º"
Printer.Print Tab(20); "Angulo de incidencia do Etb.Horz.:"; Tab(70); incidaht; " º"
Printer.Print Tab(20); "Localização do Etb.Horz. referente a L.A:"; Tab(70); htail; " m"
Printer.Print Tab(20); "Localização do C.A-asa na corda média:"; Tab(70); CAasa; " m"
Printer.Print Tab(20); "Localizaçãodo C.A-ht na corda média:"; Tab(70); CAht; " m"
Printer.Print Tab(20); "Distancia C.A da asa ao C.A-ht (LT):"; Tab(70); LT; " m"
Printer.Print Tab(20); "Localização do Ponto neutro P.N:"; Tab(70); PN; " %"
Printer.Print Tab(20); "Localização do Centro de gravidade C.G:"; Tab(70); cg; " %"
Printer.Print Tab(20); "Margem estática em porcentagem:"; Tab(70); MARGEM; " %"
Printer.Print Tab(20); "Distancia entre P.N e C.G:"; Tab(70); PNCG; " m"
Printer.Print Tab(20); "Distancia do L.R ao P.N:"; Tab(70); LRPN; " m"
Printer.Print Tab(20); "Distancia do L.R ao C.A da asa:"; Tab(70); LRCAasa; " m"
Printer.Print Tab(20); "Distancia do L.R ao C.A do Etb.Horz.:"; Tab(70); LRCAht; " m"
Printer.Print Tab(20); "Distancia do L.R ao C.G:"; Tab(70); LRCG; " m"
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print Tab(25); "NOTA:"
Printer.Print Tab(20); "O valor do C.G calculado nesta planilha deverá ser equivalente"
Printer.Print Tab(20); "ao calculado estáticamente em Pesos e Estações"
Printer.Print ""
Printer.Print ""
Printer.Print Tab(20); "Fim da página";
Printer.EndDoc

End Sub

Private Sub fileorincipia10_Click()
Form1.Visible = True
Form10.Visible = False

Form1.Refresh
End Sub

Private Sub fileperform10_Click()
Form10.Visible = False
Form8.Visible = True
Form7.Refresh
End Sub


Private Sub sai10_Click()
Form1.Visible = True
Form10.Visible = False

Form1.Refresh
End Sub

Private Sub visualiza17_Click()
Form17.Visible = True
End Sub
