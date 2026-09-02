VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Desempenho"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   ControlBox      =   0   'False
   FillColor       =   &H80000000&
   ForeColor       =   &H8000000A&
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
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
      Begin VB.TextBox vcruzeirook 
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         TabIndex        =   90
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox fuel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox vdecolsf 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox toque2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox lpousf 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox ldecsf 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox vst2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox ldtw 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox TD 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox tradin 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox FDdecol 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox npas 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox Fdsubida 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox tracestatica 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Lpouso 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox Ldecol 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox vdecol 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox FDstol 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox potreqstol 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox FD 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox potreqsub 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox cafetria 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox potreq 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox sobrahp 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox potreqdec 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox angclimb 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox hpreal 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox velplan 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox climb 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox toque 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox vmax 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox vsubida 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox vstol 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox cahel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox NI 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox avahel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox pashel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox diahel 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox pesomot 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox rpmf1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox cafechal 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox hpnominal 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox mtr 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reservatório  combustivel :                  lt"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   89
         Top             =   2760
         Width           =   2805
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/flap:"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   4680
         TabIndex        =   87
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/flap:"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   4680
         TabIndex        =   86
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/flap:"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   4680
         TabIndex        =   85
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fator W/T em Vmáxima :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   76
         Top             =   3720
         Width           =   1770
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tração Dínâmica :                                         kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   75
         Top             =   3240
         Width           =   3405
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Força de arrasto  Vdecolagem :                     kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   73
         Top             =   1320
         Width           =   3390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de pás :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   71
         Top             =   5400
         Width           =   1170
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Força de arrasto  Vsubida :                            kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   69
         Top             =   1800
         Width           =   3390
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tração Estática :                                            kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   67
         Top             =   2760
         Width           =   3420
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relação L/D  por W/T em Vmax :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   64
         Top             =   6360
         Width           =   2415
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pouso      c/F:                         s/F:                 m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   63
         Top             =   5400
         Width           =   3315
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Decolag. 15m c/F:                  s/F:                 m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   62
         Top             =   4920
         Width           =   3315
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Gerais"
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
         Left            =   8040
         TabIndex        =   61
         Top             =   4320
         Width           =   3105
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Força e Tração"
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
         Left            =   8040
         TabIndex        =   60
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vdecolag.                                s/flap:                km/h"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   58
         Top             =   1320
         Width           =   3705
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Força de arrasto  Vestol :                               Kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   56
         Top             =   840
         Width           =   3405
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Potencia Requerida em Vestol :                         Hp"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   54
         Top             =   5400
         Width           =   3570
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Força de arrasto  Vmáxima :                           kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   52
         Top             =   2280
         Width           =   3405
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   7800
         X2              =   7800
         Y1              =   720
         Y2              =   7680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         X1              =   3600
         X2              =   3600
         Y1              =   720
         Y2              =   7680
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Potencias"
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
         Left            =   3840
         TabIndex        =   51
         Top             =   4320
         Width           =   3240
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Potencia Requerida em Vsubida                        Hp"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   49
         Top             =   6360
         Width           =   3570
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "índice - C.A.F.E triviathon"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   47
         Top             =   7320
         Width           =   1815
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Potencia Requerida em Vmáxima :                    Hp"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   41
         Top             =   6840
         Width           =   3540
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sobra de potencia Vstall / Vmáxima :                Hp"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   40
         Top             =   7320
         Width           =   3540
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Potencia Requerida em Vdecolag.:                   Hp"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   39
         Top             =   5880
         Width           =   3540
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angulo de subida :                                          º"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   38
         Top             =   5880
         Width           =   3270
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Potencia Disponível real do motor :                   Hp"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   37
         Top             =   4920
         Width           =   3555
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V toque                                   s/flap:                 km/h"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   30
         Top             =   1800
         Width           =   3720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidade de planeio :                                     km/h"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   29
         Top             =   3240
         Width           =   3705
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razão de subida (Climb) :                                  ft/min"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   28
         Top             =   3720
         Width           =   3735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidade máxima calculada :                          km/h"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   27
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidade ótima de subida :                             km/h"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   26
         Top             =   2280
         Width           =   3720
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vestol                                      s/flap:                km/h"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3840
         TabIndex        =   25
         Top             =   840
         Width           =   3690
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Velocidades"
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
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   3315
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Especificação :                                    polg."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   4920
         Width           =   3060
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eficiencia em cruzeiro :                        %"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   7320
         Width           =   2835
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avanço :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   6840
         Width           =   645
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passo :                                                 polg."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   6360
         Width           =   3075
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diametro :                                             polg."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   5880
         Width           =   3090
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Hélices"
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
         Left            =   150
         TabIndex        =   13
         Top             =   4320
         Width           =   3015
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso do motor :                                    kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   2955
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Veloc. de Cruzeiro 80% hp :                 km/h"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   3090
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RPM  - original :                                    Rpm"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   3075
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Potencia do motor :                              Hp"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   2940
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "índice - C.A.F.E challenge/10000"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   4
         Top             =   6840
         Width           =   2385
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motor :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Motorização"
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
         Left            =   210
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Menu fileir8 
      Caption         =   "&Ir para ..."
      Begin VB.Menu fileestab 
         Caption         =   "Estabilidade"
      End
      Begin VB.Menu fileaerodinam 
         Caption         =   "Aerodinamica"
      End
      Begin VB.Menu filedimension 
         Caption         =   "Dimensionamento"
      End
      Begin VB.Menu fileprincipia 
         Caption         =   "Página inicial"
      End
      Begin VB.Menu filedesenho8 
         Caption         =   "Retornar ao desenho"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu filevisual8 
      Caption         =   "&Visualizar"
   End
   Begin VB.Menu filefer8 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "&Calculadora && Conversor"
      End
   End
   Begin VB.Menu fileimprimi8 
      Caption         =   "&Imprimir"
   End
   Begin VB.Menu sai8 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fileaerodinam_Click()
Form8.Visible = False
Form7.Visible = True
Form7.Refresh
End Sub

Private Sub filecalcon_Click()
Form29.Show
End Sub

Private Sub filedesenho8_Click()
Form18.Visible = True
Form8.Visible = False
Form18.Refresh
End Sub

Private Sub filedimension_Click()
Form8.Visible = False
Form2.Visible = True
Form2.Refresh
End Sub

Private Sub fileestab_Click()

'Form1.Visible = False

'Distancia da LR ao BA da asa = LRBA
LRBA = Form10.LRCAasa - (0.25 * Form7.CMaerodw)

Form10.yasa = (Form2.bw / 6) * ((1 + (2 * Form1.Afilasa) / (1 + Form1.Afilasa)))
Form10.yasa = Format(Form10.yasa, "0.000")

Form10.yht = (Form2.bht / 6) * ((1 + (2 * Form1.AfilaHt) / (1 + Form1.AfilaHt)))
Form10.yht = Format(Form10.yht, "0.000")

' Formula do Stinton
Kht = (1 - (3.6 * (((Form7.cdtotvmax) ^ 0.5) / (Form10.LT / Form7.CMaerodw)))) * Form4.effHT

Form10.eficiht = Kht
Form10.eficiht = Format(Form10.eficiht, "0.00")

Form10.DEDA = 35 * ((Form7.slopeasa * (3.1416 / 180)) / Form1.arw)
Form10.DEDA = Format(Form10.DEDA, "0.000")

varislope = Form7.slopeht / Form7.slopeasa

'Também pode calcular Ponto Neutro por:
'Pncalc = 0.25 + (Form10.eficiht * Form10.volht * varislope * (1 - Form10.DEDA))
'Form10.PN = Pncalc * 100
'Form10.PN = Format(Form10.PN, "0.000")

'Calculo de PN por Raymer - % de corda
cmfuse = ((0.005 * (Form2.wfs ^ 2) * Form2.lfs) / (Form7.CMaerodw * Form2.Sw)) * (180 / 3.1415)

Pncalc2 = (((Form7.slopeasa * (Form10.LRCAasa / Form7.CMaerodw)) - (cmfuse)) + (Form10.eficiht * (Form2.sht / Form2.Sw) * Form7.slopeht * (1 - Form10.DEDA) * (Form10.LRCAht / Form7.CMaerodw))) / (Form7.slopeasa + (Form10.eficiht * (Form2.sht / Form2.Sw) * Form7.slopeht * (1 - Form10.DEDA)))
Form10.PN = (((Pncalc2 * Form7.CMaerodw) - LRBA) / Form7.CMaerodw) * 100
Form10.PN = Format(Form10.PN, "0.000")
'Fim Raymer


margemcalc = Form1.margestat
Form10.MARGEM = margemcalc
Form10.MARGEM = Format(Form10.MARGEM, "0.000")


cgcalc = ((Abs((Form10.PN / 100) * Form7.CMaerodw) - Abs((margemcalc / 100) * Form7.CMaerodw)) / Form7.CMaerodw)
Form10.cg = cgcalc * 100
Form10.cg = Format(Form10.cg, "0.000")

'sustentação na asa
'lw é distancia CG até CA
'ltotal é distancia lw+ LT

    CAasacalc = 0.25 * Form7.CMaerodw
    Form10.CAasa = CAasacalc
    Form10.CAasa = Format(Form10.CAasa, "0.000")
    
    LRCGcalc = ((Abs(Form10.LRCAasa) - Abs(Form10.CAasa)) + Abs((Form10.cg / 100) * Form7.CMaerodw))
    Form10.LRCG = LRCGcalc
    Form10.LRCG = Format(Form10.LRCG, "0.000")
    
    Form10.LRPN = ((Abs(Form10.LRCAasa) - Abs(Form10.CAasa)) + Abs((Form10.PN / 100) * Form7.CMaerodw))
    Form10.LRPN = Format(Form10.LRPN, "0.000")

'antigo em 02/02/2002 lw = Abs(0.25 * form7.CMaerodw * 100) - Abs((cgcalc) * (form7.CMaerodw * 100))
a = (Form10.LRPN * 100) - (Form10.LRCG * 100)
B = (Form10.LRCAht * 100) - (Form10.LRPN * 100)

Lw1 = Form1.wo * (a + B) / B
Lh1 = Form1.wo - Lw1

'sustentação na asa
Form10.sustentasa = (Lw1)
Form10.sustentasa = Format(Form10.sustentasa, "0")


'sustentação Estab. horiz.
Form10.sustentaht = (Lh1)
Form10.sustentaht = Format(Form10.sustentaht, "0")


' Incidencia da asa - usando "slope do avião"
epson = 1 + ((2 * Form1.Afilasa) / (Form1.arw * (1 + Form1.Afilasa)))
slopeaviao = (Form7.slopeasa * 0.995) / (epson + (Form7.slopeasa / (3.141516 * Form1.arw)))

incidasa1 = (Form7.clk4 / slopeaviao) + (Form7.alfasazero * 3.141516 / 180)
Form10.incidasa = incidasa1 * (180 / 3.141516)
Form10.incidasa = Format(Form10.incidasa, "0.000")


wcls = Form1.wo / (Form7.cl3dht * Form2.Sw)
clh = Form10.sustentaht / (wcls * Form2.sht)
Form10.incidaht = clh / Form7.slopeht * (3.141516 / 180)
Form10.incidaht = Format(Form10.incidaht, "0.000")
'Para calcular downwash também epson = -36.5 * (clh / Form1.arw)

CAhtcalc = 0.25 * Form7.CMaerodht
Form10.CAht = CAhtcalc
Form10.CAht = Format(Form10.CAht, "0.000")

PNCGcalc = (Abs(Form10.LRPN) - Abs(Form10.LRCG))
Form10.PNCG = PNCGcalc
Form10.PNCG = Format(Form10.PNCG, "0.000")

'Diedro
If Form1.posialta.Value = True Then Form10.diedro = ((0.3 / Form4.porcAIL) + (0.5 / Form4.porcEV)) / 2.1
If Form1.posibaixa.Value = True Then Form10.diedro = 1 + (0.3 / Form4.porcAIL) + (0.5 / Form4.porcEV)
If Form1.posimedia.Value = True Then Form10.diedro = ((0.3 / Form4.porcAIL) + (0.5 / Form4.porcEV)) / 1.5


Form10.diedro = Format(Form10.diedro, "0.00")

If Form1.Afilasa < 0.5 Then Form10.washout = "-3.0"
If Form1.Afilasa >= 0.5 And Form1.Afilasa <= 0.6 Then Form10.washout = "-2.00"
If Form1.Afilasa > 0.61 And Form1.Afilasa <= 0.7 Then Form10.washout = "-1.50"
If Form1.Afilasa > 0.71 And Form1.Afilasa <= 0.8 Then Form10.washout = "-1.00"
If Form1.Afilasa > 0.81 And Form1.Afilasa <= 1 Then Form10.washout = "0.00"

Form10.Visible = True
Form10.Refresh
Form8.Visible = False

End Sub

Private Sub fileimprimi8_Click()
If LIBER <> "KFAR0852" Then
    MsgBox "A rotina de impressão sómente estará disponível após a liberação permanente", 16, "Cancelamento de impressão": Exit Sub
End If

Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "III- PERFORMANCE"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "1 - MOTORIZAÇÃO"
Printer.FontSize = 10
Printer.Print Tab(20); "Especificação do motor:"; Tab(70); mtr;
Printer.Print Tab(20); "Deslocamento do motor:"; Tab(70); dslc; " ci"
Printer.Print Tab(20); "Peso do motor:"; Tab(70); pesomot; " kgf"
Printer.Print Tab(20); "HP nominal do motor - original:"; Tab(70); hpnominal; " HP"
Printer.Print Tab(20); "RPM nominal do motor - original:"; Tab(70); rpmf1; " RPM"
Printer.Print Tab(20); "HP nominal em estudo:"; Tab(70); hpreq; " HP"
Printer.Print Tab(20); "RPM nominal em estudo:"; Tab(70); rpmestudo; " RPM"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "2 - HÉLICE"
Printer.FontSize = 10
Printer.Print Tab(20); "Hélice em:"; Tab(70); tipohel;
Printer.Print Tab(20); "Especificada por:"; Tab(70); cahel;
Printer.Print Tab(20); "Diametro da hélice:"; Tab(70); diahel; " pol."
Printer.Print Tab(20); "Passo da hélice:"; Tab(70); pashel; " pol."
Printer.Print Tab(20); "Avanço da hélice:"; Tab(70); avahel; " pol."
Printer.Print Tab(20); "Eficiencia da hélice:"; Tab(70); NI; " %"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "3 - VELOCIDADE"
Printer.FontSize = 10
Printer.Print Tab(20); "Velocidade de stall sem flap:"; Tab(70); vstol; " Km/h"; "     com flap:"; Form8.vst2; " Km/h"
Printer.Print Tab(20); "Velocidade de decolagem sem flap:"; Tab(70); Form8.vdecolsf; " Km/h"; "     com flap:"; Form8.vdecol; " Km/h"
Printer.Print Tab(20); "Velocidade de subida (ideal):"; Tab(70); vsubida; " Km/h"
Printer.Print Tab(20); "Velocidade máxima:"; Tab(70); vmax; " Km/h"
Printer.Print Tab(20); "Velocidade de toque no pouso sem flap:"; Tab(70); Form8.toque2; " Km/h"; "      com flap:"; Form8.toque; " Km/h"
Printer.Print Tab(20); "Velocidade de planeio:"; Tab(70); velplan; " Km/h"
Printer.Print Tab(20); "Razão de subida:"; Tab(70); climb; " ft/min"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "4 - POTENCIA"
Printer.FontSize = 10
Printer.Print Tab(20); "Potencia disponivel real:"; Tab(70); hpreal; " HP"
Printer.Print Tab(20); "Potencia requerida em Vstall:"; Tab(70); Form8.potreqstol; " HP"
Printer.Print Tab(20); "Potencia requerida em Vdecolagem:"; Tab(70); potreqdec; " HP"
Printer.Print Tab(20); "Potencia requerida em Vsubida:"; Tab(70); potreqsub; " HP"
Printer.Print Tab(20); "Potencia requerida em Vmáxima:"; Tab(70); potreq; " HP"
Printer.Print Tab(20); "Sobra de potencia - Vstall/Vmáxima:"; Tab(70); hpreal; " HP"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "5 - FORÇA E TRAÇÃO"
Printer.FontSize = 10
Printer.Print Tab(20); "Força de arrasto em Vstall:"; Tab(70); Form8.FDstol; " kgf"
Printer.Print Tab(20); "Força de arrasto em Vdecolagem:"; Tab(70); FDdecol; " kgf"
Printer.Print Tab(20); "Força de arrasto em Vsubida:"; Tab(70); Fdsubida; " kgf"
Printer.Print Tab(20); "Força de arrasto em Vmáxima:"; Tab(70); FD; " kgf"
Printer.Print Tab(20); "Tração estática:"; Tab(70); tracestatica; " kgf"
Printer.Print Tab(20); "Tração dinâmica:"; Tab(70); tradin; " kgf"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "6 - GERAIS"
Printer.FontSize = 10
Printer.Print Tab(20); "Angulo de subida(ideal):"; Tab(70); angclimb; " º"
Printer.Print Tab(20); "Distancia de decolagem s/flap:"; Tab(70); Form8.ldecsf; " m"; "    com flap:"; Form8.Ldecol; " m"
Printer.Print Tab(20); "Distancia de pouso s/flap:"; Tab(70); Form8.lpousf; " m"; "    com flap:"; Form8.Lpouso; " m"
Printer.Print Tab(20); "Relação L/D e W/T:"; Tab(70); ldtw;
Printer.Print Tab(20); "Alcance:"; Tab(70); Requisitos.alcanceBD; " nm"
Printer.Print Tab(20); "Razão de subida:"; Tab(70); Form8.climb; " ft/min"
Printer.Print Tab(20); "CAFE challenge:"; Tab(70); Form8.cafechal;
Printer.Print Tab(20); "CAFE Triviathon:"; Tab(70); Form8.cafetria;
Printer.Print ""
Printer.Print Tab(20); "Fim da página";
Printer.EndDoc

End Sub

Private Sub fileprincipia_Click()
Form1.Visible = True
Form8.Visible = False

Form1.Refresh
End Sub


Private Sub filevisual8_Click()
Form16.Show
End Sub

Private Sub sai8_Click()
Form1.Visible = True
Form8.Visible = False
Form1.Refresh
End Sub
