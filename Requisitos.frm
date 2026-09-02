VERSION 5.00
Begin VB.Form Requisitos 
   Caption         =   "Requisitos"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form35"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame PesVaz 
      Caption         =   "Frações de Peso Vazio"
      ForeColor       =   &H00FF00FF&
      Height          =   4335
      Left            =   120
      TabIndex        =   160
      Top             =   2640
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Image Image1 
         Height          =   3915
         Left            =   120
         Picture         =   "Requisitos.frx":0000
         Top             =   240
         Width           =   5715
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Combustivel"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   4920
      TabIndex        =   147
      Top             =   2160
      Width           =   1095
      Begin VB.CheckBox alcool 
         Caption         =   "Alcool"
         Height          =   255
         Left            =   120
         TabIndex        =   148
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Construção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   240
      TabIndex        =   141
      Top             =   2160
      Width           =   4575
      Begin VB.OptionButton Tubalum 
         Caption         =   "T-alum"
         Height          =   195
         Left            =   1080
         TabIndex        =   149
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton tubo 
         Caption         =   "T-aço"
         Height          =   195
         Left            =   240
         TabIndex        =   142
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton madeira 
         Caption         =   "Mad."
         Height          =   255
         Left            =   2040
         TabIndex        =   145
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton composto 
         Caption         =   "Fibra"
         Height          =   255
         Left            =   3720
         TabIndex        =   144
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton metal 
         Caption         =   "Chapa"
         Height          =   255
         Left            =   2880
         TabIndex        =   143
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Aerodinâmica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4095
      Left            =   6120
      TabIndex        =   94
      Top             =   240
      Width           =   5535
      Begin VB.TextBox lfR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   130
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox bwR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   129
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox swBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "0"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   127
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox wsBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   126
         Text            =   "0"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox swR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   125
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox bwBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   124
         Text            =   "0"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox lfBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   123
         Text            =   "0"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command22 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   122
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton Command24 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   121
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton Command25 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   120
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox wsR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   119
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   118
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox shBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   117
         Text            =   "0"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox shR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   116
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox swetR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   115
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox swetBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "0"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Command21 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   113
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox ARvt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "3"
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox ARht 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   111
         Text            =   "4"
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox ARhtR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   110
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox ARvtR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4680
         TabIndex        =   109
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox Afilavt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   108
         Text            =   "0,5"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox AfilaHt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   107
         Text            =   "0,5"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox AfilahtR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   106
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox AfilavtR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4680
         TabIndex        =   105
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton Command23 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   104
         Top             =   2880
         Width           =   255
      End
      Begin VB.TextBox arwBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   103
         Text            =   "0"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox arwR 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Afilasa 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   101
         Text            =   "0"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox AfilasaR 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox ldloiterBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   99
         Text            =   "0"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox ldcruzBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   98
         Text            =   "0"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox ldcruzR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   97
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox ldloiterR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4680
         TabIndex        =   96
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Command17 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   95
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comp. da fuselagem :"
         Height          =   195
         Left            =   240
         TabIndex        =   140
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Envergadura da asa :"
         Height          =   195
         Left            =   240
         TabIndex        =   139
         Top             =   720
         Width           =   1530
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área da asa :"
         Height          =   195
         Left            =   240
         TabIndex        =   138
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AR Ht e AR Vt :"
         Height          =   195
         Left            =   240
         TabIndex        =   137
         Top             =   3240
         Width           =   1125
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carga alar (lb / ft2) :"
         Height          =   195
         Left            =   240
         TabIndex        =   136
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área da empg. horizontal :"
         Height          =   195
         Left            =   240
         TabIndex        =   135
         Top             =   1440
         Width           =   1845
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Afilamento Ht e Vt :"
         Height          =   195
         Left            =   240
         TabIndex        =   134
         Top             =   3600
         Width           =   1365
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área molhada (ft2)"
         Height          =   195
         Left            =   240
         TabIndex        =   133
         Top             =   2160
         Width           =   1290
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AR e Afilamento da asa :"
         Height          =   195
         Left            =   240
         TabIndex        =   132
         Top             =   2880
         Width           =   1755
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L / D cruzeiro e loiter :  "
         Height          =   195
         Left            =   240
         TabIndex        =   131
         Top             =   2520
         Width           =   1665
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Requisitos de Performance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5370
      Left            =   240
      TabIndex        =   41
      Top             =   3000
      Width           =   5775
      Begin VB.TextBox climbR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   80
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox VcmaxR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   79
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox hpR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   78
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox alcR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   77
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox tetoR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   76
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox tempR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   75
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox daterraR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   74
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox consmotczR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   73
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox ddecolR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   72
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox ddecolBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "0"
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox CBHPcz 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "0"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox daterraBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "0"
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox tloiterBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox tetomaxBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox alcanceBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "0"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox hpBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "0"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox vcruzBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "0"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox climbBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "0"
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox CBHPlt 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "0"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox consmotltR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   61
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox NIcruzR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   60
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox NIcruz 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "0"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox NIloiterR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   58
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox NIloiter 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "0"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   56
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   55
         Top             =   2640
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   54
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   53
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton Command7 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   52
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton Command8 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   51
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   50
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton Command10 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   49
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton Command13 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   48
         Top             =   4800
         Width           =   255
      End
      Begin VB.CommandButton Command14 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   47
         Top             =   4440
         Width           =   255
      End
      Begin VB.CommandButton Command15 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   46
         Top             =   4080
         Width           =   255
      End
      Begin VB.CommandButton Command16 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   45
         Top             =   3360
         Width           =   255
      End
      Begin VB.CommandButton Command20 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   44
         Top             =   3720
         Width           =   255
      End
      Begin VB.TextBox vstallBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "0"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox vstallR 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4440
         TabIndex        =   42
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do projeto"
         Height          =   195
         Left            =   4440
         TabIndex        =   168
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do parâmetro"
         Height          =   195
         Left            =   2400
         TabIndex        =   167
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distancia de aterrissagem (ft) :"
         Height          =   195
         Left            =   240
         TabIndex        =   93
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distancia de decolagem (ft) :"
         Height          =   195
         Left            =   240
         TabIndex        =   92
         Top             =   4440
         Width           =   2025
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Potencia do motor (HP) :"
         Height          =   195
         Left            =   240
         TabIndex        =   91
         Top             =   1560
         Width           =   1740
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alcance (NM) :"
         Height          =   195
         Left            =   240
         TabIndex        =   90
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teto máximo (ft) :"
         Height          =   195
         Left            =   240
         TabIndex        =   89
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo estimado - loiter (hs) :"
         Height          =   195
         Left            =   240
         TabIndex        =   88
         Top             =   840
         Width           =   2025
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidade - cruzeiro (MPH) :"
         Height          =   195
         Left            =   240
         TabIndex        =   87
         Top             =   3360
         Width           =   2070
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razão de subida (ft/min) :"
         Height          =   195
         Left            =   240
         TabIndex        =   86
         Top             =   4080
         Width           =   1830
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consumo spc cruzeiro (lb/h) :"
         Height          =   195
         Left            =   240
         TabIndex        =   85
         Top             =   1920
         Width           =   2070
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consumo spc loiter (lb/h) :"
         Height          =   195
         Left            =   240
         TabIndex        =   84
         Top             =   2280
         Width           =   1845
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efic. da hélice em cruzeiro :"
         Height          =   195
         Left            =   240
         TabIndex        =   83
         Top             =   2640
         Width           =   1950
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Efic. da hélice em loiter :"
         Height          =   195
         Left            =   240
         TabIndex        =   82
         Top             =   3000
         Width           =   1725
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidade - estol (MPH) :"
         Height          =   195
         Left            =   240
         TabIndex        =   81
         Top             =   3720
         Width           =   1845
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Valores Máximos - Médios e Mínimos para requisitos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   240
      TabIndex        =   26
      Top             =   240
      Width           =   5775
      Begin VB.TextBox Tipomot 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3480
         TabIndex        =   169
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox UAV 
         Caption         =   "UAV"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4680
         TabIndex        =   157
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Optionbc 
         Caption         =   "Médios"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2400
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Optioncc 
         Caption         =   "Máximos"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3480
         TabIndex        =   35
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Optionac 
         Caption         =   "Mínimos"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1200
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox anvbase 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   33
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tipo "
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox ttrem 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4680
         TabIndex        =   31
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox perfilBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   720
         TabIndex        =   30
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox motorBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         TabIndex        =   29
         Top             =   1440
         Width           =   1575
      End
      Begin VB.PictureBox Picture2 
         Height          =   305
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   1995
         TabIndex        =   27
         Top             =   960
         Width           =   2055
         Begin VB.HScrollBar HScroll3 
            Enabled         =   0   'False
            Height          =   265
            Left            =   0
            Max             =   16
            TabIndex        =   28
            Top             =   0
            Width           =   1995
         End
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Escolher aeronave"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   360
         TabIndex        =   146
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T. pouso:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4680
         TabIndex        =   40
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   240
         TabIndex        =   39
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2880
         TabIndex        =   38
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aeronave paramêtro:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   2400
         TabIndex        =   37
         Top             =   720
         Width           =   1590
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Pesos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4035
      Left            =   6120
      TabIndex        =   0
      Top             =   4320
      Width           =   5535
      Begin VB.TextBox WEWOP 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1920
         TabIndex        =   161
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Kpv 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4440
         TabIndex        =   158
         Top             =   1320
         Width           =   855
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   255
         Left            =   3480
         Max             =   30
         TabIndex        =   156
         Top             =   960
         Value           =   1
         Width           =   855
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   255
         Left            =   1320
         Max             =   30
         TabIndex        =   155
         Top             =   960
         Value           =   1
         Width           =   855
      End
      Begin VB.CheckBox AltVariav 
         Caption         =   "f(We)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   152
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Exp 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4440
         TabIndex        =   151
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox FatC 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   150
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox woBD 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox pesoCR 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox pesoVR 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox weR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox wfR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox woestimadoR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox CPR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox CutilR 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Pocup 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   4080
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   7
         Top             =   600
         Width           =   1215
         Begin VB.HScrollBar HScroll2 
            Height          =   215
            LargeChange     =   5
            Left            =   0
            Max             =   10000
            SmallChange     =   5
            TabIndex        =   8
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   255
         Left            =   4080
         ScaleHeight     =   195
         ScaleWidth      =   1155
         TabIndex        =   5
         Top             =   240
         Width           =   1215
         Begin VB.HScrollBar HScroll1 
            Height          =   215
            Left            =   0
            Max             =   21
            TabIndex        =   6
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.TextBox bag 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox ocup 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   5115
         TabIndex        =   1
         Top             =   3480
         Width           =   5175
         Begin VB.CommandButton Command11 
            Caption         =   "AUTO"
            Height          =   315
            Left            =   4440
            TabIndex        =   166
            Top             =   0
            Width           =   685
         End
         Begin VB.CommandButton PezVazio 
            Caption         =   "Tabela"
            Height          =   315
            Left            =   3425
            TabIndex        =   165
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Tipo"
            Height          =   315
            Left            =   2400
            TabIndex        =   164
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Ok !"
            Height          =   315
            Left            =   1375
            TabIndex        =   163
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton Raymer 
            Caption         =   "Calcular Pesos"
            Height          =   315
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   1395
         End
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "We/Wo do parametro"
         Height          =   195
         Left            =   240
         TabIndex        =   162
         Top             =   1320
         Width           =   1560
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "We/Wo calculado"
         Height          =   195
         Left            =   3000
         TabIndex        =   159
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e"
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
         Left            =   3240
         TabIndex        =   154
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C"
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
         Left            =   1080
         TabIndex        =   153
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Vazio (lb) :"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1920
         Width           =   1140
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Combustível (lb) :"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   2640
         Width           =   1635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Total de decolagem  (lb) :"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   3000
         Width           =   2205
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carga Útil (lb) :"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ANV - exemplo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2640
         TabIndex        =   21
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ANV - projeto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4080
         TabIndex        =   20
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bagagem (lb) :"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de ocupantes e peso (lb) :"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2070
      End
   End
End
Attribute VB_Name = "Requisitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub alcool_Click()
If alcool.Value = Checked Then hpR = hpBD * 1.2
If alcool.Value = Unchecked Then hpR = hpBD
End Sub

Private Sub alcR_GotFocus()
alcR.SelStart = 0
alcR.SelLength = Len(alcR)
End Sub

Private Sub AltVariav_Click()
FatC.SetFocus
If AltVariav.Value = Unchecked Then
    HScroll4.Value = 0
    HScroll5.Value = 0
End If
End Sub

Private Sub climbR_GotFocus()
climbR.SelStart = 0
climbR.SelLength = Len(climbR)
End Sub



Private Sub fileacondionar_Click()
If tubo.Value = False And madeira.Value = False And metal.Value = False And composto.Value = False Then
        MsgBox "Escolha o tipo de construção", vbInformation
    Exit Sub
End If

' Coeficiente K
k = 1
10:
k = k + 1
'Coversão de velocidade
Vfts = VcmaxR * 1.4666

'Valor de e
EP = 2.718281

'IMPORTANTE

20:
    alcR = alcR - 10
    
 ' COMBUSTÍVEL

        'Calculo de K4 Cruzeiro
        k4 = EP ^ ((-alcR * 6076) * (consmotczR / 3600) / ((550 * NIcruzR) * ldcruzR))
        k4 = Format(k4, "0.000")
                
                'Calculo de K5 Loiter
        k5 = EP ^ ((-tempR * 60) * Vfts * (consmotltR / 3600) / ((550 * NIloiterR) * ldloiterR))
        k5 = Format(k5, "0.000")
        
'xxxxxxxxxx Calculo da porcentTRBem total xxxxxxxxxxx
ktotal = 0.975 * 0.985 * 0.995 * k4 * k5
ktotal = Format(ktotal, "0.000")
          
'xxxxxxxxxx Calculo da porcentTRBem de Peso Combustivel xxxxxxxxxxx
kpc = 1.06 * (1 - ktotal)

'xxxxxxxxxx Calculo do PESO TOTAL xxxxxxxxxxx

For i = k * (Val(Pocup) + Val(bag)) To 25000 Step 1

If Tipo.Option2.Value = True Or Tipo.Option3.Value = True Or Tipo.Option4.Value = True Or Tipo.Option5.Value = True Then
    If tubo.Value = True Then g = 2.36: Gg = -0.18
    If madeira.Value = True Then g = 1.19: Gg = -0.09
    If metal.Value = True And ocup = 2 Or ocup = 4 Then g = 2.36: Gg = -0.18
    If metal.Value = True And ocup = 6 Then g = 2.36: Gg = -0.18
    If composto.Value = True Then g = 0.99: Gg = -0.09
    
            Kpv = (g * (i ^ Gg))
            Kpv = Format(Kpv, "0.00")
End If

If Tipo.Option1.Value = True Then
    If tubo.Value = True Then g = 0.85: Gg = -0.05
    If madeira.Value = True Then g = 0.85: Gg = -0.05
    If metal.Value = True Then g = 0.95: Gg = -0.05
    If composto.Value = True Then g = 0.72: Gg = -0.04
    
            Kpv = (g * (i ^ Gg))
            Kpv = Format(Kpv, "0.000")
End If

    kk = (1 - kpc - Kpv)
    
    WOcalc = (Val(Pocup) + Val(bag)) / kk
    WOcalc = Format(WOcalc, "0.00")
    
If WOcalc < 0 Then GoTo 10
 
   hh = i
 ' barra de rolTRBem
ld.Visible = True
ld.Max = 3 * woBD
ld = hh
    
If i = Int(WOcalc) Or i > Int(WOcalc) Then
woestimadoR = WOcalc
woestimadoR = Format(woestimadoR, "0.00")

'Calculo de Peso Vazio e Combustível
weR = Kpv * woestimadoR
weR = Format(weR, "0.00")

wfR = kpc * woestimadoR
wfR = Format(wfR, "0.00")

'Calculo da carga  util
CPR = woestimadoR - weR
CPR = Format(CPR, "0.00")

kcpr = Val(CPR) - (Val(Pocup) + Val(wfR) + Val(bag))

'IMPORTANTE
If woestimadoR > 1.05 * woBD Then GoTo 20

'If kcpr < 0 Then MsgBox "A quantidade de combustível ou bagTRBem deverá ser diminuida"

ld.Visible = False
ld.Value = 0

Exit Sub
End If
Next
End Sub

Private Sub fileRaxRo_Click()
RaxRo.Visible = True
End Sub



Private Sub Command1_Click()
Tipo.Visible = True
End Sub




Private Sub Command11_Click()
AltVariav.Value = Checked



For i = 1 To 30
    For j = 1 To 30

    HScroll4.Value = i
    HScroll5.Value = j
    
    Call Raymer_Click
    
    
    'Para Peso Máximo
    'If Abs(Val(woBD) - Val(woestimadoR)) < 200 Then Exit Sub
    
    'Para Peso Minimo
    If Abs(Val(pesoVR) - Val(weR)) < 200 Then Exit Sub


    Next j
Next i

End Sub

Private Sub composto_Click()
Requisitos.AltVariav = Checked
Requisitos.HScroll4 = 10
Requisitos.HScroll5 = 10
End Sub
Private Sub HScroll4_Change()
FatC = 1 + (HScroll4.Value / 10)
End Sub

Private Sub HScroll5_Change()
Exp = -1 * (HScroll5.Value / 100)
End Sub

Private Sub Image1_Click()
PesVaz.Visible = False
End Sub

Private Sub madeira_Click()
'Requisitos.AltVariav = Checked
'Requisitos.HScroll4 = 10
'Requisitos.HScroll5 = 10
End Sub

Private Sub metal_Click()
'Requisitos.AltVariav = Checked
'Requisitos.HScroll4 = WEWOP * 100
'Requisitos.HScroll5 = 10
End Sub

Private Sub PezVazio_Click()
PesVaz.Visible = True
End Sub

Private Sub Raymer_Click()

If Pocup = 0 Then
MsgBox "Insira o Número de ocupantes", vbCritical
Exit Sub
End If

If tubo.Value = False And madeira.Value = False And metal.Value = False And composto.Value = False And Tubalum.Value = False Then
        MsgBox "Escolha o tipo de construção", vbInformation
    Exit Sub
End If

' Coeficiente K
k = 1
10:
k = k + 1
'Coversão de velocidade
Vfts = VcmaxR * 1.4666

'Valor de e
EP = 2.718281

' COMBUSTÍVEL

    'Calculo de K4 Cruzeiro
    r7 = EP ^ ((-alcR * 6076) * (consmotczR / 3600) / ((550 * NIcruzR) * ldcruzR))
    r7 = Format(r7, "0.000")
        'CczR = (consmotczR) * VcmaxR / (550 * NIcruz)
        'CczR = Format(CczR, "0.000")
        'k4 = EP ^ ((-alcR * CczR) / (VcmaxR * ldcruzR))
        'k4 = Format(k4, "0.000")
        
        'Calculo de K5 Loiter
    r8 = EP ^ ((-tempR * 60) * Vfts * (consmotltR / 3600) / ((550 * NIloiterR) * ldloiterR))
    r8 = Format(r8, "0.000")
        'CltR = (consmotltR) * VcmaxR / (550 * NIloiterR)
        'CltR = Format(CltR, "0.0000")
        'k5 = EP ^ ((-tempR * CltR) / ldloiterR)
        'k5 = Format(k5, "0.000")
        
        
'xxxxxxxxxx Calculo da porcentTRBem total xxxxxxxxxxx
EP = 2.71828182846
r1 = 0.992
r2 = 0.996
r3 = 0.996
r4 = 0.99
r5 = 0.992
r6 = 0.992


rr = (r1 * r2 * r3 * r4 * r5 * r6)
rr = Format(rr, "0.000")
ktotal = rr * r7 * r8
ktotal = Format(ktotal, "0.000")
          
          
'xxxxxxxxxx Calculo da porcentTRBem de Peso Combustivel xxxxxxxxxxx
kpc = 1.06 * (1 - ktotal)

'xxxxxxxxxx Calculo do PESO TOTAL xxxxxxxxxxx

For i = k * (Val(Pocup) + Val(bag)) To 25000 Step 20

If AltVariav.Value = Checked Then ' Acrescentado em 02/11/2005
    If Tipo.Option1.Value = True Or Tipo.Option2.Value = True Or Tipo.Option3.Value = True Or Tipo.Option4.Value = True Or Tipo.Option5.Value = True Then
        If tubo.Value = True Then g = FatC: Gg = Exp: formula = "C x Wo ^ e"
        If madeira.Value = True Then g = FatC: Gg = Exp: j = 1: formula = "C x Wo ^ e"
        If metal.Value = True Then g = FatC: Gg = Exp: formula = "C x Wo ^ e"
        If composto.Value = True Then g = FatC: Gg = Exp: formula = "C x Wo ^ e"
            Kpv = (g * (i ^ Gg))
            Kpv = Format(Kpv, "0.000")
    End If
End If


'SUBTRAI O QUE ESTÁ ABAIXO EM 27/11/2007 PARA A TESE, O CÁLCULO DE PESO SERÁ FEITO À MÃO.
'If AltVariav.Value = Unchecked Then ' Acrescentado em 02/11/2005
    'If Tipo.Option2.Value = True Or Tipo.Option3.Value = True Or Tipo.Option4.Value = True Or Tipo.Option5.Value = True Then
        'If tubo.Value = True Then g = 2.36: Gg = -0.18: formula = "2,36 x Wo ^ -0,18"
        'If madeira.Value = True Then g = 1.19: Gg = -0.09: j = 1: formula = "1,19 x Wo ^ -0,09"
        'If metal.Value = True And ocup = 2 Or ocup = 4 Then g = 2.36: Gg = -0.18: formula = "2,36 x Wo ^ -0,18"
        'If metal.Value = True And ocup = 6 Then g = 2.36: Gg = -0.18: formula = "2,36 x Wo ^ -0,18"
        'If composto.Value = True Then g = 0.99: Gg = -0.09: formula = "0,99 x Wo ^ -0,09"
            'Kpv = (g * (i ^ Gg))
            'Kpv = Format(Kpv, "0.000")
    'End If
'End If

'If AltVariav.Value = Unchecked Then ' Acrescentado em 02/11/2005
    'If Tipo.Option1.Value = True Then
        'If tubo.Value = True Then g = 0.85: Gg = -0.05: formula = "0,85 x Wo ^ -0,05"
        'If madeira.Value = True Then g = 0.85: Gg = -0.05: formula = "0,85 x Wo ^ -0,05"
        'If metal.Value = True Then g = 0.95: Gg = -0.05: formula = "0,95 x Wo ^ -0,05"
        'If composto.Value = True Then g = 0.72: Gg = -0.04: formula = "0,72 x Wo ^ -0,04"
        'If Tubalum.Value = True Then g = 1#: Gg = -0.25: formula = "1,00 x Wo ^ -0,25"
            'Kpv = (g * (i ^ Gg))
            'Kpv = Format(Kpv, "0.000")
            
    'End If
'End If
'ATÉ AQUI XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX



    kk = (1 - kpc - Kpv)
    
    WOcalc = (Val(Pocup) + Val(bag)) / kk
    WOcalc = Format(WOcalc, "0.00")
    
If WOcalc < 0 Then GoTo 10
 
    hh = i
 ' barra de rolTRBem
'ld.Visible = True
'ld.Max = 3 * woBD
'ld = hh

If i = Int(WOcalc) Or i > Int(WOcalc) Then
woestimadoR = WOcalc
woestimadoR = Format(woestimadoR, "0.00")
Ra4 = woestimadoR

'Calculo de Peso Vazio e Combustível
weR = Kpv * woestimadoR
weR = Format(weR, "0.00")
Ra1 = weR

wfR = kpc * woestimadoR
wfR = Format(wfR, "0.00")
Ra3 = wfR

'Calculo da carga  util
CPR = woestimadoR - weR
CPR = Format(CPR, "0.00")
Ra2 = CPR


kcpr = Val(CPR) - (Val(Pocup) + Val(wfR) + Val(bag))

'If kcpr < 0 Then MsgBox "A quantidade de combustível ou bagTRBem deverá ser diminuida"

'ld.Visible = False
'ld.Value = 0


Exit Sub
End If
Next

End Sub

Private Sub Command10_Click()
NIloiterR = NIloiter
End Sub

Private Sub Command12_Click()
wsR = wsBD
End Sub

Private Sub Command13_Click()
daterraR = daterraBD
End Sub

Private Sub Command14_Click()
ddecolR = ddecolBD
End Sub

Private Sub Command15_Click()
climbR = climbBD
End Sub

Private Sub Command16_Click()
VcmaxR = vcruzBD
End Sub

Private Sub Command17_Click()
ldcruzR = ldcruzBD
ldloiterR = ldloiterBD
End Sub
Private Sub Command18_Click()
shR = swBD
End Sub

Private Sub Command19_Click()
clR = clcalc
cdoR = Cdocalc
End Sub

Private Sub Command2_Click()
alcR = alcanceBD
End Sub

Private Sub Command20_Click()
vstallR = vstallBD
End Sub

Private Sub Command21_Click()
swetR = swetBD
End Sub

Private Sub Command22_Click()
swR = swBD
End Sub

Private Sub Command23_Click()
arwR = arwBD
AfilasaR = Afilasa
End Sub

Private Sub Command24_Click()
lfR = lfBD
End Sub

Private Sub Command25_Click()
bwR = bwBD
End Sub

Private Sub Command26_Click()
clmaxR = clmaxBD
cdimaxR = cdimaxBD
End Sub



Private Sub Command29_Click()

End Sub

Private Sub Command3_Click()
NIcruzR = NIcruz
End Sub

Private Sub Command4_Click()
consmotltR = CBHPlt
End Sub

Private Sub Command5_Click()
consmotczR = CBHPcz
End Sub

Private Sub Command6_Click()
Call Raymer_Click
Requisitos.Hide
Form1.Show
Form1.wo = woestimadoR / 2.2
Form1.wo = Format(Form1.wo, "0.00")

'Form1.woanexo = woestimadoR / 2.2
Form1.wo = Format(Form1.wo, "0.00")

Form1.motor = hpR

 Form1.Command9.Enabled = True
 
If alcool.Value = Checked Then hpR = hpBD * 1.2

End Sub

Private Sub consmotR_GotFocus()
consmotR.SelStart = 0
consmotR.SelLength = Len(consmotR)
End Sub

Private Sub Command7_Click()
hpR = hpBD
End Sub

Private Sub Command8_Click()
tetoR = tetomaxBD
End Sub

Private Sub Command9_Click()
tempR = tloiterBD
End Sub

Private Sub daterraR_GotFocus()
daterraR.SelStart = 0
daterraR.SelLength = Len(daterraR)
End Sub

Private Sub ddecolR_GotFocus()
ddecolR.SelStart = 0
ddecolR.SelLength = Len(ddecolR)
End Sub

Private Sub filefechar_Click()
Requisitos.Hide
End Sub

Private Sub hpR_GotFocus()
hpR.SelStart = 0
hpR.SelLength = Len(hpR)
End Sub

Private Sub HScroll1_Change()
ocup = HScroll1.Value
Pocup = 209 * ocup
End Sub

Private Sub HScroll2_Change()
bag = HScroll2.Value
End Sub


Private Sub HScroll3_Change()
HScroll4.Value = 0
HScroll5.Value = 0

If HScroll3.Value = 1 And Tipo.av1 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av1 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av1
GoTo 10
End If

If HScroll3.Value = 2 And Tipo.av2 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av2 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av2
GoTo 10
End If

If HScroll3.Value = 3 And Tipo.av2 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av3 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av3
GoTo 10
End If

If HScroll3.Value = 4 And Tipo.av4 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av4 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av4
GoTo 10
End If

If HScroll3.Value = 5 And Tipo.av5 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av5 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av5
GoTo 10
End If

If HScroll3.Value = 6 And Tipo.av6 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av6 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av6
GoTo 10
End If

If HScroll3.Value = 7 And Tipo.av7 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av7 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av7
GoTo 10
End If

If HScroll3.Value = 8 And Tipo.av8 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av8 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av8
GoTo 10
End If

If HScroll3.Value = 9 And Tipo.av9 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av9 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av9
GoTo 10
End If

If HScroll3.Value = 10 And Tipo.av10 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av10 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av10
GoTo 10
End If

If HScroll3.Value = 11 And Tipo.av11 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av11 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av11
GoTo 10
End If

If HScroll3.Value = 12 And Tipo.av12 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av12 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av12
GoTo 10
End If

If HScroll3.Value = 13 And Tipo.av13 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av13 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av13
GoTo 10
End If

If HScroll3.Value = 14 And Tipo.av14 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av14 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av14
GoTo 10
End If

If HScroll3.Value = 15 And Tipo.av15 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av15 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av15
GoTo 10
End If

If HScroll3.Value = 16 And Tipo.av16 <> Empty Then
ProcNave = "Aeronave = '" & Tipo.av16 & "'"
Tipo.Data1.Recordset.FindFirst ProcNave
anvbase = Tipo.av16
GoTo 10
End If
            
10
    alcanceBD = Tipo.Data1.Recordset("alcance")
    tetomaxBD = Tipo.Data1.Recordset("teto")
    hpBD = Tipo.Data1.Recordset("hp")
    vcruzBD = Tipo.Data1.Recordset("vcruz")
    vstallBD = Tipo.Data1.Recordset("vstall")
    climbBD = Tipo.Data1.Recordset("climb")
    ddecolBD = Tipo.Data1.Recordset("ddecol")
    daterraBD = Tipo.Data1.Recordset("daterra")
    swBD = Tipo.Data1.Recordset("sw")
    bwBD = Tipo.Data1.Recordset("bw")
    lfBD = Tipo.Data1.Recordset("lfuse")
    shBD = Tipo.Data1.Recordset("sht")
    woBD = Tipo.Data1.Recordset("pesoT")
    pesoCR = Tipo.Data1.Recordset("pesoF")
    pesoVR = Tipo.Data1.Recordset("pesov")
    CutilR = woBD - pesoVR
    Afilasa = Tipo.Data1.Recordset("afilamento")
    perfilBD = Tipo.Data1.Recordset("perfil")
    Tipomot = Tipo.Data1.Recordset("tipoprop")
    If Tipo.Data1.Recordset("Motor") <> Empty Then motorBD = Tipo.Data1.Recordset("Motor")
    
    
    ' Calculos outros
tloiterBD = 5
tloiterBD = Format(tloiterBD, "0.00")

If Tipomot = "j" Then CBHPcz = 0.8 - (0.00096 * (pesoCR ^ 0.5))
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "tp" Then CBHPcz = 2.5 * hpBD ^ -0.1905
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "p" Then CBHPcz = 0.4
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "j" Then CBHPlt = 1.2 * CBHPcz 'Decolagem 0.49 - (0.0007 * (pesoCR ^ 0.5))
CBHPlt = Format(CBHPlt, "0.00")

If Tipomot = "tp" Then CBHPlt = 1.2 * CBHPcz
CBHPlt = Format(CBHPlt, "0.00")

If Tipomot = "p" Then CBHPlt = 0.5
CBHPlt = Format(CBHPlt, "0.00")


NIcruz = 0.8
NIcruz = Format(NIcruz, "0.00")

NIloiter = 0.7
NIloiter = Format(NIloiter, "0.00")

'Calculos para aerodinamica
lfuse = Tipo.Data1.Recordset("lfuse")
sht = Tipo.Data1.Recordset("sht")
svt = Tipo.Data1.Recordset("svt")
bw = Tipo.Data1.Recordset("bw")

ttrem = Tipo.Data1.Recordset("ttrem")
If ttrem = "retratil" Then
    TT = 10.864: P = 0.5218
    Form1.Checkretratil.Value = Checked
End If

If ttrem = "fixo" Then
    TT = 8.25: P = 0.61
    Form1.Checkretratil.Value = Unchecked
End If

'Area molhada excel(10,7*4*((0,074* Z181^2 + 0,05*Z181^3- 0,006*Z181^4)))+(2,05*M21)+(2,05*N21)+(2,05*(J21-16))

swetBD = Abs(1.15 * ((2.05 * (swBD - 16)) + (8.25 * lfuse) + (2.05 * sht) + (2.05 * svt)))
swetBD = Format(swetBD, "0.00")

ldcruzBD = TT * ((((bw ^ 2) / swBD) / (swetBD / swBD)) ^ P)
ldcruzBD = Format(ldcruzBD, "0.00")

ldloiterBD = ldcruzBD * 0.866
ldloiterBD = Format(ldloiterBD, "0.00")

arwBD = bwBD ^ 2 / swBD
arwBD = Format(arwBD, "0.00")

ARht = "N.D"
ARvt = "N.D"
AfilaHt = "N.D"
Afilavt = "N.D"
ARhtR = 4
ARvtR = 3
AfilahtR = 1
AfilavtR = 0.5

wsBD = woBD / swBD
wsBD = Format(wsBD, "0.00")

'Cdocalc = (2 * woBD / ldcruzBD) / (0.00238 * ((vcruzBD * 1.4666) ^ 2) * swBD)
'Cdocalc = Format(Cdocalc, "0.000")

'clcalc = (2 * woBD) / (0.00238 * ((vcruzBD * 1.4666) ^ 2) * swBD)
'clcalc = Format(clcalc, "0.000")

'clmaxBD = (2 * woBD) / (0.00238 * ((vstallBD * 1.4666) ^ 2) * swBD)
'clmaxBD = Format(clmaxBD, "0.00")

'cdimaxBD = (clcalc ^ 2) / (3.1415 * arwBD * 0.85)
'cdimaxBD = Format(cdimaxBD, "0.00")

'Calculo da carga  util
CutilR = woBD - pesoVR
CutilR = Format(CutilR, "0.00")


'YYYYYYYYYYYY Lançando Dados YYYYYYYYYYYYYYYY
'performance
daterraR = daterraBD
ddecolR = ddecolBD
climbR = climbBD
VcmaxR = vcruzBD
alcR = alcanceBD
consmotltR = CBHPlt
consmotczR = CBHPcz
NIcruzR = NIcruz
NIloiterR = NIloiter
hpR = hpBD
tetoR = tetomaxBD
tempR = tloiterBD
'aerodinamica
ldcruzR = ldcruzBD
ldloiterR = ldloiterBD
wsR = wsBD
swR = swBD
bwR = bwBD
arwR = arwBD
lfR = lfBD
shR = swBD
vstallR = vstallBD
swetR = swetBD
AfilasaR = Afilasa

'XXXXXXXXXXXXX FÓRMULA XXXXXXXXXXXXXXXX
For i = 0.01 To 4 Step 0.1
For j = 0.01 To 4 Step 0.1
d = i
e = j

wwe = (i * (Int(woBD) ^ (-j))) * Int(woBD)
wwe = Format(wwe, "0.00")

If Int(wwe) = Int(pesoVR) Or Int(wwe) > Int(pesoVR) And Int(wwe) < 1.05 * Int(pesoVR) Then

e = Format(e, "0.00")
d = Format(d, "0.00")

Exit Sub

End If
Next
Next


End Sub

Private Sub Option1_Click()
' Trazendo valores através de Alcance
'alcanceBD = Format(alcanceBD, "0")
'ProcNave = "alcance = " & Requisitos.alcanceBD & ""
'Tipo.Data1.Recordset.FindFirst ProcNave
'Requisitos.anvbase = Tipo.Data1.Recordset("aeronave")
'alcanceBD = Format(alcanceBD, "0.00")

HScroll3.Enabled = True
HScroll3.Value = 1
Requisitos.WEWOP = Requisitos.pesoVR / Requisitos.woBD
Requisitos.WEWOP = Format(Requisitos.WEWOP, "0.00")
End Sub

Private Sub Optionac_Click()
If Tipo.Option1.Value = True Then

tloiterBD = 5
tloiterBD = Format(tloiterBD, "0.00")

        Tipo.Data1.RecordSource = ("Select Min(PesoF) as pesoCR from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(PesoV) as pesoVR from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(teto) as tetomaxBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(hp) as hpBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(climb) as climbBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(sw) as swBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(PesoT) as woBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Bw) as bwBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(lfuse) as lfBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(sht) as shBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(vstall) as vstallBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
        
                Tipo.Option1.Value = False
                Tipo.Option1.Value = True
End If

If Tipo.Option2.Value = True Then

tloiterBD = 15
tloiterBD = Format(tloiterBD, "0.00")


        Tipo.Data1.RecordSource = ("Select Min(PesoF) as pesoCR from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(PesoV) as pesoVR from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(teto) as tetomaxBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(hp) as hpBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(climb) as climbBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(sw) as swBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(PesoT) as woBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Bw) as bwBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(lfuse) as lfBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(sht) as shBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(vstall) as vstallBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option2.Value = False
                Tipo.Option2.Value = True
        
End If


If Tipo.Option3.Value = True Then

tloiterBD = 15
tloiterBD = Format(tloiterBD, "0.00")


        Tipo.Data1.RecordSource = ("Select Min(PesoF) as pesoCR from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(PesoV) as pesoVR from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(teto) as tetomaxBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(hp) as hpBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(climb) as climbBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(sw) as swBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(PesoT) as woBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Bw) as bwBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(lfuse) as lfBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(sht) as shBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(vstall) as vstallBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option3.Value = False
                Tipo.Option3.Value = True
        
End If

If Tipo.Option4.Value = True Then

tloiterBD = 20
tloiterBD = Format(tloiterBD, "0.00")


        Tipo.Data1.RecordSource = ("Select Min(PesoF) as pesoCR from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(PesoV) as pesoVR from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(teto) as tetomaxBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(hp) as hpBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(climb) as climbBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(sw) as swBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(PesoT) as woBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Bw) as bwBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(lfuse) as lfBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(sht) as shBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(vstall) as vstallBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option4.Value = False
                Tipo.Option4.Value = True
        
End If

If Tipo.Option5.Value = True Then

tloiterBD = 20
tloiterBD = Format(tloiterBD, "0.00")


        Tipo.Data1.RecordSource = ("Select Min(PesoF) as pesoCR from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(PesoV) as pesoVR from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(teto) as tetomaxBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(hp) as hpBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(climb) as climbBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(sw) as swBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(PesoT) as woBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(Bw) as bwBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(lfuse) as lfBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(sht) as shBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Min(vstall) as vstallBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option5.Value = False
                Tipo.Option5.Value = True
 
End If

' Calculos outros


If Tipomot = "j" Then CBHPcz = 0.8 - (0.00096 * (pesoCR ^ 0.5))
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "tp" Then CBHPcz = 2.5 * hpBD ^ -0.1905
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "p" Then CBHPcz = 0.4
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "j" Then CBHPlt = 1.2 * CBHPcz 'Decolagem 0.49 - (0.0007 * (pesoCR ^ 0.5))
CBHPlt = Format(CBHPlt, "0.00")

If Tipomot = "tp" Then CBHPlt = 1.2 * CBHPcz
CBHPlt = Format(CBHPlt, "0.00")

If Tipomot = "p" Then CBHPlt = 0.5
CBHPlt = Format(CBHPlt, "0.00")

NIcruz = 0.8
NIcruz = Format(NIcruz, "0.00")

NIloiter = 0.7
NIloiter = Format(NIloiter, "0.00")

'Calculos para aerodinamica
clcalc = (2 * woBD) / (0.00238 * ((vcruzBD * 1.4666) ^ 2) * swBD)
clcalc = Format(clcalc, "0.00")

arwBD = bwBD ^ 2 / swBD
arwBD = Format(arwBD, "0.00")

clmaxBD = (2 * woBD) / (0.00238 * ((vstallBD * 1.4666) ^ 2) * swBD)
clmaxBD = Format(clmaxBD, "0.00")

Cdocalc = ((250 * 1.414 * (clcalc ^ 1.5) * (0.00238 ^ 0.5) * (swBD ^ 0.5) * hpBD * 0.65) / woBD ^ 1.5)
Cdocalc = Format(Cdocalc, "0.000")
If Tipo.Option5.Value = True Then Cdocalc = Format(Cdocalc, "0.00000")
If Tipo.Option5.Value = False Then Cdocalc = Format(Cdocalc, "0.00")

cdimaxBD = (clcalc ^ 2) / (3.1415 * arwBD * 0.85)
cdimaxBD = Format(cdimaxBD, "0.00")

ldcruzBD = clcalc / Cdocalc
ldcruzBD = Format(ldcruzBD, "0.00")

ldloiterBD = ldcruzBD * 0.866
ldloiterBD = Format(ldloiterBD, "0.00")

wsBD = woBD / swBD
wsBD = Format(wsBD, "0.00")

'Calculo da carga  util
CutilR = woBD - pesoVR
CutilR = Format(CutilR, "0.00")


'YYYYYYYYYYYY Lançando Dados YYYYYYYYYYYYYYYY
'performance
daterraR = daterraBD
ddecolR = ddecolBD
climbR = climbBD
VcmaxR = vcruzBD
alcR = alcanceBD
consmotltR = CBHPlt
consmotczR = CBHPcz
NIcruzR = NIcruz
NIloiterR = NIloiter
hpR = hpBD
tetoR = tetomaxBD
tempR = tloiterBD
'aerodinamica
ldcruzR = ldcruzBD
ldloiterR = ldloiterBD
wsR = wsBD
swR = swBD
bwR = bwBD
arwR = arwBD
lfR = lfBD
shR = swBD
vstallR = vstallBD
swetR = swetBD
AfilasaR = Afilasa

Requisitos.WEWOP = Requisitos.pesoVR / Requisitos.woBD
Requisitos.WEWOP = Format(Requisitos.WEWOP, "0.00")

End Sub



Private Sub Optionbc_Click()
If Tipo.Option1.Value = True Then

tloiterBD = 15
tloiterBD = Format(tloiterBD, "0.00")

        Tipo.Data1.RecordSource = ("Select Avg(PesoF) as pesoCR from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Avg(PesoV) as pesoVR from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Avg(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(teto) as tetomaxBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(hp) as hpBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(climb) as climbBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(sw) as swBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(PesoT) as woBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Bw) as bwBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(lfuse) as lfBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(sht) as shBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(vstall) as vstallBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option1.Value = False
                Tipo.Option1.Value = True
         
End If

If Tipo.Option2.Value = True Then

tloiterBD = 20
tloiterBD = Format(tloiterBD, "0.00")

        Tipo.Data1.RecordSource = ("Select Avg(PesoF) as pesoCR from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Avg(PesoV) as pesoVR from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Avg(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(teto) as tetomaxBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(hp) as hpBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(climb) as climbBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(sw) as swBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(PesoT) as woBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Bw) as bwBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(lfuse) as lfBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(sht) as shBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(vstall) as vstallBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option2.Value = False
                Tipo.Option2.Value = True
End If


If Tipo.Option3.Value = True Then

tloiterBD = 20
tloiterBD = Format(tloiterBD, "0.00")

        Tipo.Data1.RecordSource = ("Select Avg(PesoF) as pesoCR from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Avg(PesoV) as pesoVR from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Avg(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(teto) as tetomaxBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(hp) as hpBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(climb) as climbBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(sw) as swBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(PesoT) as woBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Bw) as bwBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(lfuse) as lfBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(sht) as shBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(vstall) as vstallBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option3.Value = False
                Tipo.Option3.Value = True
End If

If Tipo.Option4.Value = True Then

tloiterBD = 25
tloiterBD = Format(tloiterBD, "0.00")

        Tipo.Data1.RecordSource = ("Select Avg(PesoF) as pesoCR from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Avg(PesoV) as pesoVR from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Avg(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(teto) as tetomaxBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(hp) as hpBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(climb) as climbBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(sw) as swBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(PesoT) as woBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Bw) as bwBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(lfuse) as lfBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(sht) as shBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(vstall) as vstallBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option4.Value = False
                Tipo.Option4.Value = True
End If

If Tipo.Option5.Value = True Then

tloiterBD = 25
tloiterBD = Format(tloiterBD, "0.00")

        Tipo.Data1.RecordSource = ("Select Avg(PesoF) as pesoCR from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Avg(PesoV) as pesoVR from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
 
        Tipo.Data1.RecordSource = ("Select Avg(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(teto) as tetomaxBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(hp) as hpBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(climb) as climbBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(sw) as swBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(PesoT) as woBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(Bw) as bwBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(lfuse) as lfBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(sht) as shBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select avg(vstall) as vstallBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option5.Value = False
                Tipo.Option5.Value = True
End If

' Calculos outros

If Tipomot = "j" Then CBHPcz = 0.8 - (0.00096 * (pesoCR ^ 0.5))
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "tp" Then CBHPcz = 2.5 * hpBD ^ -0.1905
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "p" Then CBHPcz = 0.4
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "j" Then CBHPlt = 1.2 * CBHPcz 'Decolagem 0.49 - (0.0007 * (pesoCR ^ 0.5))
CBHPlt = Format(CBHPlt, "0.00")

If Tipomot = "tp" Then CBHPlt = 1.2 * CBHPcz
CBHPlt = Format(CBHPlt, "0.00")

If Tipomot = "p" Then CBHPlt = 0.5
CBHPlt = Format(CBHPlt, "0.00")


NIcruz = 0.8
NIcruz = Format(NIcruz, "0.00")

NIloiter = 0.7
NIloiter = Format(NIloiter, "0.00")

'Calculos para aerodinamica
clcalc = (2 * woBD) / (0.00238 * ((vcruzBD * 1.4666) ^ 2) * swBD)
clcalc = Format(clcalc, "0.00")

arwBD = bwBD ^ 2 / swBD
arwBD = Format(arwBD, "0.00")

clmaxBD = (2 * woBD) / (0.00238 * ((vstallBD * 1.4666) ^ 2) * swBD)
clmaxBD = Format(clmaxBD, "0.00")

Cdocalc = ((250 * 1.414 * (clcalc ^ 1.5) * (0.00238 ^ 0.5) * (swBD ^ 0.5) * hpBD * 0.65) / woBD ^ 1.5)
If Tipo.Option5.Value = True Then Cdocalc = Format(Cdocalc, "0.00000")
If Tipo.Option5.Value = False Then Cdocalc = Format(Cdocalc, "0.00")


cdimaxBD = (clcalc ^ 2) / (3.1415 * arwBD * 0.85)
cdimaxBD = Format(cdimaxBD, "0.00")

ldcruzBD = clcalc / Cdocalc
ldcruzBD = Format(ldcruzBD, "0.00")

ldloiterBD = ldcruzBD * 0.866
ldloiterBD = Format(ldloiterBD, "0.00")

wsBD = woBD / swBD
wsBD = Format(wsBD, "0.00")

'Calculo da carga  util
CutilR = woBD - pesoVR
CutilR = Format(CutilR, "0.00")

'YYYYYYYYYYYY Lançando Dados YYYYYYYYYYYYYYYY
'performance
daterraR = daterraBD
ddecolR = ddecolBD
climbR = climbBD
VcmaxR = vcruzBD
alcR = alcanceBD
consmotltR = CBHPlt
consmotczR = CBHPcz
NIcruzR = NIcruz
NIloiterR = NIloiter
hpR = hpBD
tetoR = tetomaxBD
tempR = tloiterBD
'aerodinamica
ldcruzR = ldcruzBD
ldloiterR = ldloiterBD
wsR = wsBD
swR = swBD
bwR = bwBD
arwR = arwBD
lfR = lfBD
shR = swBD
vstallR = vstallBD
swetR = swetBD
AfilasaR = Afilasa

Requisitos.WEWOP = Requisitos.pesoVR / Requisitos.woBD
Requisitos.WEWOP = Format(Requisitos.WEWOP, "0.00")

End Sub



Private Sub Optioncc_Click()
If Tipo.Option1.Value = True Then

tloiterBD = 20
tloiterBD = Format(tloiterBD, "0.00")

        Tipo.Data1.RecordSource = ("Select Max(PesoF) as pesoCR from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Max(PesoV) as pesoVR from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Max(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(teto) as tetomaxBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(hp) as hpBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(climb) as climbBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(sw) as swBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Max(PesoT) as woBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Bw) as bwBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(lfuse) as lfBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(sht) as shBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
         
        Tipo.Data1.RecordSource = ("Select max(vstall) as vstallBD from Plan1 where (Categoria = 'TRB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option1.Value = False
                Tipo.Option1.Value = True
End If

If Tipo.Option2.Value = True Then

tloiterBD = 30
tloiterBD = Format(tloiterBD, "0.00")

        Tipo.Data1.RecordSource = ("Select Max(PesoF) as pesoCR from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Max(PesoV) as pesoVR from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Max(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(teto) as tetomaxBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(hp) as hpBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(climb) as climbBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(sw) as swBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(PesoT) as woBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Bw) as bwBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(lfuse) as lfBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(sht) as shBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
         
        Tipo.Data1.RecordSource = ("Select max(vstall) as vstallBD from Plan1 where (Categoria = 'TRP') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option2.Value = False
                Tipo.Option2.Value = True
End If


If Tipo.Option3.Value = True Then

tloiterBD = 30
tloiterBD = Format(tloiterBD, "0.00")

        Tipo.Data1.RecordSource = ("Select Max(PesoF) as pesoCR from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Max(PesoV) as pesoVR from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Max(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(teto) as tetomaxBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(hp) as hpBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(climb) as climbBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(sw) as swBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(PesoT) as woBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Bw) as bwBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(lfuse) as lfBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(sht) as shBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(vstall) as vstallBD from Plan1 where (Categoria = 'TRPB') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option3.Value = False
                Tipo.Option3.Value = True
         
End If

If Tipo.Option4.Value = True Then

tloiterBD = 30
tloiterBD = Format(tloiterBD, "0.00")

        Tipo.Data1.RecordSource = ("Select Max(PesoF) as pesoCR from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Max(PesoV) as pesoVR from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Max(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(teto) as tetomaxBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(hp) as hpBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(climb) as climbBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(sw) as swBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(PesoT) as woBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Bw) as bwBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(lfuse) as lfBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(sht) as shBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(vstall) as vstallBD from Plan1 where (Categoria = 'TRBA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option4.Value = False
                Tipo.Option4.Value = True
         
End If

If Tipo.Option5.Value = True Then

tloiterBD = 30
tloiterBD = Format(tloiterBD, "0.00")

        Tipo.Data1.RecordSource = ("Select Max(PesoF) as pesoCR from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoCR = Tipo.Data1.Recordset("pesoCR")
        pesoCR = Format(pesoCR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Max(PesoV) as pesoVR from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        pesoVR = Tipo.Data1.Recordset("pesoVR")
        pesoVR = Format(pesoVR, "0.00")
        
        Tipo.Data1.RecordSource = ("Select Max(Alcance) as alcanceBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        alcanceBD = Tipo.Data1.Recordset("alcanceBD")
        alcanceBD = Format(alcanceBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(teto) as tetomaxBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        tetomaxBD = Tipo.Data1.Recordset("tetomaxBD")
        tetomaxBD = Format(tetomaxBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(hp) as hpBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        hpBD = Tipo.Data1.Recordset("hpBD")
        hpBD = Format(hpBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(vcruz) as vcruzBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vcruzBD = Tipo.Data1.Recordset("vcruzBD")
        vcruzBD = Format(vcruzBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(climb) as climbBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        climbBD = Tipo.Data1.Recordset("climbBD")
        climbBD = Format(climbBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Ddecol) as DdecolBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        ddecolBD = Tipo.Data1.Recordset("ddecolBD")
        ddecolBD = Format(ddecolBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Daterra) as DaterraBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        daterraBD = Tipo.Data1.Recordset("daterraBD")
        daterraBD = Format(daterraBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(sw) as swBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        swBD = Tipo.Data1.Recordset("swBD")
        swBD = Format(swBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(PesoT) as woBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        woBD = Tipo.Data1.Recordset("woBD")
        woBD = Format(woBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(Bw) as bwBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        bwBD = Tipo.Data1.Recordset("bwBD")
        bwBD = Format(bwBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(lfuse) as lfBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        lfBD = Tipo.Data1.Recordset("lfBD")
        lfBD = Format(lfBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(sht) as shBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        shBD = Tipo.Data1.Recordset("shBD")
        shBD = Format(shBD, "0.00")
        
        Tipo.Data1.RecordSource = ("Select max(vstall) as vstallBD from Plan1 where (Categoria = 'TRA') and (len(Trim(Origem)) > 0)")
        Tipo.Data1.Refresh
        If Tipo.Data1.Recordset.RecordCount < 1 Then Exit Sub
        vstallBD = Tipo.Data1.Recordset("vstallBD")
        vstallBD = Format(vstallBD, "0.00")
        
                Tipo.Option5.Value = False
                Tipo.Option5.Value = True
         
 
End If

' Calculos outros

If Tipomot = "j" Then CBHPcz = 0.8 - (0.00096 * (pesoCR ^ 0.5))
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "tp" Then CBHPcz = 2.5 * hpBD ^ -0.1905
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "p" Then CBHPcz = 0.4
CBHPcz = Format(CBHPcz, "0.00")

If Tipomot = "j" Then CBHPlt = 1.2 * CBHPcz 'Decolagem 0.49 - (0.0007 * (pesoCR ^ 0.5))
CBHPlt = Format(CBHPlt, "0.00")

If Tipomot = "tp" Then CBHPlt = 1.2 * CBHPcz
CBHPlt = Format(CBHPlt, "0.00")

If Tipomot = "p" Then CBHPlt = 0.5
CBHPlt = Format(CBHPlt, "0.00")


NIcruz = 0.8
NIcruz = Format(NIcruz, "0.00")

NIloiter = 0.7
NIloiter = Format(NIloiter, "0.00")

'Calculos para aerodinamica
clcalc = (2 * woBD) / (0.00238 * ((vcruzBD * 1.4666) ^ 2) * swBD)
clcalc = Format(clcalc, "0.00")

arwBD = bwBD ^ 2 / swBD
arwBD = Format(arwBD, "0.00")

clmaxBD = (2 * woBD) / (0.00238 * ((vstallBD * 1.4666) ^ 2) * swBD)
clmaxBD = Format(clmaxBD, "0.00")

Cdocalc = ((250 * 1.414 * (clcalc ^ 1.5) * (0.00238 ^ 0.5) * (swBD ^ 0.5) * hpBD * 0.65) / woBD ^ 1.5)
Cdocalc = Format(Cdocalc, "0.000")

cdimaxBD = (clcalc ^ 2) / (3.1415 * arwBD * 0.85)
cdimaxBD = Format(cdimaxBD, "0.00")

ldcruzBD = clcalc / Cdocalc
ldcruzBD = Format(ldcruzBD, "0.00")

ldloiterBD = ldcruzBD * 0.866
ldloiterBD = Format(ldloiterBD, "0.00")

wsBD = woBD / swBD
wsBD = Format(wsBD, "0.00")

'Calculo da carga  util
CutilR = woBD - pesoVR
CutilR = Format(CutilR, "0.00")

'YYYYYYYYYYYY Lançando Dados YYYYYYYYYYYYYYYY
'performance
daterraR = daterraBD
ddecolR = ddecolBD
climbR = climbBD
VcmaxR = vcruzBD
alcR = alcanceBD
consmotltR = CBHPlt
consmotczR = CBHPcz
NIcruzR = NIcruz
NIloiterR = NIloiter
hpR = hpBD
tetoR = tetomaxBD
tempR = tloiterBD
'aerodinamica
ldcruzR = ldcruzBD
ldloiterR = ldloiterBD
wsR = wsBD
swR = swBD
bwR = bwBD
arwR = arwBD
lfR = lfBD
shR = swBD
vstallR = vstallBD
swetR = swetBD
AfilasaR = Afilasa

Requisitos.WEWOP = Requisitos.pesoVR / Requisitos.woBD
Requisitos.WEWOP = Format(Requisitos.WEWOP, "0.00")

End Sub

Private Sub tempR_GotFocus()
tempR.SelStart = 0
tempR.SelLength = Len(tempR)
End Sub



Private Sub tetoR_GotFocus()
tetoR.SelStart = 0
tetoR.SelLength = Len(tetoR)
End Sub

Private Sub VmaxR_GotFocus()
VmaxR.SelStart = 0
VmaxR.SelLength = Len(VmaxR)
End Sub

Private Sub wmotorR_GotFocus()
wmotorR.SelStart = 0
wmotorR.SelLength = Len(wmotorR)
End Sub

Private Sub Tubalum_Click()
'Requisitos.AltVariav = Checked
'Requisitos.HScroll4 = 10
'Requisitos.HScroll5 = 10
End Sub

Private Sub tubo_Click()
'Requisitos.AltVariav = Checked
'Requisitos.HScroll4 = 10
'Requisitos.HScroll5 = 10
End Sub
