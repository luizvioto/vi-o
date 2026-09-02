VERSION 5.00
Begin VB.Form Form27 
   Caption         =   "Trimagem"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form27.frx":0000
   LinkTopic       =   "Form27"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Trimagem"
      Height          =   1290
      Left            =   7800
      TabIndex        =   116
      Top             =   6840
      Width           =   3975
      Begin VB.PictureBox Picture6 
         Height          =   225
         Left            =   1560
         ScaleHeight     =   165
         ScaleWidth      =   1335
         TabIndex        =   175
         Top             =   600
         Width           =   1400
         Begin VB.HScrollBar HScroll2 
            Height          =   185
            Left            =   0
            Max             =   500
            TabIndex        =   176
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.TextBox altura 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   139
         Text            =   "0"
         Top             =   960
         Width           =   615
      End
      Begin VB.PictureBox Picture4 
         Height          =   225
         Left            =   1560
         ScaleHeight     =   165
         ScaleWidth      =   1335
         TabIndex        =   137
         Top             =   960
         Width           =   1400
         Begin VB.HScrollBar HScroll1 
            Height          =   185
            Left            =   0
            Max             =   2000
            Min             =   1
            TabIndex        =   138
            Top             =   0
            Value           =   10
            Width           =   1335
         End
      End
      Begin VB.TextBox testecl 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   119
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Velocidade máxima"
         Height          =   255
         Left            =   1920
         TabIndex        =   118
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Subida"
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CL opcional :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   174
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuste do gráfico :"
         Height          =   195
         Left            =   120
         TabIndex        =   140
         Top             =   960
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   5055
      Left            =   7680
      ScaleHeight     =   4995
      ScaleWidth      =   4035
      TabIndex        =   141
      Top             =   1800
      Width           =   4095
      Begin VB.Frame Frame1 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5055
         Left            =   0
         TabIndex        =   142
         Top             =   0
         Width           =   4095
         Begin VB.TextBox Ew 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   171
            Text            =   "0"
            Top             =   4440
            Width           =   615
         End
         Begin VB.TextBox clafaw 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   154
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox xcg 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   153
            Text            =   "0"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox xacw 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   152
            Text            =   "0"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox cmfuse 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   151
            Text            =   "0"
            Top             =   2280
            Width           =   615
         End
         Begin VB.TextBox nih 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   150
            Text            =   "0"
            Top             =   2640
            Width           =   615
         End
         Begin VB.TextBox shsw 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   149
            Text            =   "0"
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox xachcg 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   148
            Text            =   "0"
            Top             =   3360
            Width           =   615
         End
         Begin VB.TextBox clalfah 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   147
            Text            =   "0"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox cecp 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   146
            Text            =   "0"
            Top             =   3720
            Width           =   615
         End
         Begin VB.TextBox drodf 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   145
            Text            =   "0"
            Top             =   4080
            Width           =   615
         End
         Begin VB.TextBox dedalfa 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   144
            Text            =   "0"
            Top             =   4440
            Width           =   615
         End
         Begin VB.TextBox PN27 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   143
            Text            =   "0"
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Espes. da onda:                m"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   1680
            TabIndex        =   170
            Top             =   4440
            Width           =   1995
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Corda do profundor / corda do Eh :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   169
            Top             =   3720
            Width           =   2490
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Slope da asa :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   168
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Distancia do (CG - LR) / Cmw :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   167
            Top             =   1560
            Width           =   2190
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Distancia do (Caw - LR) / Cmw :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   166
            Top             =   1920
            Width           =   2280
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Coeficiente de momento - fuselagem :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   165
            Top             =   2280
            Width           =   2655
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Coeficiente NI (eficiência de Eh) :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   164
            Top             =   2640
            Width           =   2370
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Área da asa / Área do Estab. horiz. :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   163
            Top             =   3000
            Width           =   2580
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Distancia do CG - CA Ehoriz./ Cmw :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   162
            Top             =   3360
            Width           =   2580
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Slope da empenagem :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   161
            Top             =   840
            Width           =   1635
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "dr / dj  ..........................................  :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   160
            Top             =   4080
            Width           =   2655
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "de / da :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   120
            TabIndex        =   159
            Top             =   4440
            Width           =   630
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/rad"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   3555
            TabIndex        =   158
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/rad"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   3555
            TabIndex        =   157
            Top             =   840
            Width           =   300
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Distancia do (PN - LR) / Cmw :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   156
            Top             =   1200
            Width           =   2190
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Parâmetros calculados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   240
            Left            =   960
            TabIndex        =   155
            Top             =   120
            Width           =   2085
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   455
      Left            =   120
      ScaleHeight     =   390
      ScaleWidth      =   7515
      TabIndex        =   111
      Top             =   7680
      Width           =   7575
      Begin VB.CommandButton Command3 
         Caption         =   "Sair"
         Height          =   385
         Left            =   5760
         MouseIcon       =   "Form27.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   115
         Top             =   0
         Width           =   1755
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Limpar"
         Height          =   385
         Left            =   3840
         MouseIcon       =   "Form27.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   114
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Trimagem"
         Height          =   385
         Left            =   1920
         MouseIcon       =   "Form27.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   113
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cmcg x CL"
         Height          =   385
         Left            =   0
         MouseIcon       =   "Form27.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   112
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      Height          =   5835
      Left            =   120
      ScaleHeight     =   5775
      ScaleWidth      =   7515
      TabIndex        =   98
      Top             =   1800
      Width           =   7575
      Begin VB.Line Line13 
         BorderColor     =   &H00C0C0C0&
         X1              =   6720
         X2              =   6720
         Y1              =   0
         Y2              =   5760
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Deflexão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   6720
         TabIndex        =   173
         Top             =   0
         Width           =   855
      End
      Begin VB.Label defgrau 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6000
         TabIndex        =   136
         Top             =   2280
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label cmcgok1 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   840
         TabIndex        =   135
         Top             =   2280
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label kde33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   6960
         TabIndex        =   134
         Top             =   3360
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label kde22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   180
         Left            =   6960
         TabIndex        =   133
         Top             =   3120
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label kde11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   6960
         TabIndex        =   132
         Top             =   2880
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label kde4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   6960
         TabIndex        =   131
         Top             =   2520
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label kde3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   6960
         TabIndex        =   130
         Top             =   2280
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label kde2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   6960
         TabIndex        =   129
         Top             =   2040
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label kde1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   6960
         TabIndex        =   128
         Top             =   1800
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label cmg33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   120
         TabIndex        =   127
         Top             =   4080
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label cmg22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   180
         Left            =   120
         TabIndex        =   126
         Top             =   3600
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label cmg11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   120
         TabIndex        =   125
         Top             =   3120
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label cmg0 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   124
         Top             =   2400
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label cmg3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   120
         TabIndex        =   123
         Top             =   960
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label cmg2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   120
         TabIndex        =   122
         Top             =   1440
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label cmg1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   120
         TabIndex        =   121
         Top             =   1920
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6600
         TabIndex        =   120
         Top             =   120
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "CL"
         ForeColor       =   &H00C0C0FF&
         Height          =   240
         Left            =   7200
         TabIndex        =   110
         Top             =   2640
         Width           =   240
      End
      Begin VB.Shape clponto 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1380
         Shape           =   3  'Circle
         Top             =   2640
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0FF&
         X1              =   -120
         X2              =   7560
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "Cmcg"
         ForeColor       =   &H00C0C0FF&
         Height          =   240
         Left            =   1200
         TabIndex        =   109
         Top             =   240
         Width           =   525
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0FF&
         X1              =   1440
         X2              =   1440
         Y1              =   0
         Y2              =   5760
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   1935
         X2              =   1935
         Y1              =   240
         Y2              =   5520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   2445
         X2              =   2445
         Y1              =   240
         Y2              =   5520
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   2940
         X2              =   2940
         Y1              =   240
         Y2              =   5520
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   3435
         X2              =   3435
         Y1              =   240
         Y2              =   5520
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   3960
         X2              =   3945
         Y1              =   240
         Y2              =   5520
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   4440
         X2              =   4440
         Y1              =   240
         Y2              =   5520
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   4935
         X2              =   4935
         Y1              =   240
         Y2              =   5520
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   5445
         X2              =   5445
         Y1              =   240
         Y2              =   5520
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   5940
         X2              =   5940
         Y1              =   240
         Y2              =   5520
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   6435
         X2              =   6435
         Y1              =   240
         Y2              =   5520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   180
         Left            =   6240
         TabIndex        =   108
         Top             =   5520
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   180
         Left            =   5760
         TabIndex        =   107
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   180
         Left            =   5280
         TabIndex        =   106
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   180
         Left            =   4800
         TabIndex        =   105
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   180
         Left            =   4320
         TabIndex        =   104
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   180
         Left            =   3840
         TabIndex        =   103
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   180
         Left            =   3240
         TabIndex        =   102
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   180
         Left            =   2760
         TabIndex        =   101
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   180
         Left            =   2280
         TabIndex        =   100
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   180
         Left            =   1800
         TabIndex        =   99
         Top             =   5520
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400000&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   0
         TabIndex        =   172
         Top             =   0
         Width           =   11655
      End
      Begin VB.Label dd1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   6600
         TabIndex        =   97
         Top             =   240
         Width           =   75
      End
      Begin VB.Label aa1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   5880
         TabIndex        =   96
         Top             =   240
         Width           =   75
      End
      Begin VB.Label aa2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   5880
         TabIndex        =   95
         Top             =   480
         Width           =   75
      End
      Begin VB.Label dd2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   6600
         TabIndex        =   94
         Top             =   480
         Width           =   75
      End
      Begin VB.Label aa3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   5880
         TabIndex        =   93
         Top             =   720
         Width           =   75
      End
      Begin VB.Label dd3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   6600
         TabIndex        =   92
         Top             =   720
         Width           =   75
      End
      Begin VB.Label aa4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   5880
         TabIndex        =   91
         Top             =   960
         Width           =   75
      End
      Begin VB.Label dd4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   6600
         TabIndex        =   90
         Top             =   960
         Width           =   75
      End
      Begin VB.Label aa5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   5880
         TabIndex        =   89
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label dd5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   6600
         TabIndex        =   88
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label aa6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   5880
         TabIndex        =   87
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label dd6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   6600
         TabIndex        =   86
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label dd66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   8040
         TabIndex        =   85
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label aa66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   7320
         TabIndex        =   84
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label dd44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   8040
         TabIndex        =   83
         Top             =   960
         Width           =   75
      End
      Begin VB.Label aa55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   7320
         TabIndex        =   82
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label dd33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   8040
         TabIndex        =   81
         Top             =   720
         Width           =   75
      End
      Begin VB.Label aa44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   7320
         TabIndex        =   80
         Top             =   960
         Width           =   75
      End
      Begin VB.Label dd55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   8040
         TabIndex        =   79
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label aa33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   7320
         TabIndex        =   78
         Top             =   720
         Width           =   75
      End
      Begin VB.Label dd22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   8040
         TabIndex        =   77
         Top             =   480
         Width           =   75
      End
      Begin VB.Label aa22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   7320
         TabIndex        =   76
         Top             =   480
         Width           =   75
      End
      Begin VB.Label aa11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   7320
         TabIndex        =   75
         Top             =   240
         Width           =   75
      End
      Begin VB.Label dd11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   180
         Left            =   8040
         TabIndex        =   74
         Top             =   240
         Width           =   75
      End
      Begin VB.Label dd111 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   9480
         TabIndex        =   73
         Top             =   240
         Width           =   75
      End
      Begin VB.Label aa111 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   8760
         TabIndex        =   72
         Top             =   240
         Width           =   75
      End
      Begin VB.Label aa222 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   8760
         TabIndex        =   71
         Top             =   480
         Width           =   75
      End
      Begin VB.Label dd222 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   9480
         TabIndex        =   70
         Top             =   480
         Width           =   75
      End
      Begin VB.Label aa333 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   8760
         TabIndex        =   69
         Top             =   720
         Width           =   75
      End
      Begin VB.Label dd555 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   9480
         TabIndex        =   68
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label aa444 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   8760
         TabIndex        =   67
         Top             =   960
         Width           =   75
      End
      Begin VB.Label dd333 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   9480
         TabIndex        =   66
         Top             =   720
         Width           =   75
      End
      Begin VB.Label aa555 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   8760
         TabIndex        =   65
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label dd444 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   9480
         TabIndex        =   64
         Top             =   960
         Width           =   75
      End
      Begin VB.Label aa666 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   8760
         TabIndex        =   63
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label dd666 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   9480
         TabIndex        =   62
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label dd6666 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10920
         TabIndex        =   61
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label aa6666 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10200
         TabIndex        =   60
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label dd4444 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10920
         TabIndex        =   59
         Top             =   960
         Width           =   75
      End
      Begin VB.Label aa5555 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10200
         TabIndex        =   58
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label dd3333 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10920
         TabIndex        =   57
         Top             =   720
         Width           =   75
      End
      Begin VB.Label aa4444 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10200
         TabIndex        =   56
         Top             =   960
         Width           =   75
      End
      Begin VB.Label dd5555 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10920
         TabIndex        =   55
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label aa3333 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10200
         TabIndex        =   54
         Top             =   720
         Width           =   75
      End
      Begin VB.Label dd2222 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10920
         TabIndex        =   53
         Top             =   480
         Width           =   75
      End
      Begin VB.Label aa2222 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10200
         TabIndex        =   52
         Top             =   480
         Width           =   75
      End
      Begin VB.Label aa1111 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10200
         TabIndex        =   51
         Top             =   240
         Width           =   75
      End
      Begin VB.Label dd1111 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   180
         Left            =   10920
         TabIndex        =   50
         Top             =   240
         Width           =   75
      End
      Begin VB.Label d1111 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   5160
         TabIndex        =   49
         Top             =   240
         Width           =   75
      End
      Begin VB.Label a1111 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   4440
         TabIndex        =   48
         Top             =   240
         Width           =   75
      End
      Begin VB.Label a2222 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   4440
         TabIndex        =   47
         Top             =   480
         Width           =   75
      End
      Begin VB.Label d2222 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   5160
         TabIndex        =   46
         Top             =   480
         Width           =   75
      End
      Begin VB.Label a3333 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   4440
         TabIndex        =   45
         Top             =   720
         Width           =   75
      End
      Begin VB.Label d5555 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   5160
         TabIndex        =   44
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label a4444 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   4440
         TabIndex        =   43
         Top             =   960
         Width           =   75
      End
      Begin VB.Label d3333 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   5160
         TabIndex        =   42
         Top             =   720
         Width           =   75
      End
      Begin VB.Label a5555 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   4440
         TabIndex        =   41
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label d4444 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   5160
         TabIndex        =   40
         Top             =   960
         Width           =   75
      End
      Begin VB.Label a6666 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   4440
         TabIndex        =   39
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label d6666 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   5160
         TabIndex        =   38
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label d666 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3720
         TabIndex        =   37
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label a666 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3000
         TabIndex        =   36
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label d444 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3720
         TabIndex        =   35
         Top             =   960
         Width           =   75
      End
      Begin VB.Label a555 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3000
         TabIndex        =   34
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label d333 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3720
         TabIndex        =   33
         Top             =   720
         Width           =   75
      End
      Begin VB.Label a444 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3000
         TabIndex        =   32
         Top             =   960
         Width           =   75
      End
      Begin VB.Label d555 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3720
         TabIndex        =   31
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label a333 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3000
         TabIndex        =   30
         Top             =   720
         Width           =   75
      End
      Begin VB.Label d222 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3720
         TabIndex        =   29
         Top             =   480
         Width           =   75
      End
      Begin VB.Label a222 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3000
         TabIndex        =   28
         Top             =   480
         Width           =   75
      End
      Begin VB.Label a111 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   75
      End
      Begin VB.Label d111 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   180
         Left            =   3720
         TabIndex        =   26
         Top             =   240
         Width           =   75
      End
      Begin VB.Label d11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   75
      End
      Begin VB.Label a11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   75
      End
      Begin VB.Label a22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   1560
         TabIndex        =   23
         Top             =   480
         Width           =   75
      End
      Begin VB.Label d22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   2280
         TabIndex        =   22
         Top             =   480
         Width           =   75
      End
      Begin VB.Label a33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   1560
         TabIndex        =   21
         Top             =   720
         Width           =   75
      End
      Begin VB.Label d55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   2280
         TabIndex        =   20
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label a44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   1560
         TabIndex        =   19
         Top             =   960
         Width           =   75
      End
      Begin VB.Label d33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   2280
         TabIndex        =   18
         Top             =   720
         Width           =   75
      End
      Begin VB.Label a55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   1560
         TabIndex        =   17
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label d44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   2280
         TabIndex        =   16
         Top             =   960
         Width           =   75
      End
      Begin VB.Label a66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   1560
         TabIndex        =   15
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label d66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   2280
         TabIndex        =   14
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label d6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   840
         TabIndex        =   13
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label a6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label d5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   840
         TabIndex        =   11
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label a5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label d4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   840
         TabIndex        =   9
         Top             =   960
         Width           =   75
      End
      Begin VB.Label a4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   75
      End
      Begin VB.Label d3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   75
      End
      Begin VB.Label a3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   75
      End
      Begin VB.Label d2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   75
      End
      Begin VB.Label a2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   75
      End
      Begin VB.Label a1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   75
      End
      Begin VB.Label d1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   75
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cmcg  x  CL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   5160
         TabIndex        =   1
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.Menu fileir 
      Caption         =   "&Ir para ..."
      Begin VB.Menu fileddesenho 
         Caption         =   "Página de desenho"
      End
      Begin VB.Menu filepesobalanc 
         Caption         =   "Peso && balanceamento"
      End
      Begin VB.Menu fileestrut 
         Caption         =   "Estruturas"
      End
      Begin VB.Menu fileestab 
         Caption         =   "Estabilidade estática"
      End
      Begin VB.Menu fileperform 
         Caption         =   "Performance"
      End
      Begin VB.Menu fileaerodin 
         Caption         =   "Aerodinâmica"
      End
      Begin VB.Menu filedimensiona 
         Caption         =   "Dimensionamento"
      End
      Begin VB.Menu filepaginicio 
         Caption         =   "Página inicial"
      End
   End
   Begin VB.Menu filefer27 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "&Calculadora && Conversor"
      End
   End
   Begin VB.Menu fileimprimi27 
      Caption         =   "&Imprimir"
   End
   Begin VB.Menu sai27 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Frame3.Visible = False
Form27.Option4.Value = False
Form27.Option3.Value = False

'CALCULO DOS PARAMETROS

Masa = Form4.Data1.Recordset("Cmo")

fria = Masa + Abs(cmfuse)
fria = Format(fria, "0.000")

Fatoralfa = (clafaw * (xcg - xacw) + fria)
Fatoralfa = Format(Fatoralfa, "0.000")

fatclh = 1 - (dedalfa)
fatclh = Format(fatclh, "0.000")

Fatorht = nih * shsw * xachcg
Fatorht = Format(Fatorht, "0.000")

DeltaOLH = (-nih / clalfah) * (5.3) * (cecp)
DeltaOLH = Format(DeltaOLH, "0.000")

Fatorcmcg = Fatorht * clalfah * dedalfa
Fatorcmcg = Format(Fatorcmcg, "0.000")

cmcgvero = Fatoralfa - (Fatorht * clalfah * fatclh)
cmcgvero = Format(cmcgvero, "0.000")

ALFAdef = Format(cmcgvero, "0.000")
ALFAdef = Format(ALFAdef, "0.000")

DEdef = clalfah * DeltaOLH * Fatorht
DEdef = Format(DEdef, "0.000")

'FIM


For kalfa = 0 To 10 Step 1
For kde = -6 To 6 Step 1

alfa = kalfa
DE = kde

'kcmcg = (-0.195 * (ALFA * 3.141516 / 180)) - (0.43 * (DE * 3.141516 / 180))
'Calculado pelo programa :
kcmcg = (ALFAdef * (alfa * 3.141516 / 180)) + (DEdef * (DE * 3.141516 / 180))


cmcg = kcmcg
cmcg = Format(cmcg, "0.000")

cltotal = 5.05 * (alfa * 3.141516 / 180) + 0.168 * (DE * 3.141516 / 180)
cltotal = Format(cltotal, "0.000")

'PARA alfa 0 a 10 e de 0
If alfa = 0 And DE = 0 Then a1 = cmcg
If alfa = 0 And DE = 0 Then d1 = cltotal

If alfa = 2 And DE = 0 Then a2 = cmcg
If alfa = 2 And DE = 0 Then d2 = cltotal

If alfa = 4 And DE = 0 Then a3 = cmcg
If alfa = 4 And DE = 0 Then d3 = cltotal

If alfa = 6 And DE = 0 Then a4 = cmcg
If alfa = 6 And DE = 0 Then d4 = cltotal

If alfa = 8 And DE = 0 Then a5 = cmcg
If alfa = 8 And DE = 0 Then d5 = cltotal

If alfa = 10 And DE = 0 Then a6 = cmcg
If alfa = 10 And DE = 0 Then d6 = cltotal

'PARA alfa 0 a 10 e de -2
If alfa = 0 And DE = -2 Then a11 = cmcg
If alfa = 0 And DE = -2 Then d11 = cltotal

If alfa = 2 And DE = -2 Then a22 = cmcg
If alfa = 2 And DE = -2 Then d22 = cltotal

If alfa = 4 And DE = -2 Then a33 = cmcg
If alfa = 4 And DE = -2 Then d33 = cltotal

If alfa = 6 And DE = -2 Then a44 = cmcg
If alfa = 6 And DE = -2 Then d44 = cltotal

If alfa = 8 And DE = -2 Then a55 = cmcg
If alfa = 8 And DE = -2 Then d55 = cltotal

If alfa = 10 And DE = -2 Then a66 = cmcg
If alfa = 10 And DE = -2 Then d66 = cltotal

'PARA alfa 0 a 10 e de -4
If alfa = 0 And DE = -4 Then a111 = cmcg
If alfa = 0 And DE = -4 Then d111 = cltotal

If alfa = 2 And DE = -4 Then a222 = cmcg
If alfa = 2 And DE = -4 Then d222 = cltotal

If alfa = 4 And DE = -4 Then a333 = cmcg
If alfa = 4 And DE = -4 Then d333 = cltotal

If alfa = 6 And DE = -4 Then a444 = cmcg
If alfa = 6 And DE = -4 Then d444 = cltotal

If alfa = 8 And DE = -4 Then a555 = cmcg
If alfa = 8 And DE = -4 Then d555 = cltotal

If alfa = 10 And DE = -4 Then a666 = cmcg
If alfa = 10 And DE = -4 Then d666 = cltotal

'PARA alfa 0 a 10 e de -6
If alfa = 0 And DE = -6 Then a1111 = cmcg
If alfa = 0 And DE = -4 Then d1111 = cltotal

If alfa = 2 And DE = -6 Then a2222 = cmcg
If alfa = 2 And DE = -6 Then d2222 = cltotal

If alfa = 4 And DE = -6 Then a3333 = cmcg
If alfa = 4 And DE = -6 Then d3333 = cltotal

If alfa = 6 And DE = -6 Then a4444 = cmcg
If alfa = 6 And DE = -6 Then d4444 = cltotal

If alfa = 8 And DE = -6 Then a5555 = cmcg
If alfa = 8 And DE = -6 Then d5555 = cltotal

If alfa = 10 And DE = -6 Then a6666 = cmcg
If alfa = 10 And DE = -6 Then d6666 = cltotal


'.......PARA "de" POSITIVOS .......................

'PARA alfa 0 a 10 e de 0
If alfa = 0 And DE = 0 Then aa1 = cmcg
If alfa = 0 And DE = 0 Then dd1 = cltotal

If alfa = 2 And DE = 0 Then aa2 = cmcg
If alfa = 2 And DE = 0 Then dd2 = cltotal

If alfa = 4 And DE = 0 Then aa3 = cmcg
If alfa = 4 And DE = 0 Then dd3 = cltotal

If alfa = 6 And DE = 0 Then aa4 = cmcg
If alfa = 6 And DE = 0 Then dd4 = cltotal

If alfa = 8 And DE = 0 Then aa5 = cmcg
If alfa = 8 And DE = 0 Then dd5 = cltotal

If alfa = 10 And DE = 0 Then aa6 = cmcg
If alfa = 10 And DE = 0 Then dd6 = cltotal

'PARA alfa 0 a 10 e de 2
If alfa = 0 And DE = 2 Then aa11 = cmcg
If alfa = 0 And DE = 2 Then dd11 = cltotal

If alfa = 2 And DE = 2 Then aa22 = cmcg
If alfa = 2 And DE = 2 Then dd22 = cltotal

If alfa = 4 And DE = 2 Then aa33 = cmcg
If alfa = 4 And DE = 2 Then dd33 = cltotal

If alfa = 6 And DE = 2 Then aa44 = cmcg
If alfa = 6 And DE = 2 Then dd44 = cltotal

If alfa = 8 And DE = 2 Then aa55 = cmcg
If alfa = 8 And DE = 2 Then dd55 = cltotal

If alfa = 10 And DE = 2 Then aa66 = cmcg
If alfa = 10 And DE = 2 Then dd66 = cltotal

'PARA alfa 0 a 10 e de -4
If alfa = 0 And DE = 4 Then aa111 = cmcg
If alfa = 0 And DE = 4 Then dd111 = cltotal

If alfa = 2 And DE = 4 Then aa222 = cmcg
If alfa = 2 And DE = 4 Then dd222 = cltotal

If alfa = 4 And DE = 4 Then aa333 = cmcg
If alfa = 4 And DE = 4 Then dd333 = cltotal

If alfa = 6 And DE = 4 Then aa444 = cmcg
If alfa = 6 And DE = 4 Then dd444 = cltotal

If alfa = 8 And DE = 4 Then aa555 = cmcg
If alfa = 8 And DE = 4 Then dd555 = cltotal

If alfa = 10 And DE = 4 Then aa666 = cmcg
If alfa = 10 And DE = 4 Then dd666 = cltotal

'PARA alfa 0 a 10 e de -6
If alfa = 0 And DE = 6 Then aa1111 = cmcg
If alfa = 0 And DE = 6 Then dd1111 = cltotal

If alfa = 2 And DE = 6 Then aa2222 = cmcg
If alfa = 2 And DE = 6 Then dd2222 = cltotal

If alfa = 4 And DE = 6 Then aa3333 = cmcg
If alfa = 4 And DE = 6 Then dd3333 = cltotal

If alfa = 6 And DE = 6 Then aa4444 = cmcg
If alfa = 6 And DE = 6 Then dd4444 = cltotal

If alfa = 8 And DE = 6 Then aa5555 = cmcg
If alfa = 8 And DE = 6 Then dd5555 = cltotal

If alfa = 10 And DE = 6 Then aa6666 = cmcg
If alfa = 10 And DE = 6 Then dd6666 = cltotal



Next
Next

'Desenho das linhas negativas
Picture1.DrawStyle = 2

'altura = 5000

Picture1.Line (1440 + (d1), 2760 - (altura * a1))-(1440 + (5000 * d6), 2760 - (altura * a6)), &HFFFFFF
kde1 = "0º"
kde1.Top = (2760 - (altura * a6)) - kde1.Height / 2
kde1.Visible = True

Picture1.Line (1440 + (d11), 2760 - (altura * a11))-(1440 + (5000 * d66), 2760 - (altura * a66)), &HFFFF00
kde2 = "-2º"
kde2.Top = 2760 - (altura * a66) - kde2.Height / 2
kde2.Visible = True

Picture1.Line (1440 + (d111), 2760 - (altura * a111))-(1440 + (5000 * d666), 2760 - (altura * a666)), &HFF00FF
kde3 = "-4º"
kde3.Top = 2760 - (altura * a666) - kde3.Height / 2
kde3.Visible = True

Picture1.Line (1440 + (d1111), 2760 - (altura * a1111))-(1440 + (5000 * d6666), 2760 - (altura * a6666)), &HFF00&
kde4 = "-6º"
kde4.Top = 2760 - (altura * a6666) - kde4.Height / 2
kde4.Visible = True
'Desenho das linhas positivas

Picture1.Line (1440 + (dd1), 2760 - (altura * aa1))-(1440 + (5000 * dd6), 2760 - (altura * aa6)), &HFFFFFF
'ZERO

Picture1.Line (1440 + (dd11), 2760 - (altura * aa11))-(1440 + (5000 * dd66), 2760 - (altura * aa66)), &HFF8080
kde11 = "2º"
kde11.Top = 2760 - (altura * aa66) - kde11.Height / 2
kde11.Visible = True

Picture1.Line (1440 + (dd111), 2760 - (altura * aa111))-(1440 + (5000 * dd666), 2760 - (altura * aa666)), &H8080FF
kde22 = "4º"
kde22.Top = 2760 - (altura * aa666) - kde22.Height / 2
kde22.Visible = True

Picture1.Line (1440 + (dd1111), 2760 - (altura * aa1111))-(1440 + (5000 * dd6666), 2760 - (altura * aa6666)), &HC0FFC0
kde33 = "6º"
kde33.Top = 2760 - (altura * aa6666) - kde22.Height / 2
kde33.Visible = True

'CMGs
'NEGATIVO
cmg1 = a11
cmg1.Top = (2760 - (altura * a11)) - cmg1.Height / 2
cmg1.Visible = True

cmg2 = a111
cmg2.Top = (2760 - (altura * a111)) - cmg2.Height / 2
cmg2.Visible = True

cmg3 = a1111
cmg3.Top = (2760 - (altura * cmg3)) - cmg3.Height / 2
cmg3.Visible = True

cmg0 = a1
cmg0.Top = (2760 - (altura * cmg0)) - cmg0.Height / 2
cmg0.Visible = True

'CMGs
'POSITIVO
cmg11 = aa11
cmg11.Top = (2760 - (altura * cmg11)) - cmg11.Height / 2
cmg11.Visible = True

cmg22 = aa111
cmg22.Top = (2760 - (altura * cmg22)) - cmg22.Height / 2
cmg22.Visible = True

cmg33 = aa1111
cmg33.Top = (2760 - (altura * cmg33)) - cmg33.Height / 2
cmg33.Visible = True

End Sub


Private Sub Command2_Click()
If testecl = 0 Then
MsgBox " Escolha o Regime ou CL", vbInformation
Exit Sub
End If


clponto.Visible = True
vv = (5000 * testecl)
For kcl = 0 To vv Step 0.1
clponto.Left = 1375 + kcl
Next

'Traçado da linha do Cl testado
If a6 = "-" And a1 = "-" Then
    MsgBox " Cmcg x CL deve ser calculado.", vbInformation
    Call Command4_Click
    Exit Sub
End If

nicola = a6 - a1


If nicola = 0 Then nicola = 1E-100


ca = -1 * ((d6 - d1) / (nicola))
ca = Format(ca, "0.000")

cmcgok1 = testecl / ca
cmcgok1 = Format(cmcgok1, "0.000")
cmcgok1.Visible = True

cmcgok = testecl / ca
cmcgok = Format(cmcgok, "0.0000")
Picture1.DrawStyle = 0

ky = cmcgok / testecl
ky = Format(ky, "0.000")

yf = (d6 - testecl) * ky
yf = Format(yf, "0.000")

Picture1.Line (1440, 2760 - (altura * cmcgok))-(1440 + (5000 * testecl), 2760), &HFF&
Picture1.Line (1440 + (5000 * testecl), 2760)-(1440 + (5000 * d6), 2760 + (yf * altura)), &HFF&

cmcgok1.Top = (2760 - (altura * cmcgok)) - cmcgok1.Height / 2
cmcgok1.Visible = True

'Inserir grau na reta do CL
'Para "de" positivo e diferentes do que o testando
If cmg1 = 0 Then cmg1 = 0.0001

If cmcgok1 > 0 And cmcgok1 < 2 Then definal = (cmcgok1 / cmg1) * 2
If cmcgok1 > 2 And cmcgok1 < 4 Then definal = (cmcgok1 / cmg1) * 2
If cmcgok1 > 4 And cmcgok1 < 6 Then definal = (cmcgok1 / cmg1) * 2

'Para "de"negativo e diferentes do que o testando
If cmcgok1 < 0 And cmcgok1 > -2 Then definal = (cmcgok1 / cmg1) * 2
If cmcgok1 < -2 And cmcgok1 > -4 Then definal = (cmcgok1 / cmg1) * 2
If cmcgok1 < -4 And cmcgok1 > -6 Then definal = (cmcgok1 / cmg1) * 2

'Para "de"igual ao testando
'Para "de" positivo e diferentes do que o testando
If cmcgok1 = 0 Then definal = 0
If cmcgok1 = 2 Then definal = 2
If cmcgok1 = 4 Then definal = 4
If cmcgok1 = 6 Then definal = 6
If cmcgok1 = -2 Then definal = -2
If cmcgok1 = -4 Then definal = -4
If cmcgok1 = -6 Then definal = -6

If Form27.xcg < Form27.PN27 Then definalk = -1 * definal
If Form27.xcg = Form27.PN27 Then definalk = definal
If Form27.xcg > Form27.PN27 Then definalk = -1 * definal

definalk = Format(definalk, "0.00")

defgrau = definalk + "º"

defgrau.Top = 2760 + (yf * altura) - defgrau.Height / 2
defgrau.Visible = True

'Espessura do Downwash

If Form27.Option3.Value = True Then coefsust = Form7.clk2
If Form27.Option4.Value = True Then coefsust = Form7.clk4

cdmax = Form4.Data1.Recordset("cdmax")

'METODO STINTON
'cd = Abs((coefsust ^ 2) / (3.1415 * Form1.arw * 0.9) + Abs(cdmax))
'Fa = ((cd ^ 0.5) / (5 * 3.141516))
'Fb1 = Form10.LT / 0.305
'Fb2 = Form2.cmw / 0.305
'Fb = (Form1.arw * (Fb1 / Fb2)) ' A fórmula correta seria: Fb = (Form1.arw * (Fb1 / Fb2)) +7
'Ew = ((Fa * Fb) * (Form2.cmw / 0.305)) * 0.305
'Ew = Format(Ew, "0.000")

'METODO DOMMASCH
Eps = 20 * coefsust * (((1 / Form1.Afilasa) ^ 0.3) / (Form1.arw ^ 0.725)) * ((3 * Form7.CMaerodw / Form10.LT) ^ 0.25)
Ew = (Tan(Eps * 3.1415 / 180)) * Form10.LT
Ew = Format(Ew, "0.000")

Form14.filedesenho14.Enabled = True

End Sub

Private Sub Command3_Click()
Form27.Visible = False
Form14.Show
End Sub


Private Sub Command4_Click()
Form27.Picture1.AutoRedraw = True
Form27.Picture1.Cls
kde1.Visible = False
kde2.Visible = False
kde3.Visible = False
kde4.Visible = False
kde11.Visible = False
kde22.Visible = False
kde33.Visible = False

clponto.Visible = False

cmg0.Visible = False
cmg1.Visible = False
cmg2.Visible = False
cmg3.Visible = False
cmg11.Visible = False
cmg22.Visible = False
cmg33.Visible = False
defgrau.Visible = False

cmcgok1.Visible = False

End Sub

Private Sub fileaerodin_Click()
Form27.Visible = False
Form7.Visible = True
End Sub

Private Sub filecalcon_Click()
Form29.Show
End Sub

Private Sub fileddesenho_Click()
Form18.Frame22.Visible = True

'Abrir retorno ao desenho
Form1.filedesenho1.Enabled = True
Form2.filedesenho.Enabled = True
Form7.filedesenho7.Enabled = True
Form8.filedesenho8.Enabled = True
Form10.filedesenho10.Enabled = True
Form11.filedesenho11.Enabled = True
Form14.filedesenho14.Enabled = True

If Form18.afilaASA = Empty Then Form18.afilaASA = Form1.Afilasa
If Form18.AfilaHt = Empty Then Form18.AfilaHt = Form1.AfilaHt

If Form1.Afilasa <> Form18.afilaASA Then
MsgBox "O Afilamento da asa foi alterado. Clique Configurar.", vbInformation
Form18.HScroll1.Value = 0
Form18.grau = 0
Form18.bavt.Visible = False
Form18.bfvt.Visible = False
End If
If Form1.AfilaHt <> Form18.AfilaHt Then
MsgBox "O Afilamento da empenagem foi alterado. Clique Configurar.", vbInformation
Form18.HScroll2.Value = 0
Form18.grauht = 0
Form18.bavt.Visible = False
Form18.bfvt.Visible = False
End If

Form27.Visible = False
Form18.Visible = True

'Form18.vsescala = 17

Form27.Refresh

End Sub

Private Sub filedimensiona_Click()
Form27.Visible = False
Form2.Visible = True
End Sub

Private Sub fileestab_Click()
Form27.Visible = False
Form10.Visible = True
End Sub

Private Sub fileestrut_Click()
Form27.Visible = False
Form11.Visible = True
End Sub

Private Sub fileimprimi27_Click()
If LIBER <> "KFAR0852" Then
    MsgBox "A rotina de impressão sómente estará disponível após a liberação permanente", 16, "Cancelamento de impressão": Exit Sub
End If

If Form27.Option3.Value = True Then regime = "V.Subida"
If Form27.Option4.Value = True Then regime = "V.máxima"
'Pagina
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print ""
Printer.Print Tab(25); "VII - TRIMAGEM"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.FontSize = 10
Printer.Print Tab(20); "Coeficiente de momento da fuselagem:"; Tab(70); Form27.cmfuse
Printer.Print Tab(20); "Área da asa / área do estab. horizontal:"; Tab(70); Form27.shsw;
Printer.Print Tab(20); "Distancia do CG ao CA do estab. horizontal:"; Tab(70); Form27.xachcg; " /cordas"
Printer.Print Tab(20); "Coeficiente drô/dfi:"; Tab(70); Form27.drodf;
Printer.Print Tab(20); "Coeficiente depson/dalfa:"; Tab(70); Form27.dedalfa;
Printer.Print Tab(20); "Espessura 2w no estab. horizontal:"; Tab(70); Form27.Ew; " m"
Printer.Print Tab(20); "Regime estudado:"; Tab(70); regime;
Printer.Print Tab(20); "Coeficiente Cmcg:"; Tab(70); Form27.cmcgok1;
Printer.Print Tab(20); "Deflexão do estab. horizontal:"; Tab(70); Form27.defgrau;
Printer.Print ""
Printer.Print Tab(20); "Fim da página";
Printer.EndDoc
End Sub

Private Sub filepaginicio_Click()
Form27.Visible = False
Form1.Visible = True
End Sub

Private Sub fileperform_Click()
Form27.Visible = False
Form8.Visible = True
End Sub

Private Sub filepesobalanc_Click()
Form27.Visible = False
Form14.Visible = True
End Sub

Private Sub Form_Load()
HScroll1.Value = 10
End Sub

Private Sub HScroll1_Change()
valor = 500
altura = valor * HScroll1.Value
Call Command4_Click
End Sub

Private Sub Option1_Click()
testecl = Form7.clk1
End Sub

Private Sub Option2_Click()
testecl = Form7.cldecol
End Sub

Private Sub HScroll2_Change()
valcl = 0.01
testecl = valcl * HScroll2.Value
End Sub

Private Sub Option3_Click()
testecl = Form7.clk2
End Sub

Private Sub Option4_Click()
testecl = Form7.clk4
End Sub

Private Sub sai27_Click()
Form1.Visible = True
Form27.Visible = False

Form1.Refresh
End Sub
