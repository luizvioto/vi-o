VERSION 5.00
Begin VB.Form Form14 
   Caption         =   "Peso e balanceamento"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.PictureBox Picture7 
         Height          =   2535
         Left            =   5400
         ScaleHeight     =   2475
         ScaleWidth      =   6075
         TabIndex        =   93
         Top             =   0
         Width           =   6135
         Begin VB.Frame Frame2 
            Height          =   2415
            Left            =   120
            TabIndex        =   94
            Top             =   0
            Width           =   5775
            Begin VB.PictureBox Picture6 
               Height          =   375
               Left            =   3120
               ScaleHeight     =   315
               ScaleWidth      =   2355
               TabIndex        =   103
               Top             =   1920
               Width           =   2415
               Begin VB.CommandButton Option2 
                  Caption         =   "&Estimativa manual"
                  Height          =   315
                  Left            =   0
                  MouseIcon       =   "Form14.frx":030A
                  MousePointer    =   99  'Custom
                  TabIndex        =   104
                  Top             =   0
                  Width           =   2355
               End
            End
            Begin VB.PictureBox Picture3 
               Height          =   375
               Left            =   240
               ScaleHeight     =   315
               ScaleWidth      =   2355
               TabIndex        =   101
               Top             =   1920
               Width           =   2415
               Begin VB.CommandButton Option1 
                  Caption         =   "&Estimativa automática"
                  Height          =   315
                  Left            =   0
                  MouseIcon       =   "Form14.frx":0614
                  MousePointer    =   99  'Custom
                  TabIndex        =   102
                  Top             =   0
                  Width           =   2355
               End
            End
            Begin VB.Label kpeso 
               Caption         =   "Label14"
               Height          =   255
               Left            =   4800
               TabIndex        =   105
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Label37 
               Caption         =   $"Form14.frx":091E
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   97
               Top             =   1440
               Width           =   5415
            End
            Begin VB.Label Label36 
               Caption         =   $"Form14.frx":09C3
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   120
               TabIndex        =   96
               Top             =   840
               Width           =   5415
            End
            Begin VB.Label Label13 
               Caption         =   $"Form14.frx":0A7C
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   120
               TabIndex        =   95
               Top             =   240
               Width           =   5535
            End
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00C0C0C0&
         Height          =   455
         Left            =   5400
         ScaleHeight     =   390
         ScaleWidth      =   6105
         TabIndex        =   92
         Top             =   7385
         Width           =   6165
         Begin VB.CommandButton Command4 
            Caption         =   "&Cancelar"
            Height          =   385
            Left            =   4080
            MouseIcon       =   "Form14.frx":0B54
            MousePointer    =   99  'Custom
            TabIndex        =   100
            Top             =   0
            Width           =   2025
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Atualizar dados"
            Enabled         =   0   'False
            Height          =   385
            Left            =   2040
            MouseIcon       =   "Form14.frx":0E5E
            MousePointer    =   99  'Custom
            TabIndex        =   99
            Top             =   0
            Width           =   2055
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Calcular planilha"
            Height          =   385
            Left            =   0
            MouseIcon       =   "Form14.frx":1168
            MousePointer    =   99  'Custom
            TabIndex        =   98
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   4695
         Left            =   5400
         ScaleHeight     =   4635
         ScaleWidth      =   6075
         TabIndex        =   60
         Top             =   2640
         Width           =   6135
         Begin VB.Frame Frame7 
            BackColor       =   &H00400000&
            BorderStyle     =   0  'None
            ForeColor       =   &H0000FFFF&
            Height          =   5295
            Left            =   0
            TabIndex        =   61
            Top             =   0
            Width           =   6135
            Begin VB.TextBox rodadir 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   4800
               TabIndex        =   82
               Top             =   4200
               Width           =   735
            End
            Begin VB.TextBox rodaesq 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   4800
               TabIndex        =   81
               Top             =   3720
               Width           =   735
            End
            Begin VB.TextBox bequilha 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   4800
               TabIndex        =   80
               Top             =   3240
               Width           =   735
            End
            Begin VB.TextBox CGdinamico 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   3960
               TabIndex        =   66
               Top             =   1080
               Width           =   735
            End
            Begin VB.TextBox CGvazio 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   3960
               TabIndex        =   65
               Top             =   2040
               Width           =   735
            End
            Begin VB.TextBox passeio 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   3960
               TabIndex        =   64
               Top             =   2520
               Width           =   735
            End
            Begin VB.TextBox CGestatico 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   3960
               TabIndex        =   63
               Top             =   1560
               Width           =   735
            End
            Begin VB.TextBox PN 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   3960
               TabIndex        =   62
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "kgf"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   5640
               TabIndex        =   87
               Top             =   4200
               Width           =   225
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "kgf"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   5640
               TabIndex        =   86
               Top             =   3720
               Width           =   225
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "kgf"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   5640
               TabIndex        =   85
               Top             =   3240
               Width           =   225
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso nas rodas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   195
               Left            =   240
               TabIndex        =   84
               Top             =   2880
               Width           =   1095
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ponto Neutro e CGs"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   195
               Left            =   240
               TabIndex        =   83
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso na Roda direita : ..............................................................."
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   240
               TabIndex        =   79
               Top             =   4200
               Width           =   4455
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso na Roda bequilha : ............................................................"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   240
               TabIndex        =   78
               Top             =   3240
               Width           =   4500
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso na Roda esquerda : .........................................................."
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   240
               TabIndex        =   77
               Top             =   3720
               Width           =   4470
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00FFFFFF&
               Height          =   4500
               Left            =   120
               Top             =   75
               Width           =   5895
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CG calculado em Estabilidade : ..............................."
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   240
               TabIndex        =   76
               Top             =   1080
               Width           =   3645
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CG calculado nesta planilha : .................................."
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   240
               TabIndex        =   75
               Top             =   1560
               Width           =   3645
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CG do avião sem combustível : ..............................."
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   240
               TabIndex        =   74
               Top             =   2040
               Width           =   3660
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Passeio do CG : ......................................................."
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   240
               TabIndex        =   73
               Top             =   2520
               Width           =   3660
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Corda média"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   4800
               TabIndex        =   72
               Top             =   1080
               Width           =   1050
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Corda média"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   4800
               TabIndex        =   71
               Top             =   1560
               Width           =   1050
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Corda média"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   4800
               TabIndex        =   70
               Top             =   2040
               Width           =   1050
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ponto Neutro : ........................................................."
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   240
               TabIndex        =   69
               Top             =   600
               Width           =   3645
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% Corda média"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   4800
               TabIndex        =   68
               Top             =   600
               Width           =   1050
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Centímetros"
               ForeColor       =   &H0000FFFF&
               Height          =   195
               Left            =   4800
               TabIndex        =   67
               Top             =   2520
               Width           =   855
            End
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   7815
         Left            =   120
         ScaleHeight     =   7755
         ScaleWidth      =   5115
         TabIndex        =   1
         Top             =   0
         Width           =   5175
         Begin VB.TextBox text13 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   88
            Text            =   "  Piloto + Co-piloto"
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox text1 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "  Conjunto motor"
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox d1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   55
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox text4 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "  Reservatório + combustivel"
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox p2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   53
            Text            =   "0"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox d2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   52
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox text7 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "  Instrumentos"
            Top             =   3480
            Width           =   2775
         End
         Begin VB.TextBox p3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   50
            Text            =   "0"
            Top             =   3480
            Width           =   855
         End
         Begin VB.TextBox d3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   49
            Text            =   "0"
            Top             =   3480
            Width           =   735
         End
         Begin VB.TextBox text10 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "  Estabilizador Vertical"
            Top             =   3120
            Width           =   2775
         End
         Begin VB.TextBox p4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   47
            Text            =   "0"
            Top             =   3120
            Width           =   855
         End
         Begin VB.TextBox d4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   46
            Text            =   "0"
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox p5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   45
            Text            =   "0"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox d5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   44
            Text            =   "0"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox text16 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "  Ocupantes"
            Top             =   2040
            Width           =   2775
         End
         Begin VB.TextBox p6 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   42
            Text            =   "0"
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox d6 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   41
            Text            =   "0"
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox Text19 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "  Bagagem"
            Top             =   2400
            Width           =   2775
         End
         Begin VB.TextBox p7 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   39
            Text            =   "0"
            Top             =   2400
            Width           =   855
         End
         Begin VB.TextBox d7 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   38
            Text            =   "0"
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox Text28 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "  Estabilizador Horizontal"
            Top             =   2760
            Width           =   2775
         End
         Begin VB.TextBox p10 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   36
            Text            =   "0"
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox d10 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   35
            Text            =   "0"
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox Text31 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "  Asa"
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox p11 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   33
            Text            =   "0"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox d11 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   32
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox Text34 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MouseIcon       =   "Form14.frx":1472
            TabIndex        =   31
            Text            =   "  Fuselagem"
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox p12 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   30
            Text            =   "0"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox d12 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   29
            Text            =   "0"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox p13 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   28
            Text            =   "0"
            Top             =   3840
            Width           =   855
         End
         Begin VB.TextBox d13 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   27
            Text            =   "0"
            Top             =   3840
            Width           =   735
         End
         Begin VB.TextBox Text40 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "  Trem principal"
            Top             =   4200
            Width           =   2775
         End
         Begin VB.TextBox p14 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   25
            Text            =   "0"
            Top             =   4200
            Width           =   855
         End
         Begin VB.TextBox d14 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   24
            Text            =   "0"
            Top             =   4200
            Width           =   735
         End
         Begin VB.TextBox Text43 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "  Outros Pesos"
            Top             =   4560
            Width           =   2775
         End
         Begin VB.TextBox p15 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   22
            Text            =   "0"
            Top             =   4560
            Width           =   855
         End
         Begin VB.TextBox d15 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   21
            Text            =   "0"
            Top             =   4560
            Width           =   735
         End
         Begin VB.TextBox p1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3120
            TabIndex        =   20
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text37 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MouseIcon       =   "Form14.frx":177C
            TabIndex        =   19
            Text            =   "  Bequila"
            Top             =   3840
            Width           =   2775
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00400000&
            Height          =   2775
            Left            =   120
            ScaleHeight     =   2715
            ScaleWidth      =   4800
            TabIndex        =   2
            Top             =   4920
            Width           =   4860
            Begin VB.Frame Frame4 
               BackColor       =   &H00400000&
               Height          =   2610
               Left            =   120
               TabIndex        =   3
               Top             =   0
               Width           =   4575
               Begin VB.TextBox wfuel 
                  Alignment       =   2  'Center
                  BackColor       =   &H0000FFFF&
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   3360
                  TabIndex        =   89
                  Top             =   2160
                  Width           =   735
               End
               Begin VB.TextBox pt 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H00800000&
                  Height          =   285
                  Left            =   3360
                  Locked          =   -1  'True
                  TabIndex        =   8
                  Top             =   240
                  Width           =   735
               End
               Begin VB.TextBox w13 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H00800000&
                  Height          =   285
                  Left            =   3360
                  Locked          =   -1  'True
                  TabIndex        =   7
                  Top             =   600
                  Width           =   735
               End
               Begin VB.TextBox ptp 
                  Alignment       =   2  'Center
                  BackColor       =   &H0000FFFF&
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   3360
                  Locked          =   -1  'True
                  TabIndex        =   6
                  Top             =   1440
                  Width           =   735
               End
               Begin VB.TextBox pbq 
                  Alignment       =   2  'Center
                  BackColor       =   &H0000FFFF&
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   3360
                  Locked          =   -1  'True
                  TabIndex        =   5
                  Top             =   1080
                  Width           =   735
               End
               Begin VB.TextBox cg 
                  Alignment       =   2  'Center
                  BackColor       =   &H0000FFFF&
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   3360
                  Locked          =   -1  'True
                  TabIndex        =   4
                  Top             =   1800
                  Width           =   735
               End
               Begin VB.Label Label40 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Peso do avião sem combustível : "
                  ForeColor       =   &H0000FFFF&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   91
                  Top             =   2160
                  Width           =   2400
               End
               Begin VB.Label Label46 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "kgf"
                  ForeColor       =   &H0000FFFF&
                  Height          =   195
                  Left            =   4200
                  TabIndex        =   90
                  Top             =   2160
                  Width           =   225
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Peso total do avião (da tabela acima) :"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   18
                  Top             =   240
                  Width           =   2715
               End
               Begin VB.Label Label27 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Peso total do avião (calculado no projeto) :"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   17
                  Top             =   600
                  Width           =   3030
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "kgf"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   4200
                  TabIndex        =   16
                  Top             =   240
                  Width           =   225
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "kgf"
                  ForeColor       =   &H00FFFFFF&
                  Height          =   195
                  Left            =   4200
                  TabIndex        =   15
                  Top             =   600
                  Width           =   225
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Carga no trem principal :"
                  ForeColor       =   &H0000FFFF&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   14
                  Top             =   1440
                  Width           =   1710
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Carga na bequilha :"
                  ForeColor       =   &H0000FFFF&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   13
                  Top             =   1080
                  Width           =   1380
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Localização do Cg em relação à L.R"
                  ForeColor       =   &H0000FFFF&
                  Height          =   195
                  Left            =   120
                  TabIndex        =   12
                  Top             =   1800
                  Width           =   2580
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "kgf"
                  ForeColor       =   &H0000FFFF&
                  Height          =   195
                  Left            =   4200
                  TabIndex        =   11
                  Top             =   1080
                  Width           =   225
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "kgf"
                  ForeColor       =   &H0000FFFF&
                  Height          =   195
                  Left            =   4200
                  TabIndex        =   10
                  Top             =   1440
                  Width           =   225
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "cm"
                  ForeColor       =   &H0000FFFF&
                  Height          =   195
                  Left            =   4200
                  TabIndex        =   9
                  Top             =   1800
                  Width           =   210
               End
            End
         End
         Begin VB.Label Label3 
            Caption         =   "Peso (kgf)"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3160
            MouseIcon       =   "Form14.frx":1A86
            MousePointer    =   99  'Custom
            TabIndex        =   59
            ToolTipText     =   "Peso de cada equipamento isoladamente ou em conjunto"
            Top             =   30
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Local (cm)"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   4200
            MouseIcon       =   "Form14.frx":1EC8
            MousePointer    =   99  'Custom
            TabIndex        =   58
            ToolTipText     =   "Distancia em relação ao eixo da hélice"
            Top             =   25
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Equipamentos principais"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   25
            Width           =   1710
         End
      End
   End
   Begin VB.Menu fileir14 
      Caption         =   "&Ir para ..."
      Begin VB.Menu filetrim 
         Caption         =   "&Trimagem"
      End
      Begin VB.Menu filestrut 
         Caption         =   "Estrutura"
      End
      Begin VB.Menu fileestab 
         Caption         =   "Estabilidade estática"
      End
      Begin VB.Menu fileperform14 
         Caption         =   "Performance"
      End
      Begin VB.Menu fileaerodin14 
         Caption         =   "Aerodinâmica"
      End
      Begin VB.Menu filedimensiona14 
         Caption         =   "Dimensionamento"
      End
      Begin VB.Menu fileprincipia14 
         Caption         =   "Principal"
      End
      Begin VB.Menu filedesenho14 
         Caption         =   "Retornar ao desenho"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu filevisual14 
      Caption         =   "&Visualizar"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu filetabservo 
      Caption         =   "&CGservos"
   End
   Begin VB.Menu filefer14 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "&Calculadora && Conversor"
      End
   End
   Begin VB.Menu fileimprimi 
      Caption         =   "Im&primir"
   End
   Begin VB.Menu sai14 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'Trazendo o peso(wo) do form1 em gramas
Form14.w13 = Form1.wo
Form14.w13 = Format(Form14.w13, "0.00")

If p13 = 0 Or d13 = 0 Or p14 = 0 Or d14 = 0 Then
MsgBox "Peso ou distancias das rodas devem ser diferentes de zero", vbInformation
Exit Sub
End If

k1 = Form14.p1 * Form14.d1
k2 = Form14.p2 * Form14.d2
k3 = Form14.p3 * Form14.d3
k4 = Form14.p4 * Form14.d4
k5 = Form14.p5 * Form14.d5
k6 = Form14.p6 * Form14.d6
k7 = Form14.p7 * Form14.d7
k10 = Form14.p10 * Form14.d10
k11 = Form14.p11 * Form14.d11
k12 = Form14.p12 * Form14.d12
k13 = Form14.p13 * Form14.d13
k14 = Form14.p14 * Form14.d14
k15 = Form14.p15 * Form14.d15

kt = Abs(k1) + Abs(k2) + Abs(k3) + Abs(k4) + Abs(k5) + Abs(k6) + Abs(k7) + Abs(k9) + Abs(k10) + Abs(k11) + Abs(k12) + Abs(k13) + Abs(k14) + Abs(k15)

pt = Abs(p1) + Abs(p2) + Abs(p3) + Abs(p4) + Abs(p5) + Abs(p6) + Abs(p7) + Abs(p9) + Abs(p10) + Abs(p11) + Abs(p12) + Abs(p13) + Abs(p14) + (p15)
pt = Format(pt, "0.00")


If Form1.tremtric.Value = True And Val(d13) > Val(d14) Then
    MsgBox "O trem de pouso é Triciclo"
    d13.SetFocus
        Exit Sub
End If

If Form1.tremconv.Value = True And Val(d13) < Val(d14) Then
    MsgBox "O trem de pouso é Convencional"
    d13.SetFocus
        Exit Sub
End If

cg = kt / pt
cg = Format(cg, "0.00")


If Val(d14) > Val(d13) And Val(cg) > Val(d13) And Val(cg) < Val(d14) Then
    k1 = Abs(cg) - Abs(d13)
    k2 = Abs(d14) - Abs(d13)
    k3 = k1 / k2
    ptp = k3 * pt
    pbq = Abs(pt) - Abs(ptp)
    pbq = Format(pbq, "0.00")
    ptp = Format(ptp, "0.00")
End If

    
If Val(d14) < Val(d13) And Val(cg) > Val(d14) And Val(cg) < Val(d13) Then
    k1 = Abs(cg) - Abs(d14)
    k2 = Abs(d13) - Abs(d14)
    k3 = k1 / k2
    pbq = k3 * pt
    ptp = Abs(pt) - Abs(pbq)
    pbq = Format(pbq, "0.00")
    ptp = Format(ptp, "0.00")
End If

If Val(d14) = Val(cg) Then
    pbq = 0
    ptp = Val(pt)
    MsgBox "Bequilha com carga nula! Modifique a posição do trem de pouso principal. Se Convencional mover para frente e se Triciclo mover para trás", vbExclamation
End If


'Para cargas na bequilha negativas
If Val(d14) > Val(d13) And Val(cg) > Val(d14) And Val(cg) > Val(d13) Then
    k1 = Abs(cg) - Abs(d14)
    k2 = Abs(d14) - Abs(d13)
    pbq = (-1 * pt * k1) / k2
    ptp = pt
    pbq = Format(pbq, "0.00")
    ptp = Format(ptp, "0.00")
    
MsgBox "Peso excessivo na cauda do avião!  A bequilha esta levantando!  Verifique Pesos e Estações", vbInformation
Exit Sub
End If

If Val(d14) < Val(d13) And Val(cg) < Val(d14) And Val(cg) < Val(d13) Then
    k1 = Abs(d14) - Abs(cg)
    k2 = Abs(d13) - Abs(d14)
    pbq = (-1 * pt * k1) / k2
    ptp = pt
    pbq = Format(pbq, "0.00")
    ptp = Format(ptp, "0.00")
    
MsgBox "Peso excessivo no nariz do avião!  A bequilha esta levantando!  Verifique Pesos e Estações", vbOKOnly + vbInformation + vbDefaultButton1, "Peso & Balanceamento"
Exit Sub
End If

If Int(w13) > Int(pt) Then
                    MsgBox "Peso de projeto diferente do Peso calculado - Diminua o peso inicial do projeto ou aumente o peso dos equipamentos", vbOKOnly + vbCritical + vbDefaultButton1, "Peso & Balanceamento"
                    Exit Sub
End If
                    

If Int(w13) < Int(pt) Then
                    MsgBox "Peso de projeto diferente do Peso calculado - Aumente o peso inicial do projeto ou diminua o peso dos equipamentos", vbOKOnly + vbCritical + vbDefaultButton1, "Peso & Balanceamento"
                    Exit Sub
End If

Form14.PN = Form10.PN

Form14.CGdinamico = Form10.cg
Form14.CGdinamico = Format(Form14.CGdinamico, "0.000")

'distancia da LR ao B.ataq.
dba = Abs(Form10.LRCAasa * 100) - (0.25 * Form7.CMaerodw * 100)

'CGestático
Form14.CGestatico = ((Abs(Form14.cg) - Abs(dba)) / (Form7.CMaerodw * 100)) * 100
Form14.CGestatico = Format(Form14.CGestatico, "0.000")

'CG do avião sem combustível
k1 = Form14.p1 * Form14.d1
k3 = Form14.p3 * Form14.d3
k4 = Form14.p4 * Form14.d4
k5 = Form14.p5 * Form14.d5
k6 = Form14.p6 * Form14.d6
k7 = Form14.p7 * Form14.d7
k10 = Form14.p10 * Form14.d10
k11 = Form14.p11 * Form14.d11
k12 = Form14.p12 * Form14.d12
k13 = Form14.p13 * Form14.d13
k14 = Form14.p14 * Form14.d14
k15 = Form14.p15 * Form14.d15

kscomb = Abs(k1) + Abs(k3) + Abs(k4) + Abs(k5) + Abs(k6) + Abs(k7) + Abs(k9) + Abs(k10) + Abs(k11) + Abs(k12) + Abs(k13) + Abs(k14) + Abs(k15)

ptscomb = Abs(p1) + Abs(p3) + Abs(p4) + Abs(p5) + Abs(p6) + Abs(p7) + Abs(p9) + Abs(p10) + Abs(p11) + Abs(p12) + Abs(p13) + Abs(p14) + Abs(p15)
ptscomb = Format(ptscomb, "0.00")

CGscomb = kscomb / ptscomb
Form14.CGvazio = ((CGscomb - dba) / (Form7.CMaerodw * 100)) * 100
Form14.CGvazio = Format(Form14.CGvazio, "0.000")

'Calculo do Passeio
' Para que o valor do passeio fique positivo fo inserido "Abs"

cgpcm = ((CGestatico / 100) * Form7.CMaerodw) * 100
cgpcm = Format(cgpcm, "0.00")

cgvzcm = ((CGvazio / 100) * Form7.CMaerodw) * 100
cgvzcm = Format(cgvzcm, "0.00")

passeio = Abs(Abs(cgvzcm) - Abs(cgpcm))
passeio = Format(passeio, "0.00")

Form14.bequilha = Form14.pbq
Form14.rodaesq = Form14.ptp / 2
Form14.rodadir = Form14.ptp / 2
Form14.wfuel = Val(pt) - Val(p2)

    'FINAL
    Form14.p15 = Abs(Form1.wo) - ((Abs(p1) + Abs(p2) + Abs(p3) + Abs(p4) + Abs(p5) + Abs(p6) + Abs(p7) + Abs(p9) + Abs(p10) + Abs(p11) + Abs(p12) + Abs(p13) + Abs(p14)))
    Form14.p15 = Format(Form14.p15, "0.00")


'Conclusão do CG - FORA!
'If Int(Form14.CGdinamico) <> Int(Form14.CGestatico) Then
'MsgBox "CUIDADO! CG calculado em Estabilidade diferente do CG calculado em Peso/balanceamento. Modifique estações.", vbOKOnly + vbCritical + vbDefaultButton1, "Peso & Balanceamento"
'If Val(Form14.CGvazio) > Val(Form14.PN) Then MsgBox "CUIDADO! CG vazio além do ponto neutro. Instabilidade!", vbCritical
'Exit Sub
'End If

'Conclusão do CG - OK!
'If Int(Form14.CGdinamico) = Int(Form14.CGestatico) Then
'MsgBox "OK! CG calculado em estabilidade corresponde ao calculado nesta planilha (P&B).", vbOKOnly + vbInformation + vbDefaultButton1, "Peso & Balanceamento"
'If Val(Form14.CGvazio) > Val(Form14.PN) Then MsgBox "CUIDADO! CG vazio além do ponto neutro. Instabilidade!", vbCritical
'Form14.filevisual14.Enabled = True
'Form14.filevisual14.Visible = True
'Exit Sub
'End If

'Conclusão do CG - OK! Modificado em 02/05/2004 SJC

If Int(Form14.CGestatico) >= Int(Form14.CGdinamico) And Int(Form14.CGestatico) < Form14.PN Then
MsgBox "OK! CG calculado em estabilidade corresponde ao calculado nesta planilha (P&B).", vbOKOnly + vbInformation + vbDefaultButton1, "Peso & Balanceamento"
End If

'Conclusão do CG - FORA! Modificado em 02/05/2004 SJC
If Val(Form14.CGvazio) > Val(Form14.PN) Then
MsgBox "CUIDADO! CG vazio além do ponto neutro. Instabilidade!", vbCritical
Form14.filevisual14.Enabled = True
Form14.filevisual14.Visible = True
Exit Sub
End If

If Val(Form14.CGvazio) < Int(Form14.CGdinamico) Then
MsgBox "CUIDADO! CG vazio muito adiantado. Excesso de estabilidade!", vbCritical
Form14.filevisual14.Enabled = True
Form14.filevisual14.Visible = True
Exit Sub
End If

End Sub

Private Sub Command2_Click()
' ooooooooooooo ABRINDO PESO E BALANCEAMENTO ooooooooooooooo

If Form21.Data1.Recordset("Projeto") = Form20.p1 Then

    resp2000 = MsgBox("SIM = restaurar valores salvos  e  NÃO = recalcular novos valores.", vbYesNo)

        If resp2000 = vbYes Then

            Form14.p1 = Form21.Data1.Recordset("peso1")
            Form14.d1 = Form21.Data1.Recordset("dist1")

            Form14.p2 = Form21.Data1.Recordset("peso2")
            Form14.d2 = Form21.Data1.Recordset("dist2")

            Form14.p11 = Form21.Data1.Recordset("peso3")
            Form14.d11 = Form21.Data1.Recordset("dist3")

            Form14.p12 = Form21.Data1.Recordset("peso4")
            Form14.d12 = Form21.Data1.Recordset("dist4")

            Form14.p5 = Form21.Data1.Recordset("peso5")
            Form14.d5 = Form21.Data1.Recordset("dist5")

            Form14.p6 = Form21.Data1.Recordset("peso6")
            Form14.d6 = Form21.Data1.Recordset("dist6")

            Form14.p7 = Form21.Data1.Recordset("peso7")
            Form14.d7 = Form21.Data1.Recordset("dist7")

            Form14.p10 = Form21.Data1.Recordset("peso8")
            Form14.d10 = Form21.Data1.Recordset("dist8")

            Form14.p4 = Form21.Data1.Recordset("peso9")
            Form14.d4 = Form21.Data1.Recordset("dist9")

            Form14.p3 = Form21.Data1.Recordset("peso10")
            Form14.d3 = Form21.Data1.Recordset("dist10")

            Form14.p13 = Form21.Data1.Recordset("peso11")
            Form14.d13 = Form21.Data1.Recordset("dist11")

            Form14.p14 = Form21.Data1.Recordset("peso12")
            Form14.d14 = Form21.Data1.Recordset("dist12")

            Form14.p15 = Form21.Data1.Recordset("peso13")
            Form14.d15 = Form21.Data1.Recordset("dist13")
            
        End If
End If
            
End Sub




Private Sub Command4_Click()

Form14.p1 = 0
Form14.d1 = 0

Form14.p2 = 0
Form14.d2 = 0

Form14.p3 = 0
Form14.d3 = 0

Form14.p4 = 0
Form14.d4 = 0

Form14.p5 = 0
Form14.d5 = 0

Form14.p6 = 0
Form14.d6 = 0

Form14.p7 = 0
Form14.d7 = 0

Form14.p14 = 0
Form14.d14 = 0

Form14.p15 = 0
Form14.d15 = 0

Form14.p10 = 0
Form14.d10 = 0

Form14.p11 = 0
Form14.d11 = 0

Form14.p12 = 0
Form14.d12 = 0

Form14.p13 = 0
Form14.d13 = 0



End Sub


Private Sub fileaerodin14_Click()
Form7.Visible = True
Form14.Visible = False
End Sub


Private Sub filecalcon_Click()
Form29.Show
End Sub

Private Sub filedesenho14_Click()
If Form27.testecl = Empty Or Form27.testecl = 0 Then
    MsgBox "A trimagem deve ser calculada", vbInformation
    Exit Sub
End If

If Form14.CGdinamico = Empty And Form14.CGestatico = Empty Then
    MsgBox "O CG deve ser calculado", vbInformation
    Exit Sub
End If

Form18.Visible = True
Form14.Visible = False
Form18.Refresh
End Sub

Private Sub filedimensiona14_Click()
Form2.Visible = True
Form14.Visible = False
Form2.Refresh
End Sub

Private Sub fileestab_Click()
Form14.Visible = False
Form10.Visible = True
End Sub

Private Sub fileimprimi_Click()
If LIBER <> "KFAR0852" Then
    MsgBox "A rotina de impressão sómente estará disponível após a liberação permanente", 16, "Cancelamento de impressão": Exit Sub
End If

Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "VI - PESO E BALANCEAMENTO"
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 10
Printer.Print ""
Printer.Print Tab(20); "ABREVIATURAS"
Printer.Print Tab(20); "C.G = Centro de gravidade"
Printer.Print Tab(20); "L.R = Linha de referencia - eixo da hélice "
Printer.Print Tab(20); "P.N = Ponto neutro"
Printer.Print ""
Printer.Print ""
Printer.Print Tab(25); "TABELA"
Printer.Print ""
Printer.Print Tab(20); "EQUIPAMENTO:"; Tab(65); "PESO (gr)"; Tab(80); "ESTAÇÃO (cm)";
Printer.Print Tab(20); "Conjunto motor:"; Tab(65); Form14.p1; Tab(80); Form14.d1;
Printer.Print Tab(20); "Reservatório + combustível:"; Tab(65); Form14.p2; Tab(80); Form14.d2;
Printer.Print Tab(20); "Asa:"; Tab(65); Form14.p11; Tab(80); Form14.d11;
Printer.Print Tab(20); "Fuselagem:"; Tab(65); Form14.p12; Tab(80); Form14.d12;
Printer.Print Tab(20); "Piloto e co-piloto:"; Tab(65); Form14.p5; Tab(80); Form14.d5;
Printer.Print Tab(20); "Ocupantes:"; Tab(65); Form14.p6; Tab(80); Form14.d6;
Printer.Print Tab(20); "Bagagem:"; Tab(65); Form14.p7; Tab(80); Form14.d7;
Printer.Print Tab(20); "Estabilizador horizontal:"; Tab(65); Form14.p10; Tab(80); Form14.d10;
Printer.Print Tab(20); "Estabilizador vertical:"; Tab(65); Form14.p4; Tab(80); Form14.d4;
Printer.Print Tab(20); "Instrumentos:"; Tab(65); Form14.p3; Tab(80); Form14.d3;
Printer.Print Tab(20); "Bequilha:"; Tab(65); Form14.p13; Tab(80); Form14.d13;
Printer.Print Tab(20); "Trem principal:"; Tab(65); Form14.p14; Tab(80); Form14.d14;
Printer.Print Tab(20); "Outros:"; Tab(65); Form14.p15; Tab(80); Form14.d15;
Printer.Print ""
Printer.Print ""
Printer.Print Tab(25); "CALCULOS"
Printer.Print ""
Printer.Print Tab(20); "Peso total do avião nesta planilha:"; Tab(70); Form14.pt; " kgf"
Printer.Print Tab(20); "Peso total do avião de projeto:"; Tab(70); Form14.w13; " kgf"
Printer.Print Tab(20); "Carga na bequilha:"; Tab(70); Form14.pbq; " kgf"
Printer.Print Tab(20); "Carga no trem principal:"; Tab(70); Form14.ptp; " kgf"
Printer.Print Tab(20); "Localização do C.G em relação a L.R:"; Tab(70); Form14.cg; " cm"
Printer.Print Tab(20); "Peso do avião sem combustível:"; Tab(70); Form14.wfuel; " kgf"
Printer.Print ""
Printer.Print ""
Printer.Print Tab(25); "PONTO NEUTRO E C.Gs"
Printer.Print ""
Printer.Print Tab(20); "Ponto Neutro:"; Tab(70); Form14.PN; " %Cmedia"
Printer.Print Tab(20); "C.G calculado em Estabilidade:"; Tab(70); Form14.CGdinamico; " %Cmedia"
Printer.Print Tab(20); "C.G calculado nesta planilha:"; Tab(70); Form14.CGestatico; " %Cmedia"
Printer.Print Tab(20); "C.G do avião sem combustível:"; Tab(70); Form14.CGvazio; " %Cmedia"
Printer.Print Tab(20); "Passeio do C.G:"; Tab(70); Form14.passeio; "; Cm; """
Printer.Print Tab(20); "Peso na roda - bequilha:"; Tab(70); Form14.bequilha; " kgf"
Printer.Print Tab(20); "Peso na roda esquerda:"; Tab(70); Form14.rodaesq; " kgf"
Printer.Print Tab(20); "Peso na roda direita:"; Tab(70); Form14.rodadir; " kgf"
Printer.Print ""
Printer.Print ""
Printer.Print Tab(20); "Fim da página";
Printer.EndDoc

End Sub

Private Sub fileperform14_Click()
Form10.Visible = True
Form14.Visible = False
End Sub

Private Sub filepesobal14_Click()
Form10.Visible = True
Form14.Visible = False
End Sub

Private Sub fileprincipia14_Click()
Form1.Visible = True
Form14.Visible = False

Form1.Refresh
End Sub
Private Sub filestrut_Click()
Form14.Visible = False
Form11.Visible = True
End Sub



Private Sub filetrim_Click()

If Form14.p1 = Empty Or Form14.w13 = Empty Then
MsgBox "A planilha deve ser calculada!", vbCritical
Exit Sub
End If

If Form14.p1 <> Empty Or Form14.w13 <> Empty Then
Form27.Visible = True
Form27.clafaw = Form7.slopeasa
Form27.clafaw = Format(Form27.clafaw, "0.000")

Form27.clalfah = Form7.slopeht
Form27.clalfah = Format(Form27.clalfah, "0.000")

Form27.xcg = (Form10.LRCG) / Form7.CMaerodw
Form27.xcg = Format(Form27.xcg, "0.000")

Form27.xacw = Form10.LRCAasa / Form7.CMaerodw
Form27.xacw = Format(Form27.xacw, "0.000")

Form27.cmfuse = ((0.005 * (Form2.wfs ^ 2) * Form2.lfs) / (Form7.CMaerodw * Form2.Sw)) * (180 / 3.1415)
Form27.cmfuse = Format(Form27.cmfuse, "0.000")

Form27.nih = Form10.eficiht
Form27.nih = Format(Form27.nih, "0.000")

Form27.shsw = (Form2.sht / Form2.Sw)
Form27.shsw = Format(Form27.shsw, "0.000")

Form27.xachcg = Abs((Form10.LRCAht / Form7.CMaerodw) - (Form10.LRCG / Form7.CMaerodw))
Form27.xachcg = Format(Form27.xachcg, "0.000")

Form27.cecp = 0.3
Form27.cecp = Format(Form27.cecp, "0.000")

Form27.drodf = 4.2
Form27.drodf = Format(Form27.drodf, "0.000")

Form27.dedalfa = Form10.DEDA
Form27.dedalfa = Format(Form27.dedalfa, "0.000")

If Form1.margestat = 0 Then Form27.PN27 = Form27.xcg

If Form1.margestat <> 0 Then
Form27.PN27 = Form10.LRPN / Form7.CMaerodw
Form27.PN27 = Format(Form27.PN27, "0.000")

Form27.HScroll1.Value = 16
End If

End If

Form14.Visible = False

End Sub

Private Sub filevisual14_Click()
'Retirei a condição abaixo em 02/05/2004 em SJC devido a nova variação do CG, ver em no botão Calcular Planilha
''If Int(Form14.CGdinamico) <> Int(Form14.CGestatico) Then
'MsgBox "Verifique se os valores de CGs estão corretos"
'Exit Sub
'End If

Form15.Visible = True

If Form1.posialta.Value = True Or Form1.posibaixa.Value = True Or Form1.posimedia.Value = True Then

If Form1.tremconv.Value = True Then
Form15.beqconv.Visible = True
Form15.beqtri.Visible = False
Form15.rd.Left = 2520
Form15.re.Left = 2520
Form15.beqconv = Form14.pbq
End If

If Form1.tremtric.Value = True Then
Form15.beqconv.Visible = False
Form15.beqtri.Visible = True
Form15.rd.Left = 3120
Form15.re.Left = 3120
Form15.beqtri = Form14.pbq
End If

End If

Form15.re = Form14.ptp / 2
Form15.rd = Form14.ptp / 2

End Sub




Private Sub Option1_Click()
Call Option2_Click
' TRAVA DE CELULAS

Form14.p1.Locked = True

Form14.p2.Locked = True

Form14.p11.Locked = True
Form14.d11.Locked = True

Form14.p12.Locked = True
Form14.d12.Locked = True

Form14.p5.Locked = True
Form14.d5.Locked = True

Form14.p6.Locked = True
Form14.d6.Locked = True

Form14.p7.Locked = True
Form14.d7.Locked = True

Form14.p10.Locked = True
Form14.d10.Locked = True

Form14.p4.Locked = True

Form14.p3.Locked = True

Form14.p13.Locked = True

Form14.p14.Locked = True



If Form21.Data1.Recordset.RecordCount = 0 Then GoTo 100

If Form21.Data1.Recordset("projeto") = Form20.p1 Then
    resp1 = MsgBox("O projeto já está salvo portanto, os valores serão recalculados", vbYesNo)
    
        If resp1 = vbYes Then GoTo 100
        
        If resp1 = vbNo Then
            Exit Sub
        End If
End If


If Form21.Data1.Recordset("projeto") <> Form20.p1 Then GoTo 100

100:
N = 4

dar = Form1.roar * 0.00194
vemapes = Form8.vmax * 0.9113
Qd = (dar * (vemapes ^ 2)) / 2
Qd = Format(Qd, "0.0000")


'Estação Motor
Form14.p1 = 0.12 * (Form1.wo / 2.2) 'era 0.22 antes de 29/11/2007
Form14.p1 = Format(Form14.p1, "0.00")

If Form1.pull.Value = True Then
Form14.d1 = (0.05 * Form2.lfs) * 100
Form14.d1 = Format(Form14.d1, "0.00")
End If

If Form1.pusher.Value = True Then
Form14.d1 = (0.9 * Form2.lfs) * 100
Form14.d1 = Format(Form14.d1, "0.00")
End If

'Peso e Estação ASA e combustivel na asa

Form14.p11 = (0.14 * Form1.wo)
Form14.p11 = Format(Form14.p11, "0.00")

Form14.d11 = ((Form10.LRCAasa - 0.25 * Form7.CMaerodw) + (0.4 * Form2.crw)) * 100
Form14.d11 = Format(Form14.d11, "0.00")

'Em 05/12/2007
'Estação do Piloto
Form14.d5 = Val(Form14.d11) - 85

'Estação do Co-piloto
Form14.d6 = Val(Form14.d11) + 85

'xxxxxxxxxxxxxxxxx

'Estação Tanque
Form14.p2 = Val(Requisitos.wfR / 2.2) + (0.005 * Form1.wo)
Form14.p2 = Format(Form14.p2, "0.00")
'estação na asa
Form14.d2 = d11
Form14.d2 = Format(Form14.d2, "0.00")

'Peso e Estação Fuselagem
    
        Form14.p12 = (0.1 * Form1.wo)
        Form14.p12 = Format(Form14.p12, "0.00")

        Form14.d12 = 0.4 * Form2.lfs * 100
        Form14.d12 = Format(Form14.d12, "0.00")


'Peso do trem principal
    Form14.p14 = (0.03 * Form1.wo)
    Form14.p14 = Format(Form14.p14, "0.00")

If Form1.tremtric.Value = True Then
    Form14.d14 = (Form10.LRPN) * 100
    Form14.d14 = Format(Form14.d14, "0.00")
End If

If Form1.tremconv.Value = True Then
    Form14.d14 = (0.8 * Form10.LRCAasa) * 100
    Form14.d14 = Format(Form14.d14, "0.00")
End If


'Peso da bequilha
If Form1.tremtric.Value = True Then Form14.p13 = (0.015 * Form1.wo)

If Form1.tremconv.Value = True Then Form14.p13 = (0.005 * Form1.wo)

If Form1.tremtric.Value = True Then
    If Form1.pull.Value = True Then Form14.d13 = 1.25 * Form14.d1
    If Form1.pusher.Value = True Then Form14.d13 = (0.075 * Form2.lfs) * 100
    Form14.d13 = Format(Form14.d13, "0.00")
End If

If Form1.tremconv.Value = True Then
    Form14.d13 = 5 + ((Form2.lfs) * 100)
    Form14.d13 = Format(Form14.d13, "0.00")
 End If

'Piloto e copiloto
p5 = 95

'Ocupantes
 p6 = ((Requisitos.ocup - 1) * 95)

'Bagagem
p7 = (Requisitos.bag / 2.2) + ((Requisitos.CPR / 2.2) - (Requisitos.wfR / 2.2))
d7 = d11

'Estabilizador horizontal
p10 = 0.02 * (Form1.wo)

'Estabilizador Vertical
p4 = 0.01 * (Form1.wo)

'Instrumentos e Controles
p3 = 0.07 * (Form1.wo)
    
               
p3 = Format(p3, "0.00")
p4 = Format(p4, "0.00")
p5 = Format(p5, "0.00")
p6 = Format(p6, "0.00")
p7 = Format(p7, "0.00")
p10 = Format(p10, "0.00")
Form14.p13 = Format(Form14.p13, "0.00")


        d3 = Val(d5) - 60
        d3 = Format(d3, "0.00")
        
        d4 = Form10.LRCAht * 100
        d4 = Format(d4, "0.00")
        
        d10 = Form10.LRCAht * 100
        d10 = Format(d10, "0.00")
        
'Estação da Bequilha
If Form1.tremconv.Value = True Then lfs13 = Form10.LRCAht * 100
If Form1.tremtric.Value = True Then lfs13 = (Form2.lfs * 0.05) * 100

        
'Peso de Bagagem & OUTROS
Form14.p15 = Abs(Form1.wo) - ((Abs(p1) + Abs(p2) + Abs(p3) + Abs(p4) + Abs(p5) + Abs(p6) + Abs(p7) + Abs(p9) + Abs(p10) + Abs(p11) + Abs(p12) + Abs(p13) + Abs(p14)))
Form14.p15 = Format(Form14.p15, "0.000")

'If p15 >= 0 Then Exit Sub


If p15 > 0 Then
kdist = Form14.p15 / 7

p11 = p11 + kdist
p11 = Format(p11, "0.00")

p12 = p12 + kdist
p12 = Format(p12, "0.00")

p10 = p10 + kdist
p10 = Format(p10, "0.00")

p3 = p3 + kdist
p3 = Format(p3, "0.00")

p4 = p4 + kdist
p4 = Format(p4, "0.00")

p13 = p13 + kdist
p13 = Format(p13, "0.00")

p14 = p14 + kdist
p14 = Format(p14, "0.00")

Form14.p15 = Abs(Form1.wo) - ((Abs(p1) + Abs(p2) + Abs(p3) + Abs(p4) + Abs(p5) + Abs(p6) + Abs(p7) + Abs(p9) + Abs(p10) + Abs(p11) + Abs(p12) + Abs(p13) + Abs(p14)))
Form14.p15 = Format(Form14.p15, "0.000")


Exit Sub
End If

If Val(p15) < 0 Then

                    MsgBox "O valor de ""Outros Pesos"" é menor do que zero, distribua o valor entre outros pesos manualmente", vbOKOnly + vbCritical + vbDefaultButton1, "Peso & Balanceamento"
                    Form29.Show
                    Form29.Label3.Visible = True
                    Form29.difpeso.Visible = True
                    Form29.difpeso = p15
                    p15 = 0
                    
End If
              
End Sub

Private Sub Option2_Click()

Form14.d5.SetFocus


Form14.p1.Locked = False
Form14.d1.Locked = False

Form14.p2.Locked = False
Form14.d2.Locked = False

Form14.p11.Locked = False
Form14.d11.Locked = False

Form14.p12.Locked = False
Form14.d12.Locked = False

Form14.p5.Locked = False
Form14.d5.Locked = False

Form14.p6.Locked = False
Form14.d6.Locked = False

Form14.p7.Locked = False
Form14.d7.Locked = False

Form14.p10.Locked = False
Form14.d10.Locked = False

Form14.p4.Locked = False
Form14.d4.Locked = False

Form14.p3.Locked = False
Form14.d3.Locked = False

Form14.p13.Locked = False
Form14.d13.Locked = False

Form14.p14.Locked = False
Form14.d14.Locked = False

Form14.p15.Locked = False
Form14.d15.Locked = False

End Sub

Private Sub p1_GotFocus()
p1.SelStart = 0
p1.SelLength = Len(p1)
End Sub

Private Sub p1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d1.SetFocus
End Sub
Private Sub p10_GotFocus()
p10.SelStart = 0
p10.SelLength = Len(p10)
End Sub
Private Sub p10_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d10.SetFocus
End Sub
Private Sub p11_GotFocus()
p11.SelStart = 0
p11.SelLength = Len(p11)
End Sub
Private Sub p11_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d11.SetFocus
End Sub
Private Sub p12_GotFocus()
p12.SelStart = 0
p12.SelLength = Len(p12)
End Sub
Private Sub p12_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d12.SetFocus
End Sub
Private Sub p13_GotFocus()
p13.SelStart = 0
p13.SelLength = Len(p13)
End Sub
Private Sub p13_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d13.SetFocus
End Sub


Private Sub p14_GotFocus()
p14.SelStart = 0
p14.SelLength = Len(p14)
End Sub
Private Sub p14_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d14.SetFocus
End Sub
Private Sub p15_GotFocus()
p15.SelStart = 0
p15.SelLength = Len(p15)
End Sub
Private Sub p15_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d15.SetFocus
End Sub
Private Sub p16_GotFocus()
p16.SelStart = 0
p16.SelLength = Len(p16)
End Sub
Private Sub p16_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d16.SetFocus
End Sub
Private Sub p2_GotFocus()
p2.SelStart = 0
p2.SelLength = Len(p2)
End Sub
Private Sub p2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d2.SetFocus
End Sub

Private Sub p3_GotFocus()
p3.SelStart = 0
p3.SelLength = Len(p3)
End Sub
Private Sub p3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d3.SetFocus
End Sub
Private Sub p4_GotFocus()
p4.SelStart = 0
p4.SelLength = Len(p4)
End Sub
Private Sub p4_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d4.SetFocus
End Sub
Private Sub p5_GotFocus()
p5.SelStart = 0
p5.SelLength = Len(p5)
End Sub
Private Sub p5_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d5.SetFocus
End Sub
Private Sub p6_GotFocus()
p6.SelStart = 0
p6.SelLength = Len(p6)
End Sub
Private Sub p6_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d6.SetFocus
End Sub
Private Sub p7_GotFocus()
p7.SelStart = 0
p7.SelLength = Len(p7)
End Sub
Private Sub p7_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then d7.SetFocus
End Sub

'para d
Private Sub d1_GotFocus()
d1.SelStart = 0
d1.SelLength = Len(d1)
End Sub
Private Sub d1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p2.SetFocus
End Sub
Private Sub d10_GotFocus()
d10.SelStart = 0
d10.SelLength = Len(d10)
End Sub
Private Sub d10_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p4.SetFocus
End Sub
Private Sub d11_GotFocus()
d11.SelStart = 0
d11.SelLength = Len(d11)
End Sub
Private Sub d11_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p12.SetFocus
End Sub
Private Sub d12_GotFocus()
d12.SelStart = 0
d12.SelLength = Len(d12)
End Sub
Private Sub d12_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p5.SetFocus
End Sub
Private Sub d13_GotFocus()
d13.SelStart = 0
d13.SelLength = Len(d13)
End Sub

Private Sub d13_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p14.SetFocus
End Sub
Private Sub d14_GotFocus()
d14.SelStart = 0
d14.SelLength = Len(d14)
End Sub
Private Sub d14_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p15.SetFocus
End Sub
Private Sub d15_GotFocus()
d15.SelStart = 0
d15.SelLength = Len(p15)
End Sub
Private Sub d15_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command1.SetFocus
End Sub
Private Sub d16_GotFocus()
d16.SelStart = 0
d16.SelLength = Len(d16)
End Sub
Private Sub d16_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command1.SetFocus
End Sub
Private Sub d2_GotFocus()
d2.SelStart = 0
d2.SelLength = Len(d2)
End Sub
Private Sub d2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p11.SetFocus
End Sub

Private Sub d3_GotFocus()
d3.SelStart = 0
d3.SelLength = Len(d3)
End Sub
Private Sub d3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p13.SetFocus
End Sub
Private Sub d4_GotFocus()
d4.SelStart = 0
d4.SelLength = Len(d4)
End Sub
Private Sub d4_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p3.SetFocus
End Sub
Private Sub d5_GotFocus()
d5.SelStart = 0
d5.SelLength = Len(d5)
End Sub
Private Sub d5_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p6.SetFocus
End Sub
Private Sub d6_GotFocus()
d6.SelStart = 0
d6.SelLength = Len(d6)
End Sub
Private Sub d6_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p7.SetFocus
End Sub
Private Sub d7_GotFocus()
d7.SelStart = 0
d7.SelLength = Len(d7)
End Sub
Private Sub d7_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p10.SetFocus
End Sub






Private Sub sai14_Click()
Form1.Visible = True
Form14.Visible = False

Form1.Refresh
End Sub

Private Sub text1_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p1 = Val(p1) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub

Private Sub text10_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p4 = Val(p4) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub

Private Sub text13_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p5 = Val(p5) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub
Private Sub text16_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p6 = Val(p6) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub
Private Sub Text19_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p7 = Val(p7) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub


Private Sub Text28_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p10 = Val(p10) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub

Private Sub Text31_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p11 = Val(p11) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub

Private Sub Text34_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p12 = Val(p12) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub
Private Sub Text37_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p13 = Val(p13) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub

Private Sub text4_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p2 = Val(p2) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub
Private Sub Text40_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p14 = Val(p14) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub
Private Sub Text43_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p15 = Val(p15) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub

Private Sub text7_Click()
If resto = Empty Then Exit Sub
If resto <> Empty Then
p3 = Val(p3) + (pdivide)
resto = resto - 1
End If
If resto = 0 Then
pdivide = 0
Call Command1_Click
End If
End Sub

