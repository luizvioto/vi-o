VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Dimensionamento"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11880
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MouseIcon       =   "Form2.frx":030A
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Aguarde 
      Height          =   150
      Left            =   120
      TabIndex        =   89
      Top             =   8000
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000A&
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11655
      Begin VB.Frame Frame2 
         BackColor       =   &H00400000&
         Caption         =   "Fuselagem e Flap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   240
         TabIndex        =   70
         Top             =   6360
         Width           =   11175
         Begin VB.TextBox anguflap 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox volcab 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   10080
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox bflap 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   84
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox cmflap 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox tipoflap 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox hfs 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox wfs 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox lfs 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ang. flap :                  º"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   9360
            TabIndex        =   88
            Top             =   840
            Width           =   1590
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Volume :                   m3"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   9360
            TabIndex        =   85
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Envergadura :                        m."
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   6960
            TabIndex        =   81
            Top             =   840
            Width           =   2250
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Corda média :                            m."
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   4440
            TabIndex        =   80
            Top             =   840
            Width           =   2400
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Flap :"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   79
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Flap "
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
            Left            =   120
            TabIndex        =   78
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Altura máxima :                       m."
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   6960
            TabIndex        =   77
            Top             =   360
            Width           =   2265
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Largura máxima :                       m."
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   4440
            TabIndex        =   76
            Top             =   360
            Width           =   2400
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Comprimento total :                       m."
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   1800
            TabIndex        =   75
            Top             =   360
            Width           =   2550
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Fuselagem"
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
            Left            =   120
            TabIndex        =   74
            Top             =   360
            Width           =   1605
         End
      End
      Begin VB.TextBox perfasa 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox perfvt 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox svt 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox bvt 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox crvt 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox cmvt 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox cpvt 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox espesvt 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox slemdir 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox blemdir 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox crlemdir 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   5400
         Width           =   735
      End
      Begin VB.TextBox cplemdir 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   5880
         Width           =   735
      End
      Begin VB.TextBox cpprofund 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   5880
         Width           =   735
      End
      Begin VB.TextBox crprofund 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   5400
         Width           =   735
      End
      Begin VB.TextBox bprofund 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox sprofund 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox espesht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox cpht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox cmht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox crht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox bht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox sht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox perfht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox cpail 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   5880
         Width           =   735
      End
      Begin VB.TextBox crail 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   5400
         Width           =   735
      End
      Begin VB.TextBox bail 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox sail 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox espessura 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox cpw 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox cmw 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox crw 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox bw 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox sw 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   2040
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   7560
         X2              =   7560
         Y1              =   840
         Y2              =   6240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   3480
         X2              =   3480
         Y1              =   840
         Y2              =   6240
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área da empenagem vertical :                             m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   67
         Top             =   2040
         Width           =   3630
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Envergadura - empenag. vertic. :                         m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   66
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da raiz - empenag. vertic. :                         m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   65
         Top             =   3000
         Width           =   3630
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda média - empenag. vertic. :                          m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   64
         Top             =   3480
         Width           =   3630
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da ponta - empenag. vertic. :                     m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   63
         Top             =   3960
         Width           =   3615
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espessura máx. - empenag. vertic. :                    %"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   62
         Top             =   1560
         Width           =   3540
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área - leme de direção :                                       m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   61
         Top             =   4440
         Width           =   3645
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Envergadura - leme de direção :                          m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   60
         Top             =   4920
         Width           =   3600
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da raiz - leme de direção :                          m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   59
         Top             =   5400
         Width           =   3615
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da ponta - leme de direção :                      m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   58
         Top             =   5880
         Width           =   3600
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil utilizado :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7680
         TabIndex        =   57
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Estabilizador Vertical"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   7680
         TabIndex        =   45
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil utilizado :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   33
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Estabilizador Horizontal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   3570
         TabIndex        =   32
         Top             =   360
         Width           =   3795
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da ponta - profundor :                                m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   31
         Top             =   5880
         Width           =   3600
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da raiz - profundor :                                    m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   30
         Top             =   5400
         Width           =   3615
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Envergadura - profundor :                                    m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   29
         Top             =   4920
         Width           =   3600
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área - profundor :                                                 m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   28
         Top             =   4440
         Width           =   3645
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espessura máx. - empenag. horiz. :                     %"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   27
         Top             =   1560
         Width           =   3525
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil utilizado :"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Asa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Envergadura - aileron                           m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   4920
         Width           =   2895
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da ponta - aileron :                     m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   5880
         Width           =   2895
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da raiz - aileron :                         m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   5400
         Width           =   2910
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área - aileron :                                     m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espessura máxima :                             %"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   2820
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da ponta - empenag. horiz. :                      m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   15
         Top             =   3960
         Width           =   3600
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda média - empenag. horiz. :                           m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   14
         Top             =   3480
         Width           =   3615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da raiz - empenag. horiz. :                          m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   13
         Top             =   3000
         Width           =   3615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Envergadura - empenag. horiz. :                          m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   12
         Top             =   2520
         Width           =   3600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área da empenagem horizontal :                          m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   11
         Top             =   2040
         Width           =   3660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da ponta - Asa :                         m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   3960
         Width           =   2880
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda média - Asa :                              m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda da raiz - Asa :                             m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Envergadura - Asa :                             m."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   2880
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área - Asa :                                          m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   2925
      End
   End
   Begin VB.Menu fileirf2 
      Caption         =   "&Ir para ..."
      Begin VB.Menu fileaerodin 
         Caption         =   "Aerodinâmica"
      End
      Begin VB.Menu fileprincipia 
         Caption         =   "Página inicial"
      End
      Begin VB.Menu filedesenho 
         Caption         =   "Retornar ao desenho"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu fileimprimi2 
      Caption         =   "Im&primir "
   End
   Begin VB.Menu filefer2 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "&Calculadora && Conversor"
      End
   End
   Begin VB.Menu sai2 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub fileaerodin_Click()
'Numero de nervuras Form18
Form18.nnerv = (Form2.bw * 0.066)
Form18.nnerv = Format(Form18.nnerv, "0.00")

'>>>>>>>>>>>> CONFIGURAÇÃO DA ASA  - FORM18 <<<<<<<<<<<<<<<<<
 
'Enviando a eficiencia do Estabilizador

If Form4.m > 0 Then testehtail = (Form4.m * (Form2.bw / 2))
If Form4.m = 0 Then testehtail = (Form4.m * (Form2.bw / 2))
If Form4.m < 0 And Form1.posimedia.Value = True Then testehtail = (Form4.m * (Form2.bw / 3))
If Form4.m < 0 And Form1.posibaixa.Value = True Then testehtail = (Form4.m * (Form2.bw))
If Form4.m < 0 And Form1.posialta.Value = True Then testehtail = (Form4.m * (Form2.bw))


If testehtail > Form2.bvt Then
MsgBox "Diminua a eficiencia do estabilizador ou aumente o alongamento da empenagem vertical", vbInformation
Form1.Visible = True
Form4.Visible = True
Form4.Command8.SetFocus
Exit Sub
End If

'Se o teste for ok então segue tudo como abaixo:
                        If testehtail < Form2.bvt Then
                        
Form10.htail = testehtail
Form10.htail = Format(Form10.htail, "0.00")

'********************* VARIAVEIS COMUNS ***************************
'POTENCIA + RANKINE + CFuse
            hp = Form1.motor
            trank = 519
            kfuse = 2.5
                        
'Opção quanto ao acabamento
If Requisitos.madeira.Value = True Then fatpaint = 0.00046
If Requisitos.composto.Value = True Then fatpaint = 0.0002
If Requisitos.tubo.Value = True Then fatpaint = 0.00048
If Requisitos.metal.Value = True Then fatpaint = 0.00035
If Requisitos.Tubalum.Value = True Then fatpaint = 0.00048

'COLOCAR Banco Dados >>>> fatpaint = 0.00035

                'fuselagem
reyf7 = 38.21 * (Form2.lfs * 3.28 / fatpaint) ^ 1.053
reyf7 = Format(reyf7, "0.00")
lfsdia = (Form2.lfs * 3.28) / (((Form2.wfs * 3.28) * (Form2.hfs * 3.28)) ^ 0.5)
gf7 = Log(reyf7) / Log(10)

                'asa
reyw7 = 38.21 * (((Form2.cmw * 3.28) / fatpaint) ^ 1.053)
reyw7 = Format(reyw7, "0.00")
gw7 = Log(reyw7) / Log(10)

                'empenagem horizontal
reyht7 = 38.21 * (Form2.cmht * (3.28) / fatpaint) ^ 1.053
reyht7 = Format(reyht7, "0.00")
ght7 = Log(reyht7) / Log(10)

                'empenagem vertical
reyvt7 = 38.21 * (Form2.cmvt * (3.28) / fatpaint) ^ 1.053
reyvt7 = Format(reyvt7, "0.00")
gvt7 = Log(reyvt7) / Log(10)

'Arrasto trem de pouso etc

If Form1.Checkretratil.Value = 1 Then cdotp = 0
If Form1.Checkretratil.Value = 0 Then

afpneu = Abs((Form1.diambeq * Form1.espbeq) * 0.001076) + Abs(((Form1.diamtp * Form1.esptp) * 0.001076) * 2)
dqpneu = afpneu * 0.25

cdotp = 1.2 * (dqpneu) / (Form2.Sw * 10.76)
cdotp = Format(cdotp, "0.0000")

End If

' Arrasto do flap
If Form4.semflap.Value = True Then
cdoflap = 0
cdoflap = Format(cdoflap, "0.000")
End If

If Form4.fowlerflap.Value = True Or Form4.plainflap.Value = True Or Form4.splitflap.Value = True Then
areaflap = (((((Form2.cmflap * Sin(Form4.deflap * 3.1415 / 180)) * Form2.bflap) * 10.76) * 2) * 0.1)
cdoflap = 1.2 * (areaflap) / (Form2.Sw * 10.76)
cdoflap = Format(cdoflap, "0.000")
End If

'Arrasto de áreas adicionadas suplementares

If Form1.adicionarea = 0 Then adicionarea = "0.00"

If Form1.adicionarea <> 0 Then
dqoutras = (Form1.adicionarea * 10.76) * 0.1
cdosuple = 1.2 * (dqoutras) / (Form2.Sw * 10.76)
cdosuple = Format(cdosuple, "0.000")
End If
                                                  
'.......................ARRASTO...................................
'COLOCAR Banco Dados >>>> Requisitos.NIcruz = 0.8

kvelok = (Int((137 * (((Form1.motor * Requisitos.NIcruz) / (Form2.Sw * 10.76)) ^ (1 / 3))) * 1.61)) * 2.5
kvelok = Format(kvelok, "0.00")



Aguarde.Visible = True
Aguarde.Max = kvelok

'Precisão da velocidade -  step e adição em Vstall


'If Form4.km5.Value = True Then
STP = 0.5
UNI = 0.5
'End If

'If Form4.km10.Value = True Then
'STP = 0.1
'UNI = 0.1
'End If

'VELOCIDADES
'kvelok é dado em km/h
For velok = Int(Form1.vstol) To Int(kvelok) Step STP

'barra de rolagem
Aguarde.Value = velok
    
'km/h para mph
velokmph = velok / 1.61

'km/h para ft/s
velokpe = velok * 0.911

'Em Mach
velokmach = velok / 1224
velokmach = Format(velokmach, "0.000")

'Area Molhada
'29/02/2004
'12/12/2007

If Form20.p1 = "PCA" Or Form20.p1 = "PCA I" Or Form20.p1 = "PCA II" Then swetf7 = (6.283 * ((0.3819 * Form2.lfs + 0.12795 * Form2.lfs ^ 2 - 0.0091 * Form2.lfs ^ 3)))
If Form20.p1 <> "PCA" Then
    If Form2.lfs < 8.5 Then swetf7 = 2 * (6.283 * ((0.1477 * Form2.lfs + 0.08275 * Form2.lfs ^ 2 - 0.007766 * Form2.lfs ^ 3)))
    If Form2.lfs > 8.5 Then swetf7 = 2 * (6.283 * ((0.2288 * Form2.lfs + 0.0397 * Form2.lfs ^ 2 - 0.003133 * Form2.lfs ^ 3)))
    If Form2.lfs = 8.5 Then swetf7 = 2 * (6.283 * ((0.2288 * Form2.lfs + 0.0397 * Form2.lfs ^ 2 - 0.003133 * Form2.lfs ^ 3)))
End If
'antigo swetf7 = kfuse * (((Form2.wfs * 3.28) * (Form2.hfs * 3.28)) ^ 0.5) * (Form2.lfs * 3.28)
If Form1.OP1.Value = False Then
    Form7.swetfuse = swetf7
    Form7.swetfuse = Format(Form7.swetfuse, "0.000")
End If

If Form1.OP1.Value = True Then
    Form7.swetfuse = 90
    Form7.swetfuse = Format(Form7.swetfuse, "0.000")
End If

swetw7 = ((Form2.Sw * 10.76) * (1.977 + (0.52 * (Form2.espessura / 100)))) * 0.835
Form7.swetasa = swetw7 / 10.76
Form7.swetasa = Format(Form7.swetasa, "0.000")


swetht7 = ((Form2.sht * 10.76) * (1.977 + (0.52 * (Form2.espesht / 100)))) * 0.8
Form7.swethorz = swetht7 / 10.76
Form7.swethorz = Format(Form7.swethorz, "0.000")


swetvt7 = ((Form2.svt * 10.76) * (1.977 + (0.52 * (Form2.espesvt / 100)))) * 0.8
Form7.swetvert = swetvt7 / 10.76
Form7.swetvert = Format(Form7.swetvert, "0.000")


'Arrasto fuselagem
cff7 = 0.455 / ((gf7 ^ 2.58) * (1 + (0.144 * (velokmach)) ^ 2) ^ 0.65)
cff7 = Format(cff7, "0.000")

fff7 = 1 + (60 / lfsdia ^ 3) + (lfsdia / 400)
fff7 = Format(fff7, "0.000")

cdof7 = (cff7 * fff7 * swetf7) / (Form2.Sw * 10.76)
cdof7 = Format(cdof7, "0.000")

'Arrasto asa

cfw7 = 0.455 / ((gw7 ^ 2.58) * (1 + (0.144 * (velokmach)) ^ 2) ^ 0.65)
cfw7 = Format(cfw7, "0.000")

ffw7 = ((1 + ((0.6 / 0.25) * (Form2.espessura / 100)) + (100 * ((Form2.espessura / 100) ^ 4))) * (1.34 * ((velokmach) ^ 0.18)))
ffw7 = Format(ffw7, "0.00")
'Cdo stall asa
cdow7 = (cfw7 * ffw7 * swetw7) / (Form2.Sw * 10.76)
cdow7 = Format(cdow7, "0.000")

'Arrasto da empenagem horizontal

cfht7 = 0.455 / ((ght7 ^ 2.58) * (1 + (0.144 * (velokmach)) ^ 2) ^ 0.65)
cfht7 = Format(cfht7, "0.000")

ffht7 = ((1 + ((0.6 / 0.25) * (Form2.espesht / 100)) + (100 * ((Form2.espesht / 100) ^ 4))) * (1.34 * ((velokmach) ^ 0.18)))
ffht7 = Format(ffht7, "0.000")
'Cdo stall emp horz
cdoht7 = ((cfht7 * ffht7 * swetht7) / (Form2.Sw * 10.76))
cdoht7 = Format(cdoht7, "0.000")

'Arrasto da empenagem vertical
cfvt7 = 0.455 / ((gvt7 ^ 2.58) * (1 + (0.144 * (velokmach)) ^ 2) ^ 0.65)
cfvt7 = Format(cfvt7, "0.000")

ffvt7 = ((1 + ((0.6 / 0.25) * (Form2.espesvt / 100)) + (100 * ((Form2.espesvt / 100) ^ 4))) * (1.34 * ((velokmach) ^ 0.18)))
ffvt7 = Format(ffvt7, "0.00")
'Cdo stall emp vert
cdovt7 = ((cfvt7 * ffvt7 * swetvt7) / (Form2.Sw * 10.76))
cdovt7 = Format(cdoht7, "0.0000")

'Arrasto Cooling e motor
If velok < 100 Then cdocool = 0
If velk < 100 Then cdomotor = 0
If velok >= 100 Then
dqcool = ((4.9 * 10 ^ -7) * Form1.motor * (trank ^ 2)) / (velokpe)
dqmotor = ((2 * 10 ^ -4) * Form1.motor)

cdocool = ((2.5) * dqcool) / (Form2.Sw * 10.76)
cdocool = Format(cdocool, "0.000")

cdomotor = dqmotor / (Form2.Sw * 10.76)
cdomotor = Format(cdomotor, "0.000")
End If

'ARRASTO PARASITA TOTAL
cdif = 1.1
'Com arrasto do Perfil
'If Abs(cdototal7 * Form2.Sw * 10.76 * qd7) - Abs(cdi7 * Form2.Sw * 10.76 * qd7) > 0.000001 Then
If velok > 1.1 * Form1.vstol Then cdoflap = 0

cdototal7 = cdif * ((Abs(cdof7) + Abs(cdow7) + Abs(Form7.cdperfasa) + Abs(cdoht7) + Abs(cdovt7) + Abs(cdotp)) + Abs(cdocool) + Abs(cdomotor) + Abs(cdosuple) + Abs(cdoflap))
cdototal7 = Format(cdototal7, "0.0000")

'Fim de arrasto PARASITA ideal

'ARRASTO INDUZIDO
dar = 1.225 * 0.00194

e = 1.78 * (1 - (0.045 * Form1.arw ^ 0.68)) - 0.64
e = Format(e, "0.0000")

k = 1 / (3.1415 * Form1.arw * e)
k = Format(k, "0.0000")

qd7 = (dar * (velokpe ^ 2)) / 2
qd7 = Format(qd7, "0.0000")

cl7 = (Form1.wo * 2.2) / ((Form2.Sw * 10.76) * qd7)
cl7 = Format(cl7, "0.0000")

cdi7 = (Abs(k) * (cl7 ^ 2))
cdi7 = Format(cdi7, "0.0000")

'ARRASTO TOTAL
cd7 = Abs(cdototal7) + Abs(cdi7)
cd7 = Format(cd7, "0.0000")
      
            
'.............  VELOCIDADE DE PLANEIO (Subida)  .........................


            If Abs(cdototal7 * Form2.Sw * 10.76 * qd7) - Abs(cdi7 * Form2.Sw * 10.76 * qd7) < 0.000001 Then
            
            Form7.clk2 = cl7
            Form7.clk2 = Format(Form7.clk2, "0.000")
            
            Form7.cdtotsubida = cd7
            Form7.cdtotsubida = Format(Form7.cdtotsubida, "0.000")
            
            Form7.ldk2 = (cl7 / cd7)
            Form7.ldk2 = Format(Form7.ldk2, "0.000")
                                
            Form8.Fdsubida = (cd7 * Form2.Sw * 10.76 * qd7) / 2.2
            Form8.Fdsubida = Format(Form8.Fdsubida, "0.000")
            
            'Velocidade planeio ideal
            'videal = velok
            'videal = Format(videal, "0.00")
                        
            Form8.vsubida = velok
            Form8.vsubida = Format(Form8.vsubida, "0.00")
            
            End If
            
'AQUI ................. VELOCIDADE DE SUBIDA .................
T = ((550 * Form1.motor * Requisitos.NIcruz) / velokpe)
T = Format(T, "0.00")

d = Abs((cdi7 * Form2.Sw * 10.76 * qd7) / 2.2) + Abs((cd7 * Form2.Sw * 10.76 * qd7) / 2.2)
d = Format(d, "0.00")

vv = (velokpe * (Abs(T) - Abs(d))) / Form1.wo
vv = Format(vv, "0.00")


 '................. VELOCIDADE MÁXIMA ....................
forcarrasto = Abs(cd7 * Form2.Sw * 10.76 * qd7)
forcarrasto = Format(forcarrasto, "0.000")
 
tracmotor = ((550 * Form1.motor * Requisitos.NIcruz) / velokpe)
tracmotor = Format(tracmotor, "0.000")

            If Abs(forcarrasto) < 0.999999 * Abs(tracmotor) Then
            
            'Form8.FD = Abs((cdi7 * Form2.sw * 10.76 * qd7) / 2.2) + Abs((cd7 * Form2.sw * 10.76 * qd7) / 2.2)
            Form8.FD = forcarrasto / 2.2
            Form8.FD = Format(Form8.FD, "0.000")
            
            Form7.clk4 = cl7
            Form7.clk4 = Format(Form7.clk4, "0.000")
            
            Form7.cdtotvmax = cd7
            Form7.cdtotvmax = Format(Form7.cdtotvmax, "0.000")
            
            Form7.ldk4 = Form7.clk4 / Form7.cdtotvmax
            Form7.ldk4 = Format(Form7.ldk4, "0.000")
            
            Form7.placaplan = cd7 * Form2.Sw
            Form7.placaplan = Format(Form7.placaplan, "0.000")
            
            Form9.velomaxima = velok
            Form9.velomaxima = Format(Form9.velomaxima, "0.000")
            
            Form8.vmax = velok
            Form8.vmax = Format(Form8.vmax, "0.00")
            
            End If

            
'.................... VELOCIDADE DE STALL .....................

            If velok < 1.01 * Form1.vstol Then
            Form7.clk1 = cl7
            Form7.clk1 = Format(Form7.clk1, "0.000")
            
            Form7.cdtotstall = cd7
            Form7.cdtotstall = Format(Form7.cdtotstall, "0.000")
            
            Form7.ldk1 = cl7 / cd7
            Form7.ldk1 = Format(Form7.ldk1, "0.000")
            
            Form8.FDstol = Abs((cdi7 * Form2.Sw * 10.76 * qd7) / 2.2) + Abs((cd7 * Form2.Sw * 10.76 * qd7) / 2.2)
            Form8.FDstol = Format(Form8.FDstol, "0.000")
            
            Form8.vstol = velok
            Form8.vstol = Format(Form8.vstol, "0.00")
            End If
            
'................... VELOCIDADE DE DECOLAGEM ..................

            If velok <= 1.1 * Form1.vstol Then
            Form8.vdecol = velok
            Form8.vdecol = Format(Form8.vdecol, "0.00")
            
            Form7.cldecol = cl7
            Form7.cldecol = Format(Form7.cldecol, "0.000")
            
            Form7.cddecol = cd7
            Form7.cddecol = Format(Form7.cddecol, "0.000")
            
            Form7.lddecol = cl7 / cd7
            Form7.lddecol = Format(Form7.lddecol, "0.000")
            
            Form8.FDdecol = Abs((cdi7 * Form2.Sw * 10.76 * qd7) / 2.2) + Abs((cd7 * Form2.Sw * 10.76 * qd7) / 2.2)
            Form8.FDdecol = Format(Form8.FDdecol, "0.00")
            
            Form8.vdecol = velok
            Form8.vdecol = Format(Form8.vdecol, "0.00")
            End If
            
             If Abs(forcarrasto) >= Abs(tracmotor) Then
             
                Form8.FD = forcarrasto / 2.2
                Form8.FD = Format(Form8.FD, "0.00")
            
                Form7.clk4 = cl7
                Form7.clk4 = Format(Form7.clk4, "0.000")
            
                Form7.cdtotvmax = cd7
                Form7.cdtotvmax = Format(Form7.cdtotvmax, "0.000")
            
                Form7.ldk4 = Form7.clk4 / Form7.cdtotvmax
                Form7.ldk4 = Format(Form7.ldk4, "0.000")
            
                Form7.placaplan = cd7 * Form2.Sw
                Form7.placaplan = Format(Form7.placaplan, "0.000")
            
                Form9.velomaxima = velok
                Form9.velomaxima = Format(Form9.velomaxima, "0.000")
            
                Form8.vmax = velok
                Form8.vmax = Format(Form8.vmax, "0.00")
                
                GoTo 1000
                
            End If
                        
Next


1000:
'......................... OUTROS .............................

'Angulo de planeio
Form7.anplan = Atn(1 / Form7.ldk2) * (180 / 3.141516)
Form7.anplan = Format(Form7.anplan, "0.000")



'Slope Estab. Horiz.
deltanguloht = Abs(Form7.alfahtmax) + Abs(Form7.alfahtzero)

Form7.slopeht = (Form1.clmaxht / deltanguloht) * 57.3
Form7.slopeht = Format(Form7.slopeht, "0.00")
clmax3d = ((Form7.slopeht / (1 + (Abs(Form7.slopeht / (3.141516 * Form1.arht))))) / 57.3) * deltanguloht
Form7.cl3dht = clmax3d
Form7.cl3dht = Format(Form7.cl3dht, "0.000")

'Correção para Cl Stall
If Form7.clk1 > Form7.cl3dFLAP Then Form7.clk1 = Form7.cl3dFLAP

'Comparando Vsubida e Vstall
'If Form1.vstol > videal Then
    'MsgBox "A velocidade de subida está abaixo da velocidade de estol", vbCritical
    'Form2.Visible = False
    'Form1.Visible = True
    'Form2.Aguarde.Visible = False
    'Form1.Refresh
        'Exit Sub
'End If

'carga alar
Form7.wsproj = Form1.wo / Form2.Sw
Form7.wsproj = Format(Form7.wsproj, "0.000")

Form7.perfasa = Form4.perfil1
Form7.perfesth = Form4.perfil2
Form7.perfvh = Form4.perfil3

Form7.Visible = True
Form2.Visible = False
Aguarde.Visible = False

                    End If
                    
                    
                                       
End Sub

Private Sub filecalcon_Click()
Form29.Show
End Sub

Private Sub filedesenho_Click()
Form18.Visible = True
Form2.Visible = False
Form18.Refresh
End Sub

Private Sub fileimprimi2_Click()
If LIBER <> "KFAR0852" Then
    MsgBox "A rotina de impressão sómente estará disponível após a liberação permanente", 16, "Cancelamento de impressão": Exit Sub
End If

'Pagina 2
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print ""
Printer.Print Tab(25); "I - DIMENSIONAMENTO"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "1 - ASA"
Printer.FontSize = 10
Printer.Print Tab(20); "Aerofólio utilizado:"; Tab(70); Form2.perfasa
Printer.Print Tab(20); "Espessura do aerofólio da asa:"; Tab(70); Form2.espessura; " %"
Printer.Print Tab(20); "Área da Asa:"; Tab(70); Form2.Sw; " m2"
Printer.Print Tab(20); "Envergadura da Asa:"; Tab(70); Form2.bw; " m"
Printer.Print Tab(20); "Corda da Raiz da Asa:"; Tab(70); Form2.crw; " m"
Printer.Print Tab(20); "Corda Média da Asa:"; Tab(70); Form2.cmw; " m"
Printer.Print Tab(20); "Corda da Ponta da Asa:"; Tab(70); Form2.cpw; " m"
Printer.Print Tab(20); "Envergadura do Aileron:"; Tab(70); Form2.bail; " m"
Printer.Print Tab(20); "Corda da Raiz do Aileron:"; Tab(70); Form2.crail; " m"
Printer.Print Tab(20); "Corda da Ponta do Aileron:"; Tab(70); Form2.cpail; " m"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "2 - EMPENAGEM HORIZONTAL"
Printer.FontSize = 10
Printer.Print Tab(20); "Aerofólio utilizado:"; Tab(70); Form2.perfht
Printer.Print Tab(20); "Espessura do aerofólio:"; Tab(70); Form2.espesht; " %"
Printer.Print Tab(20); "Área da Emp.Horizontal:"; Tab(70); Form2.sht; " m2"
Printer.Print Tab(20); "Envergadurada Emp.Horizontal:"; Tab(70); Form2.bht; " m"
Printer.Print Tab(20); "Corda da Raiz da Emp. Horizontal:"; Tab(70); Form2.crht; " m"
Printer.Print Tab(20); "Corda Média da Emp. Horizontal:"; Tab(70); Form2.cmht; " m"
Printer.Print Tab(20); "Corda da Ponta da Emp. Horizontal:"; Tab(70); Form2.cpht; " m"
Printer.Print Tab(20); "Área do Profundor:"; Tab(70); Form2.sprofund; " m2"
Printer.Print Tab(20); "Envergadura do Profundor:"; Tab(70); Form2.bprofund; " m"
Printer.Print Tab(20); "Corda da Raiz do Profundor:"; ; Tab(70); Form2.crprofund; " m"
Printer.Print Tab(20); "Corda da Ponta do Profundor:"; Tab(70); Form2.cpprofund; " m"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "3 - EMPENAGEM VERTICAL"
Printer.FontSize = 10
Printer.Print Tab(20); "Aerofólio utilizado:"; Tab(70); Form2.perfvt
Printer.Print Tab(20); "Espessura do aerofólio:"; Tab(70); Form2.espesvt; " %"
Printer.Print Tab(20); "Área da Emp. Vertical:"; Tab(70); Form2.svt; " m2"
Printer.Print Tab(20); "Envegadura da Emp. Vertical:"; Tab(70); Form2.bvt; " m"
Printer.Print Tab(20); "Corda da Raiz da Emp. Vertical:"; Tab(70); Form2.crvt; " m"
Printer.Print Tab(20); "Corda Média da Emp. Vertical:"; Tab(70); Form2.cmvt; " m"
Printer.Print Tab(20); "Corda da Ponta da Emp. Vertical:"; Tab(70); Form2.cpvt; " m"
Printer.Print Tab(20); "Área do Leme de direção:"; Tab(70); Form2.slemdir; " m2"
Printer.Print Tab(20); "Envergadura do Leme de direção:"; Tab(70); Form2.blemdir; " m2"
Printer.Print Tab(20); "Corda da Raiz do Leme de direção:"; Tab(70); Form2.crlemdir; " m"
Printer.Print Tab(20); "Corda da Ponta do Leme de direção:"; Tab(70); Form2.cplemdir; " m"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print ""
Printer.Print Tab(25); "4 - FUSELAGEM"
Printer.FontSize = 10
Printer.Print Tab(20); "Comprimento da Fuselagem:"; Tab(70); Form2.lfs; " m"
Printer.Print Tab(20); "Largura da Fuselagem:"; Tab(70); Form2.wfs; " m"
Printer.Print Tab(20); "Altura da Fuselagem:"; Tab(70); Form2.hfs; " m"
Printer.Print Tab(20); "Volume da Fuselagem:"; Tab(70); Form2.volcab; " m3"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print ""
Printer.Print Tab(25); "5 - FLAP"
Printer.FontSize = 10
Printer.Print Tab(20); "Tipo de Flap:"; Tab(70); Form2.tipoflap
Printer.Print Tab(20); "Envegadura do Flap:"; Tab(70); Form2.bflap; " m"
Printer.Print Tab(20); "Corda Média do Flap:"; Tab(70); Form2.cmflap; " m"
Printer.Print ""
Printer.Print Tab(20); "Fim da página";
Printer.EndDoc

End Sub

Private Sub fileprincipia_Click()
Form1.Visible = True
Form2.Visible = False
Form1.Refresh
End Sub

Private Sub sai2_Click()
Form1.Visible = True
Form2.Visible = False
Form1.Refresh
End Sub
