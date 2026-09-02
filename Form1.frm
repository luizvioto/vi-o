VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Principal"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11880
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Aguarde1 
      Height          =   150
      Left            =   120
      TabIndex        =   142
      Top             =   8000
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox Picture5 
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1875
      ScaleWidth      =   11595
      TabIndex        =   4
      Top             =   1680
      Width           =   11655
      Begin VB.Frame Frame7 
         Height          =   1815
         Left            =   9960
         TabIndex        =   86
         Top             =   0
         Width           =   1545
         Begin VB.OptionButton tremconv 
            Caption         =   "&Convencional"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton tremtric 
            Caption         =   "&Triciclo"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox Checkretratil 
            Caption         =   "&Retrátil ?"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   1440
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1815
         Left            =   8280
         TabIndex        =   65
         Top             =   0
         Width           =   1705
         Begin VB.TextBox clmaxFLAP 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox clmaxht 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   1400
            Width           =   615
         End
         Begin VB.TextBox clmaximo 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CL c/flap"
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "CL est.hor."
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   1395
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "CL s/flap"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   6360
         TabIndex        =   55
         Top             =   0
         Width           =   2055
         Begin VB.TextBox afilavt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1080
            TabIndex        =   58
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox afilaht 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1080
            TabIndex        =   57
            Text            =   "0"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox afilasa 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1080
            TabIndex        =   56
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "l"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   14.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   240
            TabIndex        =   64
            Top             =   120
            Width           =   165
         End
         Begin VB.Label Label28 
            Caption         =   "l"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   11.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   535
            Width           =   135
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Vertical"
            Height          =   195
            Left            =   240
            TabIndex        =   62
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Horizontal"
            Height          =   195
            Left            =   240
            TabIndex        =   61
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "empenagens"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   480
            TabIndex        =   60
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "da asa"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   480
            TabIndex        =   59
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1815
         Left            =   4440
         TabIndex        =   47
         Top             =   0
         Width           =   2055
         Begin VB.TextBox arht 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1080
            TabIndex        =   49
            Text            =   "0"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox arvt 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1080
            TabIndex        =   50
            Text            =   "0"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox arw 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1080
            TabIndex        =   48
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Horizontal"
            Height          =   195
            Left            =   240
            TabIndex        =   54
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Vertical"
            Height          =   195
            Left            =   240
            TabIndex        =   53
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "AR  empenagens"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   52
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "AR da asa"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   240
            Width           =   750
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1815
         Left            =   3000
         TabIndex        =   113
         Top             =   0
         Width           =   1575
         Begin VB.OptionButton posialta 
            Caption         =   "&Alta"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton posimedia 
            Caption         =   "&Média"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton posibaixa 
            Caption         =   "&Baixa"
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   1440
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   30
         TabIndex        =   5
         Top             =   0
         Width           =   2970
         Begin VB.PictureBox Picture12 
            Height          =   305
            Left            =   2040
            ScaleHeight     =   240
            ScaleWidth      =   675
            TabIndex        =   129
            Top             =   1440
            Width           =   735
            Begin VB.HScrollBar HScroll1 
               Height          =   255
               Left            =   0
               Max             =   10
               TabIndex        =   130
               Top             =   0
               Width           =   680
            End
         End
         Begin VB.TextBox altitude 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2040
            TabIndex        =   127
            Text            =   "0"
            Top             =   1040
            Width           =   735
         End
         Begin VB.CheckBox alterapeso 
            Caption         =   "alterar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox roar 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2040
            TabIndex        =   124
            Text            =   "1,225"
            Top             =   400
            Width           =   735
         End
         Begin VB.TextBox pas 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1140
            TabIndex        =   122
            Text            =   "2"
            Top             =   400
            Width           =   735
         End
         Begin VB.PictureBox Picture8 
            Height          =   305
            Left            =   1140
            ScaleHeight     =   240
            ScaleWidth      =   675
            TabIndex        =   119
            Top             =   1440
            Width           =   735
            Begin VB.HScrollBar rotac 
               Height          =   255
               Left            =   0
               Max             =   4000
               Min             =   1
               SmallChange     =   100
               TabIndex        =   120
               Top             =   0
               Value           =   2800
               Width           =   680
            End
         End
         Begin VB.TextBox RPM 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   118
            Text            =   "2800"
            Top             =   1040
            Width           =   735
         End
         Begin VB.TextBox motor 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   400
            Width           =   735
         End
         Begin VB.TextBox wo 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   1040
            Width           =   735
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Altitude"
            Height          =   195
            Left            =   2160
            TabIndex        =   128
            Top             =   795
            Width           =   525
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "kg/m3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   2280
            TabIndex        =   125
            Top             =   165
            Width           =   420
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "r"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2040
            TabIndex        =   123
            Top             =   165
            Width           =   105
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº  pás"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1200
            TabIndex        =   121
            Top             =   165
            Width           =   540
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "rpm"
            Height          =   195
            Left            =   1200
            TabIndex        =   117
            Top             =   795
            Width           =   255
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motor - hp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   240
            TabIndex        =   10
            Top             =   165
            Width           =   735
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Peso - kg"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   360
            TabIndex        =   9
            Top             =   795
            Width           =   675
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   3150
      ScaleHeight     =   435
      ScaleWidth      =   8565
      TabIndex        =   3
      Top             =   1080
      Width           =   8625
      Begin VB.CommandButton Command11 
         Enabled         =   0   'False
         Height          =   455
         Left            =   6930
         MouseIcon       =   "Form1.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   0
         Width           =   1655
      End
      Begin VB.CommandButton Command14 
         Enabled         =   0   'False
         Height          =   455
         Left            =   5280
         MouseIcon       =   "Form1.frx":0A56
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   0
         Width           =   1670
      End
      Begin VB.CommandButton Command13 
         Enabled         =   0   'False
         Height          =   455
         Left            =   3360
         MouseIcon       =   "Form1.frx":11A2
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":14AC
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton Command12 
         Enabled         =   0   'False
         Height          =   455
         Left            =   1440
         MouseIcon       =   "Form1.frx":18EE
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1BF8
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton Command10 
         Enabled         =   0   'False
         Height          =   455
         Left            =   0
         MouseIcon       =   "Form1.frx":203A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":2344
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   11595
      TabIndex        =   2
      Top             =   3600
      Width           =   11655
      Begin VB.Frame Frame13 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   0
         TabIndex        =   126
         Top             =   0
         Width           =   8490
         Begin VB.Image Image1 
            Height          =   5025
            Left            =   0
            Picture         =   "Form1.frx":2786
            Top             =   -120
            Width           =   9075
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Não obrigatório "
         ForeColor       =   &H000000FF&
         Height          =   4255
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   8295
         Begin VB.Frame Frame17 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Adicionar áreas frontais"
            ForeColor       =   &H00FF0000&
            Height          =   1095
            Left            =   5520
            TabIndex        =   43
            Top             =   305
            Width           =   2655
            Begin VB.TextBox adicionarea 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   1560
               TabIndex        =   44
               Text            =   "0"
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "m2"
               Height          =   195
               Left            =   2280
               TabIndex        =   46
               Top             =   480
               Width           =   210
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Adicionar área de :"
               Height          =   195
               Left            =   120
               TabIndex        =   45
               Top             =   480
               Width           =   1335
            End
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Alterar Volumes de empenagens"
            ForeColor       =   &H00FF0000&
            Height          =   2415
            Left            =   2760
            TabIndex        =   33
            Top             =   305
            Width           =   2655
            Begin VB.TextBox volumeht 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1200
               TabIndex        =   37
               Text            =   "0"
               Top             =   480
               Width           =   615
            End
            Begin VB.HScrollBar hsvolumeht 
               Height          =   255
               Left            =   960
               Max             =   199
               TabIndex        =   36
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox volumevt 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1200
               TabIndex        =   35
               Text            =   "0"
               Top             =   1560
               Width           =   615
            End
            Begin VB.HScrollBar hsvolumevt 
               Height          =   255
               Left            =   960
               Max             =   299
               TabIndex        =   34
               Top             =   1920
               Width           =   1095
            End
            Begin VB.TextBox volempemV 
               Height          =   285
               Left            =   960
               TabIndex        =   76
               Top             =   1920
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.TextBox volempemH 
               Height          =   285
               Left            =   960
               TabIndex        =   75
               Top             =   840
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Horizontal                      %"
               Height          =   195
               Left            =   240
               TabIndex        =   39
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Vertical                          %"
               Height          =   195
               Left            =   240
               TabIndex        =   38
               Top             =   1560
               Width           =   1815
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Alterar Margem estática"
            ForeColor       =   &H00FF0000&
            Height          =   1155
            Left            =   5520
            TabIndex        =   29
            Top             =   1545
            Width           =   2655
            Begin VB.TextBox margestat 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1560
               TabIndex        =   31
               Text            =   "0"
               Top             =   360
               Width           =   615
            End
            Begin VB.HScrollBar hsmargem 
               Height          =   255
               Left            =   1320
               Max             =   50
               Min             =   1
               TabIndex        =   30
               Top             =   720
               Value           =   10
               Width           =   1095
            End
            Begin VB.TextBox marg 
               Height          =   285
               Left            =   1320
               TabIndex        =   78
               Top             =   720
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Margem estática                   %"
               Height          =   195
               Left            =   240
               TabIndex        =   32
               Top             =   360
               Width           =   2145
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Alterar Localização da asa"
            ForeColor       =   &H00FF0000&
            Height          =   1215
            Left            =   2760
            TabIndex        =   25
            Top             =   2815
            Width           =   2655
            Begin VB.TextBox nariz 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1200
               TabIndex        =   27
               Text            =   "0"
               Top             =   480
               Width           =   615
            End
            Begin VB.HScrollBar hsnariz 
               Height          =   255
               Left            =   960
               Max             =   199
               TabIndex        =   26
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox distanasa 
               Height          =   285
               Left            =   960
               TabIndex        =   77
               Top             =   840
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Distancia                       %"
               Height          =   195
               Left            =   240
               TabIndex        =   28
               Top             =   480
               Width           =   1815
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Alterar Fuselagem "
            ForeColor       =   &H00FF0000&
            Height          =   3735
            Left            =   120
            TabIndex        =   14
            Top             =   305
            Width           =   2580
            Begin VB.OptionButton OP1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Entrar c/ Swet fuselagem"
               Height          =   255
               Left            =   240
               TabIndex        =   143
               Top             =   3120
               Width           =   2175
            End
            Begin VB.HScrollBar hsalto 
               Height          =   255
               Left            =   1080
               Max             =   199
               Min             =   1
               TabIndex        =   20
               Top             =   2520
               Value           =   1
               Width           =   1095
            End
            Begin VB.TextBox alto 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   19
               Text            =   "0"
               Top             =   2160
               Width           =   615
            End
            Begin VB.HScrollBar hslargo 
               Height          =   255
               Left            =   1080
               Max             =   199
               Min             =   1
               TabIndex        =   18
               Top             =   1680
               Value           =   1
               Width           =   1095
            End
            Begin VB.TextBox largo 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   17
               Text            =   "0"
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox fuse 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   16
               Text            =   "0"
               Top             =   480
               Width           =   615
            End
            Begin VB.HScrollBar hsfuse 
               Height          =   255
               Left            =   1080
               Max             =   199
               Min             =   1
               TabIndex        =   15
               Top             =   840
               Value           =   1
               Width           =   1095
            End
            Begin VB.TextBox valfuse 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1200
               TabIndex        =   72
               Top             =   840
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.TextBox valfuseL 
               Height          =   285
               Left            =   1200
               TabIndex        =   73
               Top             =   1680
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.TextBox valfuseH 
               Height          =   285
               Left            =   1080
               TabIndex        =   74
               Top             =   2520
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Altura máxima                       %"
               Height          =   195
               Left            =   240
               TabIndex        =   23
               Top             =   2160
               Width           =   2130
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Largura máxima                    %"
               Height          =   195
               Left            =   240
               TabIndex        =   22
               Top             =   1320
               Width           =   2130
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Comprimento                         %"
               Height          =   195
               Left            =   240
               TabIndex        =   21
               Top             =   480
               Width           =   2160
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Alterar Peso"
            ForeColor       =   &H00FF0000&
            Height          =   1215
            Left            =   5520
            TabIndex        =   24
            Top             =   2815
            Width           =   2655
            Begin VB.PictureBox Picture13 
               Height          =   430
               Left            =   260
               ScaleHeight     =   375
               ScaleWidth      =   2055
               TabIndex        =   134
               Top             =   480
               Width           =   2115
               Begin VB.CommandButton Command2 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Alterar  Peso"
                  Height          =   375
                  Left            =   -120
                  MaskColor       =   &H00E0E0E0&
                  Style           =   1  'Graphical
                  TabIndex        =   135
                  Top             =   0
                  UseMaskColor    =   -1  'True
                  Width           =   2175
               End
            End
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         ForeColor       =   &H80000008&
         Height          =   4280
         Left            =   8520
         ScaleHeight     =   4245
         ScaleWidth      =   3030
         TabIndex        =   11
         Top             =   0
         Width           =   3055
         Begin VB.PictureBox Picture10 
            Height          =   1335
            Left            =   75
            ScaleHeight     =   1275
            ScaleWidth      =   2820
            TabIndex        =   91
            Top             =   120
            Width           =   2880
            Begin VB.CommandButton Command8 
               Caption         =   "&Dimensões de rodas"
               Enabled         =   0   'False
               Height          =   1275
               Left            =   0
               MouseIcon       =   "Form1.frx":97030
               MousePointer    =   99  'Custom
               Picture         =   "Form1.frx":9733A
               Style           =   1  'Graphical
               TabIndex        =   92
               ToolTipText     =   "Escolha as dimensôes das rodas"
               Top             =   0
               Width           =   2820
            End
            Begin VB.PictureBox Picture11 
               Height          =   795
               Left            =   2480
               MouseIcon       =   "Form1.frx":98FAC
               MousePointer    =   99  'Custom
               ScaleHeight     =   735
               ScaleWidth      =   225
               TabIndex        =   102
               Top             =   360
               Width           =   285
               Begin VB.VScrollBar VScroll1 
                  Height          =   375
                  Left            =   0
                  Max             =   60
                  TabIndex        =   105
                  Top             =   360
                  Value           =   55
                  Width           =   225
               End
               Begin VB.VScrollBar VScroll2 
                  Height          =   375
                  Left            =   0
                  Max             =   50
                  MouseIcon       =   "Form1.frx":992B6
                  MousePointer    =   99  'Custom
                  TabIndex        =   103
                  Top             =   0
                  Value           =   40
                  Width           =   225
               End
            End
            Begin VB.TextBox espbeq 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   800
               TabIndex        =   98
               Text            =   "20"
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox diambeq 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1400
               TabIndex        =   97
               Text            =   "40"
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox esptp 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   800
               TabIndex        =   96
               Text            =   "25"
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox diamtp 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1400
               TabIndex        =   95
               Text            =   "55"
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox y1 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   2000
               Locked          =   -1  'True
               TabIndex        =   94
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox y2 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   2000
               Locked          =   -1  'True
               TabIndex        =   93
               Top             =   360
               Width           =   375
            End
            Begin VB.CommandButton Command18 
               BackColor       =   &H00E0E0E0&
               Caption         =   "?"
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
               Left            =   2490
               Style           =   1  'Graphical
               TabIndex        =   107
               Top             =   60
               Width           =   255
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "H/D"
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
               Left            =   2040
               TabIndex        =   106
               Top             =   120
               Width           =   300
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "D(cm)"
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
               Left            =   1410
               MouseIcon       =   "Form1.frx":995C0
               MousePointer    =   99  'Custom
               TabIndex        =   104
               ToolTipText     =   "Diametro do pneu"
               Top             =   120
               Width           =   465
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Bequilha"
               Height          =   195
               Left            =   75
               TabIndex        =   101
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Principal"
               Height          =   195
               Left            =   75
               TabIndex        =   100
               Top             =   840
               Width           =   600
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "L(cm)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   180
               Left            =   840
               MouseIcon       =   "Form1.frx":998CA
               MousePointer    =   99  'Custom
               TabIndex        =   99
               ToolTipText     =   "Espessura do pneu (tala)"
               Top             =   120
               Width           =   435
            End
         End
         Begin VB.PictureBox Picture9 
            Height          =   975
            Left            =   75
            ScaleHeight     =   915
            ScaleWidth      =   2820
            TabIndex        =   79
            Top             =   1560
            Width           =   2880
            Begin VB.CommandButton Command15 
               BackColor       =   &H80000000&
               Caption         =   "&Velocidade de stall"
               Enabled         =   0   'False
               Height          =   915
               Left            =   0
               MouseIcon       =   "Form1.frx":99BD4
               MousePointer    =   99  'Custom
               Picture         =   "Form1.frx":99EDE
               Style           =   1  'Graphical
               TabIndex        =   85
               ToolTipText     =   "Clique aqui para calcular a velocidade de estol"
               Top             =   0
               Width           =   2820
            End
            Begin VB.TextBox vstol 
               Alignment       =   2  'Center
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1560
               TabIndex        =   81
               Text            =   "0"
               Top             =   120
               Width           =   615
            End
            Begin VB.HScrollBar hsvstol 
               Enabled         =   0   'False
               Height          =   255
               Left            =   1200
               Max             =   199
               TabIndex        =   80
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox vstolanexo 
               Alignment       =   2  'Center
               Height          =   195
               Left            =   1560
               TabIndex        =   84
               Top             =   480
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.TextBox stallsalva 
               Height          =   255
               Left            =   2040
               TabIndex        =   83
               Top             =   480
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.Label Label35 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00404080&
               Caption         =   """ Com flap """
               ForeColor       =   &H00FFFF00&
               Height          =   240
               Left            =   120
               TabIndex        =   71
               Top             =   480
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Velocidade stall                    km/h"
               Height          =   195
               Left            =   240
               TabIndex        =   82
               Top             =   120
               Width           =   2385
            End
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00808080&
            Height          =   575
            Left            =   75
            ScaleHeight     =   510
            ScaleWidth      =   2820
            TabIndex        =   40
            Top             =   3600
            Width           =   2880
            Begin VB.CommandButton Command16 
               BackColor       =   &H80000000&
               Caption         =   "&Calcular o projeto"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   515
               Left            =   0
               MouseIcon       =   "Form1.frx":9BA8C
               MousePointer    =   99  'Custom
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "Clique aqui para calcular o projeto"
               Top             =   0
               Width           =   2820
            End
         End
         Begin VB.PictureBox Picture4 
            Height          =   855
            Left            =   75
            ScaleHeight     =   795
            ScaleWidth      =   2820
            TabIndex        =   12
            Top             =   2640
            Width           =   2880
            Begin VB.CommandButton Command17 
               Caption         =   "&Alterarações do projeto"
               Height          =   795
               Left            =   0
               MouseIcon       =   "Form1.frx":9BD96
               MousePointer    =   99  'Custom
               Picture         =   "Form1.frx":9C0A0
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   0
               Width           =   2820
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   11595
      TabIndex        =   0
      Top             =   0
      Width           =   11650
      Begin VB.CommandButton Command4 
         Caption         =   "&Trem de pouso"
         Enabled         =   0   'False
         Height          =   935
         Left            =   9960
         MouseIcon       =   "Form1.frx":9C3AA
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":9C6B4
         Style           =   1  'Graphical
         TabIndex        =   141
         ToolTipText     =   "Escolha o trem de pouso"
         Top             =   0
         Width           =   1655
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Perfis e coeficientes"
         Enabled         =   0   'False
         Height          =   935
         Left            =   8300
         MouseIcon       =   "Form1.frx":9F102
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":9F40C
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "Escolha os perfis aerodinâmicos"
         Top             =   0
         Width           =   1675
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Afilamento ( Lambda )"
         Enabled         =   0   'False
         Height          =   935
         Left            =   6360
         MouseIcon       =   "Form1.frx":A0DE6
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":A10F0
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "Escolha o afilamento"
         Top             =   0
         Width           =   1955
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Alongamento ( AR )"
         Enabled         =   0   'False
         Height          =   935
         Left            =   4440
         MouseIcon       =   "Form1.frx":A3B3E
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":A3E48
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Escolha o alongamento"
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Posição da asa"
         Enabled         =   0   'False
         Height          =   935
         Left            =   3000
         MouseIcon       =   "Form1.frx":A6896
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":A6BA0
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Escolha a posição da asa"
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Requisitos"
         Enabled         =   0   'False
         Height          =   935
         Left            =   1440
         MouseIcon       =   "Form1.frx":A95EE
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":A98F8
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Clique aqui para iniciar"
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Tipo"
         Enabled         =   0   'False
         Height          =   935
         Left            =   0
         MouseIcon       =   "Form1.frx":A9D3A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":AA044
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Clique aqui para iniciar"
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Caption         =   "Motor"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   131
      Top             =   960
      Width           =   3030
      Begin VB.OptionButton pusher 
         Caption         =   "Cauda (pusher)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   133
         Top             =   250
         Width           =   1455
      End
      Begin VB.OptionButton pull 
         Caption         =   "Nariz (pull)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   250
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Menu fileinicio1 
      Caption         =   "&Iniciar"
      Begin VB.Menu filenovo 
         Caption         =   "Novo"
      End
      Begin VB.Menu fileabrexiste 
         Caption         =   "Abrir existente"
      End
      Begin VB.Menu filesalva1 
         Caption         =   "Salvar"
      End
      Begin VB.Menu filedel 
         Caption         =   "Excluir projeto"
      End
      Begin VB.Menu fileir1 
         Caption         =   "Ir para"
         Begin VB.Menu filedesenho1 
            Caption         =   "Retornar ao desenho"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu filefer1 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "Calculadora && Conversor"
      End
   End
   Begin VB.Menu fileimprimi1 
      Caption         =   "Im&primir"
   End
   Begin VB.Menu filecadastro 
      Caption         =   "&Cadastrar"
      Begin VB.Menu fileperf 
         Caption         =   "Perfil"
         Begin VB.Menu fileperfilasa 
            Caption         =   "Asa"
         End
         Begin VB.Menu fileperht 
            Caption         =   "Estabilizador horizontal"
         End
         Begin VB.Menu fileprfvt 
            Caption         =   "Estabilizador vertical"
         End
      End
      Begin VB.Menu filemotor 
         Caption         =   "Motores"
      End
      Begin VB.Menu filematerial 
         Caption         =   "Materiais"
      End
   End
   Begin VB.Menu filesobre 
      Caption         =   "So&bre.."
      Begin VB.Menu fileCred 
         Caption         =   "Creditos"
      End
      Begin VB.Menu fileLiber 
         Caption         =   "Liberação"
      End
   End
   Begin VB.Menu filesair 
      Caption         =   "&Sair do Programa"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub adicionarea_GotFocus()
    adicionarea.SelStart = 0
    adicionarea.SelLength = Len(adicionarea)
End Sub
Private Sub adicionarea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 13 Then espbeq.SetFocus
End Sub
Private Sub afilaht_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 13 Then afilavt.SetFocus
End Sub
Private Sub afilaht_GotFocus()
If Not IsNumeric(afilasa) Or afilasa = "0" Or afilasa > 1 Then
    MsgBox "Insira o valor do afilamento conforme recomendado", vbOKOnly
    afilasa.SetFocus
    afilasa = 0
    Exit Sub
End If
afilaht.SelStart = 0
afilaht.SelLength = Len(afilaht)
End Sub
Private Sub afilasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then afilaht.SetFocus
End Sub
Private Sub afilasa_GotFocus()
If Not IsNumeric(arvt) Or arvt = "0" Then
    MsgBox "Insira o valor do alongamento conforme recomendado", vbOKOnly
    arvt.SetFocus
    arvt = 0
    Exit Sub
End If

afilasa.SelStart = 0
afilasa.SelLength = Len(afilasa)

Command7.Enabled = True
End Sub
Private Sub afilavt_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command7.Enabled = True
If KeyAscii = 13 Then Command7.SetFocus
End Sub
Private Sub afilavt_GotFocus()
If Not IsNumeric(afilaht) Or afilaht = "0" Or afilaht > 1 Then
    MsgBox "Insira o valor do afilamento conforme recomendado", vbOKOnly
    afilaht.SetFocus
    afilaht = 0
    Exit Sub
End If
afilavt.SelStart = 0
afilavt.SelLength = Len(afilavt)
End Sub
Private Sub afilavt_LostFocus()

Command7.Enabled = True
End Sub
Private Sub arht_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then arvt.SetFocus
End Sub
Private Sub arht_GotFocus()
If Not IsNumeric(arw) Or arw = "0" Then
    MsgBox "Insira o valor do alongamento conforme recomendado", vbOKOnly
    arw.SetFocus
    arw = 0
    Exit Sub
End If


arht.SelStart = 0
arht.SelLength = Len(arht)
End Sub

Private Sub arvt_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command6.SetFocus
End Sub
Private Sub arvt_GotFocus()
If Not IsNumeric(arht) Or arht = "0" Then
    MsgBox "Insira o valor do alongamento conforme recomendado", vbOKOnly
    arht.SetFocus
    arht = 0
    Exit Sub
End If

arvt.SelStart = 0
arvt.SelLength = Len(arvt)
End Sub
Private Sub arw_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then arht.SetFocus
End Sub
Private Sub arw_GotFocus()
arw.SelStart = 0
arw.SelLength = Len(arw)
Command6.Enabled = True
End Sub
Private Sub alterapeso_Click()
If Form1.alterapeso.Value = Unchecked Then
Form1.alterapeso.Caption = "Original"
Form1.alterapeso.Enabled = False
Form1.alterapeso.Value = Unchecked
End If
End Sub

Private Sub Command1_Click()
'JAN. 2004
Tipo.Visible = True

Form1.wo.Locked = False
Form1.motor.Locked = False
Form1.Command15.Enabled = True
Form1.Command3.Enabled = True
Frame2.Enabled = True
pusher.Enabled = True
pull.Enabled = True
End Sub

Private Sub Command10_Click()
Form24.Visible = True
End Sub

Private Sub Command11_Click()
Form23.Visible = True
End Sub

Private Sub Command12_Click()
Form22.Visible = True
End Sub

Private Sub Command13_Click()
Form3.Visible = True
End Sub

Private Sub Command14_Click()
Form26.Visible = True
End Sub



Private Sub Command15_Click()
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
' INICIO PCA USP
Form1.Command15.Visible = False
Command16.Enabled = True

' Teste carga alar
WSPCA = Val(Requisitos.wsR)
ro = Form1.roar * 0.00194

If Tipo.PCA.Value = True Then vstolFTS = ((2 * WSPCA) / (ro * Form1.clmaxFLAP)) ^ 0.5

Form1.vstol = vstolFTS * 1.09728 'em km/h

Form1.vstol = Format(Form1.vstol, "0.00")

Form1.vstolanexo = Form1.vstol
Form1.vstolanexo = Format(Form1.vstolanexo, "0.00")

stallsalva = hsvstol.Value

Command17.Enabled = True

'FIM PCA
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
End Sub

Private Sub Command16_Click()
        
If Form1.arw = Empty Or Form1.arht = Empty Or Form1.arvt = Empty Then
    MsgBox "Célula vazia em alongamento!", vbCritical
    Exit Sub
End If

If Form1.arw = 0 Or Form1.arht = 0 Or Form1.arvt = 0 Then
    MsgBox "O alongamento deve ser diferente de zero!", vbCritical
    Exit Sub
End If

If Form1.afilasa = Empty Or Form1.afilaht = Empty Or Form1.afilavt = Empty Then
    MsgBox "Célula vazia em afilamento!", vbCritical
    Exit Sub
End If

If Form1.afilasa = 0 Or Form1.afilaht = 0 Or Form1.afilavt = 0 Then
    MsgBox "O afilamento deve ser diferente de zero!", vbCritical
    Exit Sub
End If

If vstol = Empty Then
    MsgBox "Célula vazia em velocidade de estol !", vbCritical
    Exit Sub
End If


If vstol = 0 Then
    MsgBox "A velocidade de estol deve ser maior que zero!", vbCritical
    Exit Sub
End If



If Form1.tremconv.Value = False And Form1.tremtric.Value = False Then
    MsgBox "Escolha o tipo de trem de pouso", vbInformation
    Exit Sub
End If

If Form1.clmaxFLAP = "" Or Form1.clmaximo = "" Or Form1.clmaxht = "" Then
    MsgBox "Escolha o perfil para asa e empenagens (horizontal e vertical)", vbCritical
    Form4.Show
    variavelw = "Select * from Aerofolios where Perfil=Perfil"
    Form4.Data1.RecordSource = variavelw
    Form4.Data1.Refresh

    variavelht = "Select * from PerfilHT where Perfil=Perfil"
    Form4.Data2.RecordSource = variavelht
    Form4.Data2.Refresh

    variavelvt = "Select * from PerfilVT where Perfil=Perfil"
    Form4.Data3.RecordSource = variavelvt
    Form4.Data3.Refresh

    Form4.DBGrid1.Height = 695
    Form4.DBGrid2.Height = 695
    Form4.DBGrid3.Height = 695

    Form4.Frame5.Visible = False
    Form4.Command28.Enabled = True
    Form4.Command3.Enabled = True
    Form4.Command4.Enabled = True

    Form4.DBGrid1.Enabled = True
    Form4.DBGrid2.Enabled = True
    Form4.DBGrid3.Enabled = True

    Form4.DBGrid1.Height = 2040
    Form4.DBGrid1.Enabled = True
    Form4.DBGrid1.Visible = True

    Form4.DBGrid2.Visible = False
    Form4.DBGrid3.Visible = False
    MsgBox "Inicie escolhendo o perfil para asa.", vbInformation
    Exit Sub
End If

'ROTINA 14/11/2001

'CORREÇÃO DE REYNOLDS 23/12/01
densidade = roar ' kg/m3

sestim = (Form1.wo * 2.2) / (((densidade * 0.00194 * ((Form1.vstol * 0.911) ^ 2)) / 2) * Form1.clmaxFLAP * 0.85) ' Pes2
bestim = (sestim * Form1.arw) ^ 0.5  ' pe
lestim = (sestim / bestim)
speed = 1.15 * ((146618 * (Form1.motor / (0.1 * sestim))) ^ (1 / 3)) * 1.466 'ft/s
reyestim = 6363 * speed * lestim
reyestim = Format(reyestim, "0.00")
Form7.rey3d = Int(reyestim)
cofrey = 0.00000005
Form7.rey2d = Form4.Data1.Recordset("Reynolds")
difrey = (Form7.rey2d - Form7.rey3d) * cofrey
'fim correção

'Slope sem Flap
Form7.alfasamax = Form4.Data1.Recordset("angulo")
Form7.alfasamax = Format(Form7.alfasamax, "0.000")
Form7.alfasazero = Form4.Data1.Recordset("angulo0")
Form7.alfasazero = Format(Form7.alfasazero, "0.000")
deltaSF = Abs(Form7.alfasamax) + Abs(Form7.alfasazero)
slopesf = (Form1.clmaximo / deltaSF) * 57.3
Form7.slopeasa = slopesf
Form7.slopeasa = Format(Form7.slopeasa, "0.000")
Form7.cl3dasa = (((Form7.slopeasa / (1 + (Abs(Form7.slopeasa / (3.1415 * Form1.arw))))) / 57.3) * deltaSF) - difrey
Form7.cl3dasa = Format(Form7.cl3dasa, "0.000")

'Slope com Flap
difCL = Abs(Form1.clmaxFLAP) - Abs(Form1.clmaximo)
FK = 0.055
FK2 = 0.12
k = FK * Form4.deflap * difCL
k1 = k
k2 = FK2 * Form4.deflap * difCL
k11 = Abs(Form7.alfasamax) - Abs(k)
k22 = (Form7.alfasazero) - (k2)

kk = Abs(k11) + Abs(k22)
Form7.alfaFLAPm = k11
Form7.alfaFLAPm = Format(Form7.alfaFLAPm, "0.000")
Form7.alfaFLAPo = k22
Form7.alfaFLAPo = Format(Form7.alfaFLAPo, "0.000")
slopeCF = (Form1.clmaxFLAP / kk) * 57.3
Form7.slopeFLAP = slopeCF
Form7.slopeFLAP = Format(Form7.slopeFLAP, "0.000")
Form7.cl3dFLAP = (((Form7.slopeFLAP / (1 + (Abs(Form7.slopeFLAP / (3.1415 * Form1.arw))))) / 57.3) * kk) - difrey
Form7.cl3dFLAP = Format(Form7.cl3dFLAP, "0.000")


If Form1.clmaximo = Form1.clmaxFLAP Then cofsust = Form7.cl3dasa
If Form1.clmaximo > Form1.clmaxFLAP Then cofsust = Form7.cl3dasa
If Form1.clmaximo < Form1.clmaxFLAP Then cofsust = Form7.cl3dFLAP

If Form2.Sw = Empty Then GoTo 10

If Form2.Sw <> Empty Then

respstol = MsgBox("A área da asa será mantida?", vbYesNo)
        If respstol = vbYes Then
            Form2.Sw = Form2.Sw
            dar = 1.225 * 0.00194
            
                vteste = (((2 * Form1.wo * 2.2) / (cofsust * dar * Form2.Sw * 10.72)) ^ 0.5)
                vteste = Format(vteste, "0.00")
            
                Form1.vstol = vteste * 1.09728 ' em km/h
                Form1.vstol = Format(Form1.vstol, "0.00")
                
                Form8.vstol = vteste * 1.09728 ' em km/h
                Form8.vstol = Format(Form8.vstol, "0.00")
                
                Form1.vstolanexo = Form1.vstol
                Form1.vstolanexo = Format(Form1.vstolanexo, "0.00")
            
                    If Form4.semflap.Value = False Then Form1.Label35.Visible = True
            
If Form4.semflap.Value = True Then Form1.Label35.Visible = False
                                
                respcontinua = MsgBox("Iniciar o dimensionamento?", vbYesNo)
                If respcontinua = vbYes Then GoTo 20
                If respcontinua = vbNo Then Exit Sub
                
                Exit Sub
            End If
            
            If respstol = vbNo Then
            
                dar = 1.225 * 0.00194

                Form2.Sw = ((2 * Form1.wo * 2.2) / (dar * ((Form1.vstol * 0.9113) ^ 2) * cofsust)) / 10.76
                Form2.Sw = Format(Form2.Sw, "0.000")
                
                Form1.vstolanexo = Form1.vstol
                Form1.vstolanexo = Format(Form1.vstolanexo, "0.00")
                
                GoTo 20
            End If
            
  
End If

'respcontinua = MsgBox("Deseja continuar os cálculos??", vbYesNo)
'If respcontinua = vbYes Then GoTo 10
'If respcontinua = vbNo Then Exit Sub

Form7.cl3dFLAP = Format(Form7.cl3dFLAP, "0.000")
Form7.cl3dasa = Format(Form7.cl3dasa, "0.000")

' vvvvvvvvvv VAZIAS vvvvvvvvvvvvv
If Form1.wo = Empty Then
MsgBox "Voce deve escolher o motor para o avião.", vbInformation
Exit Sub
End If

If Tipo.PCA.Value = False Then
MsgBox "Voce não escolheu o tipo de avião.", vbInformation
Exit Sub
End If




'Lembrar de escolher Perfil
If clmaximo = Empty Then
MsgBox "Voce deve escolher os perfis para asa e estabilizadores.", vbInformation
Form4.Visible = True
Exit Sub
End If

                                               
'Lembrar de clicar Vstall
If Form1.vstol = 0 Then
    MsgBox "Voce deve calcular a velocidade de stall", vbInformation
    Command15.Enabled = True
    Exit Sub
End If


'Dimensionamento

If peso < 0 Then kpeso = ((100 + peso) / 100)
If peso > 0 Then kpeso = 1 + (peso / 100)
If peso = 0 Then kpeso = 1

wo = woanexo * kpeso
wo = Format(wo, "0.00")




'DIMENSIONAMENTO - Áreas e Asa

10:

dar = 1.225 * 0.00194

Form2.Sw = ((2 * Form1.wo * 2.2) / (dar * ((Form1.vstol * 0.9113) ^ 2) * cofsust)) / 10.76
Form2.Sw = Format(Form2.Sw, "0.000")

20:
Form2.bw = (Form2.Sw * Form1.arw) ^ 0.5
Form2.bw = Format(Form2.bw, "0.000")

Form2.crw = Form2.Sw / (((1 + Form1.afilasa) / 2) * Form2.bw)
Form2.crw = Format(Form2.crw, "0.00")

Form2.cmw = (Form2.Sw / (((1 + Form1.afilasa) / 2) * Form2.bw) + Form2.crw * Form1.afilasa) / 2
Form2.cmw = Format(Form2.cmw, "0.00")

Form7.CMaerodw = ((2 / 3) * (Form2.crw) * ((1 + Form1.afilasa) + (Form1.afilasa ^ 2))) / (Form1.afilasa + 1)
Form7.CMaerodw = Format(Form7.CMaerodw, "0.000")

'Volumes de cauda
    
'PARA PCA USP
If Tipo.PCA.Value = True Then
        'ht1 = 0.35 ' Inicial, o ideal é 0.5
        'Cancelado em 29/11/2007
        'If Tipo.Option1.Value = True Then ht1 = (-0.000005 * Form1.wo * 2.2) + 0.4592
        'If Tipo.Option2.Value = True Then ht1 = (0.0002 * Form1.wo * 2.2) + 0.2235
        'If Tipo.Option3.Value = True Then ht1 = (0.00009 * Form1.wo * 2.2) + 0.3187
        'If Tipo.Option4.Value = True Then ht1 = (0.0001 * Form1.wo * 2.2) + 0.1327
        'If Tipo.Option5.Value = True Then ht1 = 0.4
        
        'Criado em 29/11/2007 para TRmilitar
        If Tipo.Option1.Value = True Or Tipo.Option2.Value = True Or Tipo.Option3.Value = True Or Tipo.Option4.Value = True Or Tipo.Option5.Value = True Then Ht1 = 0.000006 * Form1.wo * 2.2 + 0.6

        
If volumeht < 0 Then kvolumeht = ((100 + volumeht) / 100)
If volumeht > 0 Then kvolumeht = 1 + (volumeht / 100)
If volumeht = 0 Then kvolumeht = 1

    fht = (kvolumeht * Ht1)
    
        'xxxxxxxxxxxxxxxx vt1 = 0.02 ' Inicial, o ideal é 0.04 xxxxxxxxxxxxxxxxxxxxx
        'Cancelado em 29/11/2007
        'If Tipo.Option1.Value = True Then vt1 = (-0.000002 * Form1.wo * 2.2) + 0.038
        'If Tipo.Option2.Value = True Then vt1 = (0.00002 * Form1.wo * 2.2) - 0.0006
        'If Tipo.Option3.Value = True Then vt1 = (0.000005 * Form1.wo * 2.2) + 0.0242
        'If Tipo.Option4.Value = True Then vt1 = (-0.000003 * Form1.wo * 2.2) + 0.0278
        'If Tipo.Option5.Value = True Then vt1 = 0.035
        
        'Criado em 29/11/2007 para TRmilitar
        
If Tipo.Option1.Value = True Or Tipo.Option2.Value = True Or Tipo.Option3.Value = True Then vt1 = 0.000003 * Form1.wo * 2.2 + 0.03
    If Tipo.Option4.Value = True Or Tipo.Option5.Value = True Then vt1 = 0.000005 * Form1.wo * 2.2 + 0.05
        
If volumevt < 0 Then kvolumevt = ((100 + volumevt) / 100)
If volumevt > 0 Then kvolumevt = 1 + (volumevt / 100)
If volumevt = 0 Then kvolumevt = 1

    fvt = (kvolumevt * vt1)
    
        'superficies de controle
        'Com flap
    If Form4.fowlerflap.Value = True Or Form4.plainflap.Value = True Or Form4.splitflap.Value = True And Form4.semflap.Value = False Then
        fbail = 0.42
        fcail = Form4.porcAIL
    End If
        'Sem flap
    If Form4.fowlerflap.Value = False And Form4.plainflap.Value = False And Form4.splitflap.Value = False And Form4.semflap.Value = True Then
        fbail = Form4.porcenvail
        fcail = Form4.porcAIL
    End If
    
    
        fcp = Form4.porcEH
        fcld = Form4.porcEH
        eh = 9
        ev = 9
End If
'Calculo de lt

'Fatores de fuselagem

'PCA USP
If Tipo.PCA.Value = True Then fatlargo = 1
If Tipo.PCA.Value = True And Tipo.Option1.Value = True Then fatalto = 1.45 'Era 1 em 30/12/2007
If Tipo.PCA.Value = True And Tipo.Option2.Value = True Then fatalto = 1.55 'Era 1 em 30/12/2007
If Tipo.PCA.Value = True And Tipo.Option3.Value = True Then fatalto = 1.65 'Era 1 em 30/12/2007
If Tipo.PCA.Value = True And Tipo.Option4.Value = True Then fatalto = 1.7 'Era 1 em 30/12/2007
If Tipo.PCA.Value = True And Tipo.Option5.Value = True Then fatalto = 1.8 'Era 1 em 30/12/2007

'Fuselagem - comprimento
'validando fator de fuselagem - fuse
If fuse < 0 Then kfuse = ((100 + fuse) / 100)
If fuse > 0 Then kfuse = 1 + (fuse / 100)
If fuse = 0 Then kfuse = 1
'largura
If largo < 0 Then klargo = ((100 + largo) / 100)
If largo > 0 Then klargo = 1 + (largo / 100)
If largo = 0 Then klargo = 1
'altura
If alto < 0 Then kalto = ((100 + alto) / 100)
If alto > 0 Then kalto = 1 + (alto / 100)
If alto = 0 Then kalto = 1

'Comprimento de fuselagem para PCA
'Form2.lfs = ((Form1.wo * 4.37) ^ 0.23) * kfuse 'Era este em 30/12/2007
Form2.lfs = ((Form1.wo * 4.75) ^ 0.23) * kfuse
Form2.lfs = Format(Form2.lfs, "0.000")

If Requisitos.ocup > 15 Then
Form2.lfs = ((Form1.wo * 20) ^ 0.23) * kfuse
Form2.lfs = Format(Form2.lfs, "0.000")
End If

'Nariz
If Tipo.PCA.Value = True And Tipo.Option1.Value = True Then fdnc = 0.32 * Form2.lfs
If Tipo.PCA.Value = True And Tipo.Option2.Value = True Then fdnc = 0.32 * Form2.lfs
If Tipo.PCA.Value = True And Tipo.Option3.Value = True Then fdnc = 0.35 * Form2.lfs
If Tipo.PCA.Value = True And Tipo.Option4.Value = True Then fdnc = 0.375 * Form2.lfs
If Form1.pusher.Value = True Then fdnc = 0.5 * Form2.lfs

If nariz < 0 Then knazo = ((100 + nariz) / 100)
If nariz > 0 Then knazo = 1 + (nariz / 100)
If nariz = 0 Then knazo = 1

dnc = knazo * fdnc

'.............. LOOP LT .............................

Aguarde1.Visible = True
Aguarde1.Max = Form2.lfs
For klt = 0.2 To Form2.lfs Step 0.01 'era 0.001 em 24/01/04
ltteste = klt

'barra de rolagem
Aguarde1.Value = klt

Form2.sht = fht * Form2.Sw * Form7.CMaerodw / (ltteste)
Form2.sht = Format(Form2.sht, "0.000")
Form2.svt = fvt * Form2.Sw * Form2.bw / (ltteste)
Form2.svt = Format(Form2.svt, "0.000")

If Tipo.PCA.Value = True Then
    fel = 0.35 * Form2.sht
    fld = 0.35 * Form2.svt
End If

Form2.sprofund = fel
Form2.sprofund = Format(Form2.sprofund, "0.000")
Form2.slemdir = fld
Form2.slemdir = Format(Form2.slemdir, "0.000")

'Envergadura
Form2.bail = fbail * (Form2.bw / 2)
Form2.bail = Format(Form2.bail, "0.000")
Form2.bht = (Form2.sht * Form1.arht) ^ 0.5
Form2.bht = Format(Form2.bht, "0.000")
Form2.bprofund = Form2.bht
Form2.bvt = (Form2.svt * Form1.arvt) ^ 0.5
Form2.bvt = Format(Form2.bvt, "0.000")
Form2.blemdir = Form2.bvt

'Cordas
Form2.crw = Form2.Sw / (((1 + Form1.afilasa) / 2) * Form2.bw)
Form2.crw = Format(Form2.crw, "0.000")
Form2.cpw = Form2.crw * Form1.afilasa
Form2.cpw = Format(Form2.cpw, "0.000")
Form2.cmw = (Form2.Sw / (((1 + Form1.afilasa) / 2) * Form2.bw) + Form2.crw * Form1.afilasa) / 2
Form2.cmw = Format(Form2.cmw, "0.000")

Form2.crail = Form2.crw * fcail
Form2.crail = Format(Form2.crail, "0.000")
Form2.cpail = Form2.cpw * fcail
Form2.cpail = Format(Form2.cpail, "0.000")
Form2.sail = ((Abs(Form2.crail) + Abs(Form2.cpail)) / 2) * Form2.bail
Form2.sail = Format(Form2.sail, "0.00")

Form2.crht = Form2.sht / (((1 + Form1.afilaht) / 2) * Form2.bht)
Form2.crht = Format(Form2.crht, "0.000")
Form2.cpht = Form2.crht * Form1.afilaht
Form2.cpht = Format(Form2.cpht, "0.000")
Form2.cmht = (Form2.sht / (((1 + Form1.afilaht) / 2) * Form2.bht) + (Form2.crht * Form1.afilaht)) / 2
Form2.cmht = Format(Form2.cmht, "0.000")

Form7.CMaerodht = ((2 / 3) * (Form2.crht) * ((1 + Form1.afilaht) + (Form1.afilaht ^ 2))) / (Form1.afilaht + 1)
Form7.CMaerodht = Format(Form7.CMaerodht, "0.000")

Form2.crprofund = fcp * Form2.cpht
Form2.crprofund = Format(Form2.crprofund, "0.000")
Form2.cpprofund = Form2.crprofund
Form2.cpprofund = Format(Form2.cpprofund, "0.000")

    If Form2.bvt = 0 Then
        MsgBox " Diminua a velocidade de estol.", vbCritical
        Exit Sub
    End If
Form2.crvt = Form2.svt / (((1 + Form1.afilavt) / 2) * Form2.bvt)
Form2.crvt = Format(Form2.crvt, "0.000")
Form2.cpvt = Form2.crvt * Form1.afilavt
Form2.cpvt = Format(Form2.cpvt, "0.000")
Form2.cmvt = (Form2.svt / (((1 + Form1.afilavt) / 2) * Form2.bvt) + (Form2.crvt * Form1.afilavt)) / 2
Form2.cmvt = Format(Form2.cmvt, "0.000")

Form2.crlemdir = fcld * Form2.cpvt
Form2.crlemdir = Format(Form2.crlemdir, "0.000")
Form2.cplemdir = Form2.crlemdir
Form2.cplemdir = Format(Form2.cplemdir, "0.000")

'Calculo do lta
lta = (Abs(Form2.lfs) - (Abs(Form2.cmht * 0.5) + Abs(dnc)))
lta = Format(lta, "0.000")

If Int(ltteste * 100) - Int(lta * 100) = 0 Then
LT = lta

Form10.LT = LT
Form10.LT = Format(Form10.LT, "0.000")

Form10.LRCAasa = dnc
Form10.LRCAasa = Format(Form10.LRCAasa, "0.000")

Form10.LRCAht = Abs(Form10.LRCAasa) + Abs(Form10.LT)
Form10.LRCAht = Format(Form10.LRCAht, "0.000")


'velocidade de stall
'dar = form1.roar * 0.00194

'vst = ((wo * 2.2) / ((dar / 2) * Form2.Sw * 10.76 * clmaximo3d)) ^ 0.5
'vstall = vst * 1.097
'vstall = Format(vstall, "0.000")

'carga alar calculada wscalculada
wscalculada = Form1.wo / Form2.Sw
wscalculada = Format(wscalculada, "0.000")

'Formula planilha excel Tese
'Largura fuselagem com espessura parede
Form2.wfs = 0.305 * (3.5 * (Form2.lfs ^ 0.018) * fatlargo) * klargo
Form2.wfs = Format(Form2.wfs, "0.000")
'altura fuselagem com espessura parede
Form2.hfs = 0.305 * (2.2 * (Form2.lfs ^ 0.18) * fatalto) * kalto
Form2.hfs = Format(Form2.hfs, "0.000")
'volume de cabine
Form2.volcab = Form2.hfs * Form2.wfs * (Requisitos.ocup * 0.6)
Form2.volcab = Format(Form2.volcab, "0.00")

Form10.volht = fht
Form10.volht = Format(Form10.volht, "0.000")

Form10.volvt = fvt
Form10.volvt = Format(Form10.volvt, "0.000")


'FLAP
    If Form4.fowlerflap.Value = True Or Form4.plainflap.Value = True Or Form4.splitflap.Value = True And Form4.semflap.Value = False Then
        Form2.bflap = 0.4 * (Form2.bw / 2)
        Form2.bflap = Format(Form2.bflap, "0.000")
        Form2.cmflap = 0.3 * Form2.cmw
        Form2.cmflap = Format(Form2.cmflap, "0.000")
        If Form4.fowlerflap.Value = True Then Form2.tipoflap = "Fowler 30%Cm"
        If Form4.plainflap.Value = True Then Form2.tipoflap = "Plain 30%Cm"
        If Form4.splitflap.Value = True Then Form2.tipoflap = "Split30% Cm"
    End If
    
    If Form4.fowlerflap.Value = False And Form4.plainflap.Value = False And Form4.splitflap.Value = False And Form4.semflap.Value = True Then
        Form2.tipoflap = "Sem flap"
        Form2.bflap = "0,000"
        Form2.cmflap = "0,000"
    End If

Form2.anguflap = Form4.deflap

Aguarde1.Visible = False

' Form1.Visible = False
Form2.Visible = True
Form2.Refresh

'Modificando botões do Form1
Frame13.Visible = True
Form1.Command17.Enabled = True
Form1.Refresh
'Fim

Exit Sub
End If
                        
Next

If klt > 1 Then
MsgBox "A velocidade de stall está muito baixa, aumente a velocidade de stall.", vbCritical
Aguarde1.Visible = False
Call Command17_Click
Exit Sub
End If


'FIM LOOP LT

'Form1.Visible = False
Form2.Visible = True
Command8.Visible = False
   
End Sub

Private Sub Command17_Click()
Frame13.Visible = False
hsvstol.Enabled = True
Command15.Enabled = True
Command15.Visible = True
End Sub

Private Sub Command18_Click()
Command11.Enabled = True
Call Command11_Click
End Sub







Private Sub Command2_Click()
Requisitos.Show
End Sub

Private Sub Command3_Click()
Command10.Enabled = True
MsgBox "Escolha a posição da asa do avião na fuselagem ", vbInformation
End Sub

Private Sub Command4_Click()
Command11.Enabled = True
MsgBox "Escolha a posição da bequilha. No nariz (triciclo) ou na cauda (convencional)", vbInformation
End Sub

Private Sub Command5_Click()
Command12.Enabled = True
MsgBox "Relação entre o comprimento(envergadura) e a largura(corda) de superfícies.         Detalhes em Ajuda", vbInformation
arw.SetFocus
End Sub

Private Sub Command6_Click()
Command13.Enabled = True
MsgBox "Porcentagem entre a corda da ponta e corda da raiz de superficíes. Detalhes em Ajuda", vbInformation
afilasa.SetFocus
End Sub

Private Sub Command7_Click()
If Not IsNumeric(afilavt) Or afilavt = "0" Or afilavt > 1 Then
    MsgBox "Insira o valor do afilamento conforme recomendado", vbOKOnly
    afilavt.SetFocus
    afilavt = 0
    Exit Sub
End If


Command14.Enabled = True
Form4.Show
If Form1.posialta.Value = True Then Form4.Lasa.y1 = 720
If Form1.posialta.Value = True Then Form4.Lasa.y2 = 720

If Form1.posimedia.Value = True Then Form4.Lasa.y1 = 840
If Form1.posimedia.Value = True Then Form4.Lasa.y2 = 840

If Form1.posibaixa.Value = True Then Form4.Lasa.y1 = 960
If Form1.posibaixa.Value = True Then Form4.Lasa.y2 = 960

End Sub

Private Sub Command8_Click()
Command8.Visible = False
End Sub
Private Sub Command9_Click()
Requisitos.Show
End Sub

Private Sub diambeq_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then esptp.SetFocus
End Sub
Private Sub diambeq_GotFocus()
diambeq.SelStart = 0
diambeq.SelLength = Len(diambeq)
End Sub
Private Sub diambeq_LostFocus()
If diambeq = Empty Then
    MsgBox "Célula vazia", vbOKOnly
    diambeq.SetFocus
    diambeq = 0
End If
End Sub

Private Sub diamtp_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then KeyAscii = 13
End Sub
Private Sub diamtp_GotFocus()
diamtp.SelStart = 0
diamtp.SelLength = Len(diamtp)
End Sub
Private Sub diamtp_LostFocus()
If diamtp = Empty Then
    MsgBox "Célula vazia", vbOKOnly
    diamtp.SetFocus
    diamtp = 0
End If
End Sub


Private Sub espbeq_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then diambeq.SetFocus
End Sub
Private Sub espbeq_GotFocus()
espbeq.SelStart = 0
espbeq.SelLength = Len(espbeq)
End Sub
Private Sub espbeq_LostFocus()
If espbeq = Empty Then
    MsgBox "Célula vazia", vbOKOnly
    espbeq.SetFocus
    espbeq = 0
End If
End Sub

Private Sub esptp_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then diamtp.SetFocus
End Sub
Private Sub esptp_GotFocus()
esptp.SelStart = 0
esptp.SelLength = Len(esptp)
End Sub
Private Sub esptp_LostFocus()
If esptp = Empty Then
    MsgBox "Célula vazia", vbOKOnly
    esptp.SetFocus
    esptp = 0
End If
End Sub

Private Sub fileabrexiste_Click()
If Form21.Data1.Recordset.RecordCount = 0 Then
MsgBox "Não existem projetos salvos", vbInformation
Exit Sub
End If

Form21.Show
Form21.Command1.Visible = True
Form21.Command4.Visible = False
Form21.Command2.Visible = False

varabre = "Select * from Salvar where Projeto=Projeto"
Form21.Data1.RecordSource = varabre
Form21.Data1.Refresh
End Sub


Private Sub filecalcon_Click()
Form29.Show
End Sub

Private Sub fileCred_Click()
Form33.Show
End Sub

Private Sub filedel_Click()
Form21.Show
Form21.Command1.Visible = False
Form21.Command4.Visible = False
Form21.Command2.Visible = True
End Sub

Private Sub filedesenho1_Click()
Form18.Visible = True
'Form1.Visible = False
Form18.Refresh
End Sub

Private Sub fileimprimi1_Click()
If LIBER <> "KFAR0852" Then
    MsgBox "A rotina de impressão sómente estará disponível após a liberação permanente", 16, "Cancelamento de impressão": Exit Sub
End If

We = Requisitos.weR / 2.2
We = Format(We, "0.00")
Wf = Requisitos.wfR / 2.2
Wf = Format(Wf, "0.00")
Wcu = Requisitos.CPR / 2.2
Wcu = Format(Wcu, "0.00")
Wtot = Requisitos.woestimadoR / 2.2
Wtot = Format(Wtot, "0.00")

If Form1.posialta.Value = True Then asa = "Asa alta"
If Form1.posimedia.Value = True Then asa = "Asa média"
If Form1.posibaixa.Value = True Then asa = "Asa baixa"

If Form1.tremconv.Value = True Then trem = "Convencional -"
If Form1.tremtric.Value = True Then trem = "Triciclo -"

If Form1.Checkretratil = Checked Then sistema = "Trem retrátil"
If Form1.Checkretratil = Unchecked Then sistema = " Trem fixo"

Printer.FontBold = True
Printer.FontSize = 14
Printer.Print ""
Printer.Print ""
Printer.Print Tab(27.5); "PROJETO CONCEITUAL DE AERONAVES"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print ""
Printer.Print Tab(22.5); "REQUISITOS OPERACIONAIS"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 10
Printer.Print Tab(20); "Nome do projeto:"; Tab(85); Form20.p1
Printer.Print Tab(20); "Nome do projetista:"; Tab(85); Form20.p2
Printer.Print Tab(20); "Data do projeto:"; Tab(85); Form20.datatual
Printer.Print ""
Printer.Print Tab(20); "Posição da asa:"; Tab(85); asa
Printer.Print Tab(20); "Trem de pouso:"; Tab(85); trem + sistema
Printer.Print Tab(20); "Motor:"; Tab(85); Form1.motor; " Hp"
Printer.Print Tab(20); "Peso do avião:"; Tab(85); Form1.wo; " kgf"
Printer.Print Tab(20); "Coeficiente de sustentação máx. Asa Cl:"; Tab(85); Form1.clmaximo
Printer.Print Tab(20); "Coeficiente de sustentação máx. Asa/flap Clflap:"; Tab(85); Form1.clmaxFLAP
Printer.Print Tab(20); "Coeficiente de sustentação máx. Empenagem Clht:"; Tab(85); Form1.clmaxht
Printer.Print Tab(20); "Alongamento da asa :"; Tab(85); Form1.arw
Printer.Print Tab(20); "Alongamento da empenagem horizontal:"; Tab(85); Form1.arht
Printer.Print Tab(20); "Alongamento da empenagem vertical:"; Tab(85); Form1.arvt
Printer.Print Tab(20); "Afilamento da asa :"; Tab(85); Form1.afilasa
Printer.Print Tab(20); "Afilamento da empenagem horizontal:"; Tab(85); Form1.afilaht
Printer.Print Tab(20); "Afilamento da empenagem vertical:"; Tab(85); Form1.afilavt
Printer.Print Tab(20); "Área adicional"; Tab(85); Form1.adicionarea; " m2"
Printer.Print Tab(20); "Diametro da bequilha"; Tab(85); Form1.diambeq; " cm"
Printer.Print Tab(20); "Espessura da bequilha"; Tab(85); Form1.espbeq; " cm"
Printer.Print Tab(20); "Diametro do trem principal"; Tab(85); Form1.diamtp; " cm"
Printer.Print Tab(20); "Espessura do trem principal"; Tab(85); Form1.esptp; " cm"
Printer.Print Tab(20); "Velocidade de Stall"; Tab(85); Form1.vstol; "  km/h"
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print ""; Tab(22.5); "ALTERAÇÕES"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 10
Printer.Print Tab(20); "Alteração - Volume de Ht"; Tab(85); Form1.volumeht; " %"
Printer.Print Tab(20); "Alteração - Volume de Vt"; Tab(85); Form1.volumevt; " %"
Printer.Print Tab(20); "Alteração - Comprimento da fuselagem"; Tab(85); Form1.fuse; " %"
Printer.Print Tab(20); "Alteração - Largura da fuselagem"; Tab(85); Form1.largo; " %"
Printer.Print Tab(20); "Alteração - Altura da fuselagem"; Tab(85); Form1.alto; " %"
Printer.Print Tab(20); "Alteração - Nariz (Hélice-CAasa)"; Tab(85); Form1.nariz; " %"
Printer.Print Tab(20); "Alteração - Margem estática "; Tab(85); Form1.margestat; " %"
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print ""; Tab(22.5); "CÁLCULOS INICIAIS"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 10
Printer.Print Tab(20); "Peso Vazio"; Tab(85); We; " kgf"
Printer.Print Tab(20); "Peso de Combustivel"; Tab(85); Wf; " kgf"
Printer.Print Tab(20); "Peso de Carga Util"; Tab(85); Wcu; " kgf"
Printer.Print Tab(20); "Peso Total"; Tab(85); Wtot; " kgf"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 10
Printer.Print ""
Printer.Print Tab(20); "Fim da página"
'14/03/2004 altereiPrinter.Print Tab(5); "Alteração - Peso"; Tab(65); Form1.wo; " %"
Printer.EndDoc
End Sub

Private Sub fileLiber_Click()
Form32.Label1.Width = 2985
Form32.Label4.Width = 2985
Form32.Label6.Width = 2985

    Form32.Show 1
End Sub

Private Sub filematerial_Click()
Form13.Visible = False
Form28.Frame1.Visible = False
Form28.Frame2.Visible = False
Form28.Frame3.Visible = True
Form28.Visible = True
Form28.n13.SetFocus
End Sub

Private Sub filemotor_Click()
Form28.Visible = True
Form28.Frame1.Visible = True
Form28.Frame2.Visible = False
Form28.Frame3.Visible = False
Form28.t1.SetFocus
End Sub

Private Sub filenovo_Click()
Form1.Label35.Visible = False
Form1.Command1.Enabled = False
Form1.Command3.Enabled = False
Form1.Command4.Enabled = False
Form1.Command5.Enabled = False
Form1.Command6.Enabled = False
Form1.Command7.Enabled = False
Form1.Command8.Enabled = False
Form1.Command16.Enabled = False
Form1.Command15.Enabled = False
Form1.Command9.Enabled = False

'Ajudas

Form1.Command10.Enabled = False
Form1.Command11.Enabled = False
Form1.Command12.Enabled = False
Form1.Command13.Enabled = False
Form1.Command14.Enabled = False


Form1.Command17.Enabled = False

Form20.Visible = True
Form20.p1.SetFocus

' ZERANDO TUDO

'zerar inclusive o form2.sw
Form2.Sw = Empty

Form1.motor = 0
'Form1.woanexo = 0
Form1.wo = 0


Form1.posibaixa.Value = False
Form1.posimedia.Value = False
Form1.posialta.Value = False

Form1.arw = 0
Form1.arht = 0
Form1.arvt = 0

Form1.afilasa = 0
Form1.afilaht = 0
Form1.afilavt = 0

Form1.clmaximo = 0
Form1.clmaxFLAP = 0
Form1.clmaxht = 0

Form1.tremconv.Value = False
Form1.tremtric.Value = False
Form1.Checkretratil.Value = Unchecked

'ZERANDO FORM4
Form4.perfil1 = Empty
Form4.perfil2 = Empty
Form4.perfil3 = Empty
Form4.semflap.Value = True
Form4.flap.x1 = 1800
Form4.flap.y1 = 480
Form4.flap.X2 = 2160
Form4.flap.y2 = 480
Form4.clsflap = Empty
Form4.clmaxfinal = Empty

Form4.Vseficht.Value = 4

variavelw = "Select * from Aerofolios where Perfil=Perfil"
Form4.Data1.RecordSource = variavelw
Form4.Data1.Refresh

variavelht = "Select * from PerfilHT where Perfil=Perfil"
Form4.Data2.RecordSource = variavelht
Form4.Data2.Refresh

variavelvt = "Select * from PerfilVT where Perfil=Perfil"
Form4.Data3.RecordSource = variavelvt
Form4.Data3.Refresh

Form4.DBGrid1.Height = 2040
Form4.DBGrid2.Height = 695
Form4.DBGrid3.Height = 695

Form4.Frame5.Visible = False
Form4.Command28.Enabled = True
Form4.Command3.Enabled = True
Form4.Command4.Enabled = True

Form4.DBGrid1.Enabled = True
Form4.DBGrid2.Enabled = True
Form4.DBGrid3.Enabled = True
Form4.Command23.Enabled = False
'FIM FORM4

'ZERANDO O FORM27
Form27.Option4.Value = False

Form1.hsfuse.Value = 100
Form1.hsalto.Value = 100
Form1.hslargo.Value = 100
Form1.hsvolumeht.Value = 100
Form1.hsvolumevt.Value = 100
Form1.hsnariz.Value = 100
Form1.adicionarea = 100
Form1.hsmargem.Value = 15
'Form1.hspeso.Value = 100
Form1.hsvstol.Value = 0

Form1.fuse = 0
Form1.largo = 0
Form1.alto = 0
Form1.volumeht = 0
Form1.volumevt = 0
Form1.nariz = 0
Form1.adicionarea = 0
'Form1.peso = 0
Form1.vstol = 0




'ZERANDO O FORM13- planilha estruturas
Form13.vigaret.Value = True
Form13.matmesa = Empty
Form13.matalma = Empty
Form13.sigcad = Empty
Form13.sigtad = Empty
Form13.cizad = Empty
Form13.elasto = Empty
Form13.gsatual = Empty
Form13.gsnegatual = Empty
Form13.checkg.Value = Unchecked
Form13.Checkcgorig.Value = Unchecked
Form13.Command3.Enabled = True
Form13.Command4.Enabled = True
variavelest = "Select * from materiais where nome=nome"
Form13.Data1.RecordSource = variavelest
Form13.Data1.Refresh
'FIM FORM13



'ZERANDO FORM14
Form14.filevisual14.Enabled = False
Form14.filevisual14.Visible = False

'Form14.Option1.Value = False
'Form14.Option2.Value = False
'Form14.Option1.Enabled = True
'Form14.Option2.Enabled = True

Form14.p1 = 0
Form14.d1 = 0

Form14.p2 = 0
Form14.d2 = 0

Form14.p11 = 0
Form14.d11 = 0

Form14.p12 = 0
Form14.d12 = 0

Form14.p5 = 0
Form14.d5 = 0

Form14.p6 = 0
Form14.d6 = 0

Form14.p7 = 0
Form14.d7 = 0

Form14.p10 = 0
Form14.d10 = 0

Form14.p4 = 0
Form14.d4 = 0

Form14.p3 = 0
Form14.d3 = 0

Form14.p13 = 0
Form14.d13 = 0

Form14.p14 = 0
Form14.d14 = 0

Form14.p15 = 0
Form14.d15 = 0
'FIM FORM14

'ZERANDO O FORM 27
Form27.clafaw = 0
Form27.clalfah = 0
Form27.PN27 = 0
Form27.xcg = 0
Form27.cmfuse = 0
Form27.nih = 0
Form27.shsw = 0
Form27.xachcg = 0
Form27.cecp = 0
Form27.drodf = 0
Form27.Ew = 0
Form27.dedalfa = 0
'Form27.testecl = 0

Form27.Picture1.AutoRedraw = True
Form27.Picture1.Cls
Form27.kde1.Visible = False
Form27.kde2.Visible = False
Form27.kde3.Visible = False
Form27.kde4.Visible = False
Form27.kde11.Visible = False
Form27.kde22.Visible = False
Form27.kde33.Visible = False

Form27.clponto.Visible = False

Form27.cmg0.Visible = False
Form27.cmg1.Visible = False
Form27.cmg2.Visible = False
Form27.cmg3.Visible = False
Form27.cmg11.Visible = False
Form27.cmg22.Visible = False
Form27.cmg33.Visible = False
Form27.defgrau.Visible = False

Form27.cmcgok1.Visible = False

Form27.Frame3.Visible = True
'Fim form27

Form1.Command8.Visible = True
Form1.Command15.Visible = True
Form1.Frame13.Visible = True

Form1.alterapeso.Value = 0
End Sub

Private Sub fileperfilasa_Click()
Form28.Visible = True
Form4.Visible = False
Form28.Frame1.Visible = False
Form28.Frame2.Visible = True
Form28.Frame3.Visible = False
End Sub

Private Sub fileperht_Click()
Form28.Visible = True
Form4.Visible = False
Form28.Frame1.Visible = False
Form28.Frame2.Visible = True
Form28.Frame3.Visible = False
End Sub

Private Sub fileprfvt_Click()
Form28.Visible = True
Form4.Visible = False
Form28.Frame1.Visible = False
Form28.Frame2.Visible = True
Form28.Frame3.Visible = False
End Sub

Private Sub filesair_Click()

If Form20.p1 = Empty Then End
If Form20.p1 = Empty Or Form21.Data1.Recordset("Projeto") <> Form20.p1 Then
respsv = MsgBox(" Deseja salvar o projeto antes de sair?", vbYesNoCancel)

            If respsv = vbYes Then
                Form20.Show
                Form20.Command2.Visible = True
                Exit Sub
            End If
 
            If respsv = vbNo Then
            End
            End If
    
End If


If Form21.Data1.Recordset("Projeto") = Form20.p1 Then

respsv = MsgBox(" Deseja atualizar ou renomear o projeto antes de sair?", vbYesNoCancel)

        If respsv = vbYes Then
            Form20.Show
            Form20.Command2.Visible = True
            Exit Sub
        End If
 
            If respsv = vbNo Then
                End
            End If
End If

End Sub

Private Sub filesalva1_Click()
If Form20.p1 = Empty Then
    Form20.Visible = True
    Exit Sub
End If

If Form20.p1 <> Empty Then
 Form21.Show
 Form21.Command1.Visible = False
 Form21.Command4.Visible = True
 Form21.Command2.Visible = False
End If
End Sub

Private Sub Form_Activate()


Exit Sub

On Error GoTo Trata
    Set TB = DB.OpenRecordset("Acesso")
    If TB.RecordCount > 0 Then
       TB.MoveLast
       LIBER = TB("tipo")
       TPsof = TB("Tps")
       If TB("Tipo") <> "KFAR0852" Then
             If TB("DtAce") > Date Then
                MsgBox "Programa já utilizado em uma data Posterior a de Hoje", vbCritical, "Erro de Acesso"
                Inise = Val(Minute(Time)) + Val(Second(Time)) + Val(Date)
                Div# = IIf(Right(Inise, 1) <> 0, Right(Inise, 1), 1)
                SenhaC = Format(Mid(Int((Inise / Div#) + Inise), 1, 4), "0000")
                Inise = Format$(Inise, "0000")
                Inise = Trim(Trim(Mid(VarDisc, 1, 9)) + Trim(Inise))
                Sen = InputBox("Digite senha para Liberacao" + Chr(13) + Inise, "Senha de Liberação")
                
                If SenhaC <> Mid(Sen, 1, 4) Then MsgBox "Você não está Autorizado", 16, "Erro de Acesso": End
                Data2.RecordSource = ("Select * from Acesso")
                Data2.Refresh
                Data2.Recordset.MoveLast
                Dta = Data2.Recordset("Data")
                Tipo1 = Data2.Recordset("Tipo")
                Data2.Recordset.AddNew
                Data2.Recordset("Data") = Dta
                Data2.Recordset("DtAce") = Date
                Data2.Recordset("Tipo") = Tipo1
                Data2.Recordset("Tps") = IIf(IsNumeric(Mid(Sen, 15, 1)), Mid(Sen, 15, 1), 0)
                Data2.Recordset.Update
                Data2.Refresh
                'End
             End If
             Data2.RecordSource = ("Select * from Acesso")
             Data2.Refresh
             Data2.Recordset.MoveLast
             Data2.Recordset.Edit
             Data2.Recordset("Data") = Data2.Recordset("Data")
             Data2.Recordset("DtAce") = Date
             Data2.Recordset("Tipo") = Data2.Recordset("Tipo")
             Data2.Recordset("Tps") = IIf(IsNumeric(Mid(Sen, 15, 1)), Mid(Sen, 15, 1), 0)
             Data2.Recordset.Update
             Data2.Refresh
             Set TB = DB.OpenRecordset("Acesso")
             TB.MoveLast
             
             If Date >= TB("Data") Then
                Inise = Val(Minute(Time)) + Val(Second(Time)) + Val(Date)
                Div# = IIf(Right(Inise, 1) <> 0, Right(Inise, 1), 1)
                SenhaC = Format(Mid(Int((Inise / Div#) + Inise), 1, 4), "0000")
                Inise = Format$(Inise, "0000")
                Inise = Trim(Trim(Mid(VarDisc, 1, 9)) + Trim(Inise))
                Sen = InputBox("Digite senha para Liberacao" + Chr(13) + Inise, "Senha de Liberação")
                
                If SenhaC <> Mid(Sen, 1, 4) Then
                    MsgBox "Voce não esta Autorizado": End
                Else
                    Data2.RecordSource = ("Select * from Acesso")
                    Data2.Refresh
                    Data2.Recordset.AddNew
                    LLI = Date + Val(Mid(Sen, 5, 2))
                    Data2.Recordset("Data") = LLI
                    Data2.Recordset("DtAce") = Date
                    Data2.Recordset("Tipo") = Mid(Sen, 7, 8)
                    Data2.Recordset("Tps") = IIf(IsNumeric(Mid(Sen, 15, 1)), Mid(Sen, 15, 1), 0)
                    Data2.Recordset.Update
                    Data2.Refresh
                End If
             End If
       End If
    Else
        MsgBox "Consulte o seu fornecedor !", 16, "Erro de sistema": End: Exit Sub
    End If
    Exit Sub
Trata:
    Select Case Err
        Case 3078
            DB.Execute "CREATE TABLE Acesso " _
            & "(Data Date, " _
            & "DtAce Date, " _
            & "Tps integer, " _
            & "Tipo Text(10));"
            
            DFim = Date + 15
            Data2.DatabaseName = SR
            Data2.RecordSource = ("Select * from Acesso")
            Data2.Refresh
            Data2.Recordset.AddNew
            Data2.Recordset("Data") = DFim
            Data2.Recordset("DtAce") = Date
            Data2.Recordset("Tipo") = "DEMO"
            Data2.Recordset("Tps") = 0
            Data2.Recordset.Update
            Data2.Refresh
    End Select
End Sub

Private Sub Form_Initialize()
    Call Acesso
    Sair = 0
    Set WR = Workspaces(0)
    'Set DB = OpenDatabase("C:\Arquivos de programas\MEavcd\Alvaro.mdb", True, False, ";PWD=AF33216028")
    'SR = "C:\Arquivos de programas\MEavcd\Alvaro.mdb"
    SR = App.Path + "\Alvaro.mdb"
    Set DB = OpenDatabase(SR, True, False, ";PWD=AF33216028")
End Sub

Private Sub Form_Load()

Data2.DatabaseName = SR


hsfuse.Value = 100
fuse = 0

hslargo = 100
largo = 0

hsalto = 100
alto = 0

hsvolumeht = 100
volumeht = 0

hsvolumevt = 100
volumevt = 0

hsnariz = 100
nariz = 0

hspeso = 100
peso = 0

hsmargem = 50
margestat = 0

alterapeso.Value = Unchecked

'ABRINDO O CADASTRO DE PROJETOS ANTES DO FORM!
'Form1.Visible = False
'Form21.Visible = True

Form1.Visible = False
Form25.Visible = True
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub hsalto_Change()
If hsalto.Value < 100 Then alto = -1 * (100 - hsalto.Value)
    
If hsalto.Value > 100 Then alto = 1 + (hsalto.Value - 100)
    
If hsalto.Value = 100 Then alto = 0

valfuseH = hsalto.Value
    
End Sub


Private Sub HScroll1_Change()
If HScroll1.Value = 0 Then
altitude = 0
roar = 1.225
'dar1 = 0.00238
End If

If HScroll1.Value = 1 Then
altitude = 500
roar = 1.1673
'dar1 = 0.00226
End If

If HScroll1.Value = 2 Then
altitude = 1000
roar = 1.1117
'dar1 = 0.00215
End If

If HScroll1.Value = 3 Then
altitude = 1500
roar = 1.0581
'dar1 = 0.00205
End If

If HScroll1.Value = 4 Then
altitude = 2000
roar = 1.0065
'dar1 = 0.00195
End If

If HScroll1.Value = 5 Then
altitude = 2500
roar = 0.9569
'dar1 = 0.00185
End If

If HScroll1.Value = 6 Then
altitude = 3000
roar = 0.9091
'dar1 = 0.00176
End If

If HScroll1.Value = 7 Then
altitude = 3500
roar = 0.8632
'dar1 = 0.00167
End If

If HScroll1.Value = 8 Then
altitude = 4000
roar = 0.8191
'dar1 = 0.00159
End If

If HScroll1.Value = 9 Then
altitude = 4500
roar = 0.7768
'dar1 = 0.00151
End If

If HScroll1.Value = 10 Then
altitude = 5000
roar = 0.7351
'dar1 = 0.00142
End If
End Sub

Private Sub hsfuse_Change()
If hsfuse.Value < 100 Then fuse = -1 * (100 - hsfuse.Value)
    
If hsfuse.Value > 100 Then fuse = 1 + (hsfuse.Value - 100)
    
If hsfuse.Value = 100 Then fuse = 0

valfuse = hsfuse.Value
End Sub

Private Sub hslargo_Change()
If hslargo.Value < 99 Then largo = -1 * (100 - hslargo.Value)
   
If hslargo.Value > 99 Then largo = 1 + (hslargo.Value - 100)
   
If hslargo.Value = 99 Then largo = 0

valfuseL = hslargo.Value
   
End Sub

Private Sub hsmargem_Change()
If hsmargem.Value < 99 Then margestat = -1 * (99 - hsmargem.Value)
   
If hsmargem.Value > 99 Then margestat = 1 + (hsmargem.Value - 100)
   
If hsmargem.Value = 99 Then margestat = 0

Form1.margestat = Form1.hsmargem.Value

marg = hsmargem.Value
End Sub

Private Sub hsnariz_Change()
If hsnariz.Value < 99 Then nariz = -1 * (99 - hsnariz.Value)
   
If hsnariz.Value > 99 Then nariz = 1 + (hsnariz.Value - 100)
   
If hsnariz.Value = 99 Then nariz = 0

distanasa = hsnariz.Value
End Sub

'Private Sub hspeso_Change()
'If hspeso.Value < 99 Then peso = -1 * (99 - hspeso.Value)
   
'If hspeso.Value > 99 Then peso = 1 + (hspeso.Value - 100)
   
'If hspeso.Value = 99 Then peso = 0
   
'If peso < 0 Then kpeso = ((100 + peso) / 100)
'If peso > 0 Then kpeso = 1 + (peso / 100)
'If peso = 0 Then kpeso = 1
 
'wo = woanexo * kpeso
'wo = Format(wo, "0.00")

'Form1.alterapeso.Value = Checked
'Form1.alterapeso.Caption = "Alterar"
'Form1.alterapeso.Enabled = True

'pesosalva = hspeso.Value
'End Sub

Private Sub hsvolumeht_Change()
If hsvolumeht.Value < 99 Then volumeht = -1 * (99 - hsvolumeht.Value)
   
If hsvolumeht.Value > 99 Then volumeht = 1 + (hsvolumeht.Value - 100)
   
If hsvolumeht.Value = 99 Then volumeht = 0
  
volempemH = hsvolumeht.Value
End Sub

Private Sub hsvolumevt_Change()
If hsvolumevt.Value < 99 Then volumevt = -1 * (99 - hsvolumevt.Value)
   
If hsvolumevt.Value > 99 Then volumevt = 1 + (hsvolumevt.Value - 100)
   
If hsvolumevt.Value = 99 Then volumevt = 0

volempemV = hsvolumevt.Value
   
End Sub

Private Sub hsvstol_Change()
If hsvstol.Value < 50 Then vstol = Int(vstolanexo * ((2 * hsvstol.Value) - 1) / 100)
   
If hsvstol.Value > 50 Then vstol = Int(vstolanexo * (1 + (2 * hsvstol.Value - 100) / 100))
   
If hsvstol.Value = 50 Then vstol = Int(vstolanexo * 1)

vstol = Format(vstol, "0.00")

stallsalva = hsvstol.Value

Command15.Visible = False
End Sub

Private Sub PCA_Click()
If Form21.Data1.Recordset.RecordCount = 0 Then GoTo 100

If Form21.Data1.Recordset("projeto") = Form20.p1 Then Exit Sub

If Form21.Data1.Recordset("projeto") <> Form20.p1 Then GoTo 100

100:
Form1.hsmargem = 15

Form1.Command3.Enabled = True

' Lançando o Peso PCA em anexo

End Sub


Private Sub posialta_Click()
Form1.Command5.Enabled = True

End Sub

Private Sub posibaixa_Click()
Command5.Enabled = True
End Sub

Private Sub posimedia_Click()
Command5.Enabled = True
End Sub

Private Sub tipoacrof3a_Click()
If Form21.Data1.Recordset.RecordCount = 0 Then GoTo 100

If Form21.Data1.Recordset("projeto") = Form20.p1 Then Exit Sub

If Form21.Data1.Recordset("projeto") <> Form20.p1 Then GoTo 100

100:
Form1.hsmargem = 15

Command3.Enabled = True
End Sub

Private Sub tipoacroili_Click()
If Form21.Data1.Recordset.RecordCount = 0 Then GoTo 100

If Form21.Data1.Recordset("projeto") = Form20.p1 Then Exit Sub

If Form21.Data1.Recordset("projeto") <> Form20.p1 Then GoTo 100

100:
Form1.hsmargem = 15

Command3.Enabled = True
End Sub



Private Sub tipotreino_Click()
If Form21.Data1.Recordset.RecordCount = 0 Then GoTo 100

If Form21.Data1.Recordset("projeto") = Form20.p1 Then Exit Sub

If Form21.Data1.Recordset("projeto") <> Form20.p1 Then GoTo 100

100:
Form1.hsmargem = 15

Form1.Command3.Enabled = True

End Sub

Private Sub rotac_Change()
RPM = rotac.Value
End Sub

Private Sub tremconv_Click()
Command8.Enabled = True
Call Command8_Click
Form1.VScroll1.Value = 5
Form1.VScroll2.Value = 0
End Sub

Private Sub tremtric_Click()
Call Command8_Click
Form1.VScroll1.Value = 5
Form1.VScroll2.Value = 5
End Sub

Private Sub VScroll1_Change()
y1 = 1 + Form1.VScroll1 / 10
End Sub

Private Sub VScroll2_Change()
y2 = 1 + Form1.VScroll2 / 10
End Sub

