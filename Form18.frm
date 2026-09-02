VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form18 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Esboço"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form18.frx":0000
   LinkTopic       =   "Form18"
   MouseIcon       =   "Form18.frx":030A
   MousePointer    =   2  'Cross
   ScaleHeight     =   8280
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame22 
      Height          =   1335
      Left            =   0
      MouseIcon       =   "Form18.frx":045C
      MousePointer    =   99  'Custom
      TabIndex        =   71
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame23 
         Caption         =   "Estabilizador vertical"
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   5640
         TabIndex        =   79
         Top             =   120
         Width           =   5295
         Begin VB.TextBox gradvt 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1920
            TabIndex        =   81
            Text            =   "0,00"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox vt 
            Alignment       =   2  'Center
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4440
            TabIndex        =   92
            Text            =   "0,00"
            Top             =   240
            Width           =   495
         End
         Begin VB.HScrollBar HScroll4 
            Enabled         =   0   'False
            Height          =   255
            Left            =   2760
            Max             =   7
            TabIndex        =   91
            Top             =   240
            Value           =   7
            Width           =   1575
         End
         Begin VB.PictureBox Picture13 
            Height          =   375
            Left            =   120
            ScaleHeight     =   315
            ScaleWidth      =   4875
            TabIndex        =   88
            Top             =   600
            Width           =   4935
            Begin VB.CommandButton Command34 
               Caption         =   "Configur&ar"
               Height          =   315
               Left            =   0
               MouseIcon       =   "Form18.frx":05AE
               MousePointer    =   99  'Custom
               Style           =   1  'Graphical
               TabIndex        =   89
               ToolTipText     =   "Configurar o estabilizador vertical"
               Top             =   0
               Width           =   4875
            End
         End
         Begin VB.HScrollBar HScroll3 
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            Max             =   50
            TabIndex        =   80
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox hs4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4440
            TabIndex        =   106
            Text            =   "0,00"
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Frame Frame19 
         Height          =   1095
         Left            =   11040
         TabIndex        =   77
         Top             =   120
         Width           =   735
         Begin VB.CommandButton Command27 
            Caption         =   "OK !"
            Height          =   795
            Left            =   120
            MouseIcon       =   "Form18.frx":08B8
            MousePointer    =   99  'Custom
            TabIndex        =   78
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000000&
         Caption         =   "Asa e estabiizadoe horizontal"
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   120
         MouseIcon       =   "Form18.frx":0BC2
         MousePointer    =   99  'Custom
         TabIndex        =   72
         Top             =   120
         Width           =   5415
         Begin VB.HScrollBar HScroll2 
            Enabled         =   0   'False
            Height          =   255
            Left            =   2760
            Max             =   100
            MouseIcon       =   "Form18.frx":0D14
            MousePointer    =   99  'Custom
            SmallChange     =   10
            TabIndex        =   84
            Top             =   240
            Value           =   1
            Width           =   1575
         End
         Begin VB.PictureBox Picture2 
            Height          =   375
            Left            =   240
            ScaleHeight     =   315
            ScaleWidth      =   4875
            TabIndex        =   86
            Top             =   625
            Width           =   4935
            Begin VB.CommandButton Command28 
               Caption         =   "Confi&gurar"
               Height          =   315
               Left            =   0
               MouseIcon       =   "Form18.frx":101E
               MousePointer    =   99  'Custom
               TabIndex        =   87
               ToolTipText     =   "Configurar a asa e o estabilizador horizontal"
               Top             =   0
               Width           =   4875
            End
         End
         Begin VB.TextBox afilaHT 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3480
            TabIndex        =   85
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox grauht 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   83
            Text            =   "0,00"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox grau 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   75
            Text            =   "0,00"
            Top             =   240
            Width           =   495
         End
         Begin VB.HScrollBar HScroll1 
            Enabled         =   0   'False
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   100
            MouseIcon       =   "Form18.frx":1328
            MousePointer    =   99  'Custom
            SmallChange     =   10
            TabIndex        =   74
            Top             =   240
            Value           =   1
            Width           =   1575
         End
         Begin VB.TextBox afilaASA 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   480
            TabIndex        =   73
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox hsasaval 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            TabIndex        =   93
            Text            =   "0,00"
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox hshtval 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2880
            TabIndex        =   94
            Text            =   "0,00"
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CheckBox outros 
            Caption         =   "salvar outros desenhos"
            Height          =   255
            Left            =   240
            TabIndex        =   107
            Top             =   600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "º"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5040
            TabIndex        =   82
            Top             =   240
            Width           =   60
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "º"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2520
            TabIndex        =   76
            Top             =   240
            Width           =   60
         End
      End
   End
   Begin VB.Frame Frame6 
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame10 
         Caption         =   "&Vertical"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   9120
         MousePointer    =   1  'Arrow
         TabIndex        =   67
         Top             =   120
         Width           =   1815
         Begin VB.PictureBox Picture4 
            Height          =   735
            Left            =   1320
            ScaleHeight     =   675
            ScaleWidth      =   255
            TabIndex        =   69
            Top             =   240
            Width           =   315
            Begin VB.VScrollBar vslinha 
               Height          =   675
               LargeChange     =   500
               Left            =   0
               Max             =   8000
               Min             =   1500
               MouseIcon       =   "Form18.frx":1632
               MousePointer    =   99  'Custom
               SmallChange     =   500
               TabIndex        =   70
               Top             =   0
               Value           =   3000
               Width           =   255
            End
         End
         Begin VB.TextBox posilinha 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H80000000&
         Caption         =   "&Expandir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   6720
         MousePointer    =   1  'Arrow
         TabIndex        =   63
         Top             =   120
         Width           =   2255
         Begin VB.PictureBox Picture5 
            Height          =   275
            Left            =   1200
            ScaleHeight     =   210
            ScaleWidth      =   795
            TabIndex        =   65
            Top             =   240
            Width           =   855
            Begin VB.HScrollBar vsescala 
               Height          =   215
               LargeChange     =   10
               Left            =   0
               Max             =   32000
               Min             =   10
               MouseIcon       =   "Form18.frx":193C
               MousePointer    =   99  'Custom
               TabIndex        =   66
               Top             =   0
               Value           =   10
               Width           =   795
            End
         End
         Begin VB.TextBox esc18 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   64
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Vistas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   120
         MouseIcon       =   "Form18.frx":1C46
         MousePointer    =   99  'Custom
         TabIndex        =   58
         Top             =   120
         Width           =   3135
         Begin VB.OptionButton Option1 
            Caption         =   "Lateral"
            Height          =   195
            Left            =   240
            TabIndex        =   62
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Planta"
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   720
            Width           =   855
         End
         Begin VB.PictureBox Picture12 
            Height          =   735
            Left            =   1320
            ScaleHeight     =   675
            ScaleWidth      =   1575
            TabIndex        =   59
            Top             =   240
            Width           =   1635
            Begin VB.CommandButton Command32 
               Height          =   675
               Left            =   0
               MouseIcon       =   "Form18.frx":1D98
               MousePointer    =   99  'Custom
               Picture         =   "Form18.frx":20A2
               Style           =   1  'Graphical
               TabIndex        =   90
               ToolTipText     =   "Vista lateral"
               Top             =   0
               Width           =   1575
            End
            Begin VB.CommandButton Command33 
               Height          =   675
               Left            =   0
               MouseIcon       =   "Form18.frx":4C04
               MousePointer    =   99  'Custom
               Picture         =   "Form18.frx":4F0E
               Style           =   1  'Graphical
               TabIndex        =   60
               ToolTipText     =   "Vista de topo ou planta"
               Top             =   0
               Visible         =   0   'False
               Width           =   1575
            End
         End
      End
      Begin VB.Frame Frame18 
         Height          =   1095
         Left            =   11040
         TabIndex        =   56
         Top             =   120
         Width           =   735
         Begin VB.CommandButton Command29 
            Caption         =   "OK !"
            Height          =   735
            Left            =   120
            MouseIcon       =   "Form18.frx":795C
            MousePointer    =   99  'Custom
            TabIndex        =   57
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture10 
         Height          =   375
         Left            =   6720
         MouseIcon       =   "Form18.frx":7C66
         MousePointer    =   99  'Custom
         ScaleHeight     =   315
         ScaleWidth      =   2175
         TabIndex        =   54
         Top             =   840
         Width           =   2235
         Begin VB.CommandButton Command30 
            Caption         =   "&Abrir"
            Enabled         =   0   'False
            Height          =   315
            Left            =   0
            MouseIcon       =   "Form18.frx":7DB8
            MousePointer    =   99  'Custom
            TabIndex        =   55
            ToolTipText     =   "Modificar a expanção"
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "O que é?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   3360
         MouseIcon       =   "Form18.frx":80C2
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   120
         Width           =   3255
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expandir : Amplia o desenho na tela."
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   2580
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Vertical : Posição do desenho na tela."
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   720
            Width           =   2685
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   1335
      Left            =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   43
      Top             =   -120
      Visible         =   0   'False
      Width           =   11895
      Begin VB.PictureBox Picture14 
         Height          =   735
         Left            =   10560
         ScaleHeight     =   675
         ScaleWidth      =   195
         TabIndex        =   95
         Top             =   360
         Width           =   255
         Begin VB.VScrollBar VScroll3 
            Enabled         =   0   'False
            Height          =   675
            Left            =   0
            Max             =   7
            Min             =   3
            TabIndex        =   96
            Top             =   0
            Value           =   3
            Width           =   195
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000000&
         Caption         =   "Inserir no desenho"
         Height          =   855
         Left            =   120
         MouseIcon       =   "Form18.frx":8214
         MousePointer    =   1  'Arrow
         TabIndex        =   46
         Top             =   240
         Width           =   4695
         Begin VB.PictureBox Picture6 
            Height          =   435
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   4335
            TabIndex        =   47
            Top             =   240
            Width           =   4400
            Begin VB.CommandButton Command6 
               BackColor       =   &H80000000&
               Caption         =   "CG componentes"
               Height          =   375
               Left            =   2880
               Style           =   1  'Graphical
               TabIndex        =   108
               ToolTipText     =   "Inserir superfícies"
               Top             =   0
               Width           =   1450
            End
            Begin VB.CommandButton Command10 
               BackColor       =   &H80000000&
               Caption         =   "Rodas"
               Enabled         =   0   'False
               Height          =   375
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   49
               ToolTipText     =   "Inserir rodas"
               Top             =   0
               Width           =   1450
            End
            Begin VB.CommandButton Command5 
               BackColor       =   &H80000000&
               Caption         =   "Superfícies"
               Enabled         =   0   'False
               Height          =   375
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   48
               ToolTipText     =   "Inserir superfícies"
               Top             =   0
               Width           =   1450
            End
         End
      End
      Begin VB.Frame Frame15 
         Height          =   1125
         Left            =   11040
         TabIndex        =   44
         Top             =   120
         Width           =   735
         Begin VB.CommandButton Command25 
            Caption         =   "OK !"
            Height          =   735
            Left            =   120
            MouseIcon       =   "Form18.frx":8366
            MousePointer    =   99  'Custom
            TabIndex        =   45
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Asa média"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   9000
         TabIndex        =   97
         Top             =   120
         Width           =   1935
         Begin VB.OptionButton Option3 
            Caption         =   "Abrir"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox VSasamed 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            TabIndex        =   101
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox hasamed 
            Alignment       =   2  'Center
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   99
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox valasa 
            Alignment       =   2  'Center
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   98
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cm."
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
            Left            =   840
            TabIndex        =   100
            Top             =   720
            Width           =   270
         End
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   2
         X1              =   8400
         X2              =   8880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outros"
         Height          =   195
         Left            =   7200
         TabIndex        =   116
         Top             =   960
         Width           =   465
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   8400
         X2              =   8880
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fuselagem"
         Height          =   195
         Left            =   7200
         TabIndex        =   115
         Top             =   720
         Width           =   765
      End
      Begin VB.Line Line6 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         X1              =   8400
         X2              =   8880
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Combustivel"
         Height          =   195
         Left            =   7200
         TabIndex        =   114
         Top             =   480
         Width           =   855
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   8400
         X2              =   8880
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrumentos"
         Height          =   195
         Left            =   7200
         TabIndex        =   113
         Top             =   240
         Width           =   900
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   6480
         X2              =   6960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bagagem"
         Height          =   195
         Left            =   5040
         TabIndex        =   112
         Top             =   960
         Width           =   675
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         X1              =   6480
         X2              =   6960
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ocupantes"
         Height          =   195
         Left            =   5040
         TabIndex        =   111
         Top             =   720
         Width           =   780
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF00FF&
         BorderWidth     =   2
         X1              =   6480
         X2              =   6960
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Piloto e copiloto"
         Height          =   195
         Left            =   5040
         TabIndex        =   110
         Top             =   480
         Width           =   1125
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   6480
         X2              =   6960
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motor"
         Height          =   195
         Left            =   5040
         TabIndex        =   109
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.Frame Frame11 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   22
      Top             =   7560
      Width           =   11895
      Begin VB.Frame Frame24 
         Caption         =   "Nervuras (metros)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   6000
         TabIndex        =   103
         Top             =   0
         Width           =   1815
         Begin VB.TextBox nnerv 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   104
            Text            =   "0,00"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Espaçamento"
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000000&
         Caption         =   "Arco "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   7920
         MousePointer    =   1  'Arrow
         TabIndex        =   35
         Top             =   0
         Width           =   3975
         Begin VB.TextBox raio 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   720
            TabIndex        =   39
            Text            =   "0,000"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox alfa 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   38
            Text            =   "0,000"
            Top             =   240
            Width           =   615
         End
         Begin VB.PictureBox Picture8 
            Height          =   405
            Left            =   3100
            ScaleHeight     =   345
            ScaleWidth      =   720
            TabIndex        =   36
            Top             =   165
            Width           =   780
            Begin VB.CommandButton Command19 
               Caption         =   "0K!"
               Height          =   345
               Left            =   0
               MouseIcon       =   "Form18.frx":8670
               MousePointer    =   99  'Custom
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   0
               Width           =   720
            End
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "º"
            Height          =   195
            Left            =   2760
            TabIndex        =   42
            Top             =   240
            Width           =   60
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Raio                  m."
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "a"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   1800
            TabIndex        =   40
            Top             =   240
            Width           =   120
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000000&
         Caption         =   "Cavernas (entre metros)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   28
         Top             =   0
         Width           =   3135
         Begin VB.TextBox base 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   480
            TabIndex        =   32
            Text            =   "0,000"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox altura 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   31
            Text            =   "0,000"
            Top             =   240
            Width           =   615
         End
         Begin VB.PictureBox Picture11 
            Height          =   405
            Left            =   2250
            ScaleHeight     =   345
            ScaleWidth      =   720
            TabIndex        =   29
            Top             =   165
            Width           =   780
            Begin VB.CommandButton Command31 
               Caption         =   "0K!"
               Height          =   345
               Left            =   0
               MouseIcon       =   "Form18.frx":897A
               MousePointer    =   99  'Custom
               TabIndex        =   30
               Top             =   0
               Width           =   720
            End
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "b :"
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   180
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "h :"
            Height          =   195
            Left            =   1200
            TabIndex        =   33
            Top             =   240
            Width           =   180
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         Caption         =   "Linhas retas (metro)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   2655
         Begin VB.TextBox xx 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   600
            TabIndex        =   25
            Text            =   "0,000"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox yy 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1920
            TabIndex        =   24
            Text            =   "0,000"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y="
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1560
            TabIndex        =   27
            Top             =   285
            Width           =   285
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X="
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   285
            Width           =   285
         End
      End
   End
   Begin VB.Frame Frame12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   735
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00C0C0C0&
         Height          =   6165
         Left            =   60
         ScaleHeight     =   6105
         ScaleWidth      =   555
         TabIndex        =   9
         Top             =   1355
         Width           =   615
         Begin VB.CommandButton Command1 
            Height          =   475
            Left            =   0
            MouseIcon       =   "Form18.frx":8C84
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":8F8E
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Linhas elásticas"
            Top             =   0
            Width           =   555
         End
         Begin VB.CommandButton Command20 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   0
            MouseIcon       =   "Form18.frx":9840
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":9B4A
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Linha reta em X"
            Top             =   480
            Width           =   555
         End
         Begin VB.CommandButton Command22 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   0
            MouseIcon       =   "Form18.frx":A3FC
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":A706
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Linha reta em Y"
            Top             =   960
            Width           =   555
         End
         Begin VB.CommandButton Command3 
            Height          =   480
            Left            =   0
            MouseIcon       =   "Form18.frx":AFB8
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":B2C2
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Pontilhada"
            Top             =   1440
            Width           =   555
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0C0&
            Height          =   480
            Left            =   0
            MouseIcon       =   "Form18.frx":BB74
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":BE7E
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Tracejada"
            Top             =   1920
            Width           =   555
         End
         Begin VB.CommandButton Command14 
            Height          =   480
            Left            =   0
            MouseIcon       =   "Form18.frx":C730
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":CA3A
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Elipse"
            Top             =   2400
            Width           =   555
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   0
            MouseIcon       =   "Form18.frx":D2EC
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":D5F6
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Circulo"
            Top             =   2880
            Width           =   555
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H80000000&
            Height          =   480
            Left            =   0
            MouseIcon       =   "Form18.frx":DEA8
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":E1B2
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Retangulo"
            Top             =   3360
            Width           =   555
         End
         Begin VB.CommandButton Command21 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   0
            MouseIcon       =   "Form18.frx":EA64
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":ED6E
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Texto"
            Top             =   3840
            Width           =   555
         End
         Begin VB.CommandButton Command23 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            MouseIcon       =   "Form18.frx":F620
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":F92A
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Dimensão em X"
            Top             =   4320
            Width           =   555
         End
         Begin VB.CommandButton Command24 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            MouseIcon       =   "Form18.frx":101DC
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":104E6
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Dimensão em Y"
            Top             =   4920
            Width           =   555
         End
         Begin VB.CommandButton apaga 
            Height          =   575
            Left            =   0
            MouseIcon       =   "Form18.frx":10D98
            MousePointer    =   99  'Custom
            Picture         =   "Form18.frx":110A2
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Apaga pontos."
            Top             =   5535
            Width           =   555
         End
      End
      Begin VB.Frame Frame13 
         Height          =   1215
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   495
         Begin VB.CommandButton Command26 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            MouseIcon       =   "Form18.frx":11954
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   240
            Width           =   255
         End
      End
   End
   Begin VB.TextBox med2 
      BorderStyle     =   0  'None
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
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Text            =   "dy"
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox med1 
      BorderStyle     =   0  'None
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
      Height          =   195
      Left            =   11280
      TabIndex        =   4
      Text            =   "dx"
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox label2 
      BorderStyle     =   0  'None
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
      Left            =   11280
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "X"
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox label3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Y"
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   10560
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line bfvt 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   10440
      X2              =   9840
      Y1              =   4200
      Y2              =   5040
   End
   Begin VB.Line bavt 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   8760
      X2              =   9480
      Y1              =   5040
      Y2              =   4200
   End
   Begin VB.Line raizht 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   840
      X2              =   1440
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line medht 
      BorderColor     =   &H00404040&
      Visible         =   0   'False
      X1              =   840
      X2              =   1440
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line raizasa 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   840
      X2              =   1440
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line medasa 
      BorderColor     =   &H00404040&
      Visible         =   0   'False
      X1              =   840
      X2              =   1320
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line pontht 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   840
      X2              =   1440
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line pontasa 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   840
      X2              =   1440
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line bfht 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   1200
      X2              =   1200
      Y1              =   5280
      Y2              =   1920
   End
   Begin VB.Line baht 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   1320
      X2              =   1320
      Y1              =   5280
      Y2              =   1920
   End
   Begin VB.Line bfasa 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   960
      X2              =   960
      Y1              =   5280
      Y2              =   1920
   End
   Begin VB.Line basa 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   840
      X2              =   840
      Y1              =   5280
      Y2              =   1920
   End
   Begin VB.Line linhaasa1 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   5760
      Y2              =   4320
   End
   Begin VB.Line linhaleme 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   0
      X2              =   11880
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line linhafinal 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   4320
      Y2              =   5760
   End
   Begin VB.Line linhasf 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   0
      X2              =   11880
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line linhaif 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   0
      X2              =   12000
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line linhaY 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   0
      Y2              =   8400
   End
   Begin VB.Line linhaX 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   0
      X2              =   11880
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   10560
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu fileMenu 
      Caption         =   "&Iniciar"
      Begin VB.Menu filevista 
         Caption         =   "&Vista"
         Begin VB.Menu filelateral 
            Caption         =   "Lateral"
            Shortcut        =   {F3}
         End
         Begin VB.Menu fileplanta 
            Caption         =   "Em planta"
            Shortcut        =   {F4}
         End
         Begin VB.Menu menusepara 
            Caption         =   "-"
         End
         Begin VB.Menu filesuperficie 
            Caption         =   "Superfícies"
            Begin VB.Menu fileasa18 
               Caption         =   "Asa"
            End
            Begin VB.Menu fileEH18 
               Caption         =   "Estabilizador Horiz."
            End
         End
         Begin VB.Menu menusepar 
            Caption         =   "-"
         End
         Begin VB.Menu filetabservo 
            Caption         =   "&Tabela CG servos"
         End
      End
      Begin VB.Menu FileNew 
         Caption         =   "Novo desenho"
      End
      Begin VB.Menu salvproj18 
         Caption         =   "Salvar projeto"
      End
      Begin VB.Menu word 
         Caption         =   "Salvar Arquivo Texto - WORD"
      End
      Begin VB.Menu salvdes 
         Caption         =   "Salvar desenhos"
         Begin VB.Menu filesave 
            Caption         =   "Desenho principal"
         End
         Begin VB.Menu filesaveas 
            Caption         =   "Outros desenhos"
         End
      End
      Begin VB.Menu Fileab 
         Caption         =   "Abrir desenhos"
         Begin VB.Menu fileopen 
            Caption         =   "Desenho principal"
         End
         Begin VB.Menu fileopout 
            Caption         =   "Outros desenhos"
         End
      End
      Begin VB.Menu fileir18 
         Caption         =   "I&r para ..."
         Begin VB.Menu filetrim 
            Caption         =   "Trimagem"
         End
         Begin VB.Menu filePB 
            Caption         =   "Peso && Balanceamento"
         End
         Begin VB.Menu fileestrut18 
            Caption         =   "Estrutura"
         End
         Begin VB.Menu fileestab18 
            Caption         =   "Estabilidade estática"
         End
         Begin VB.Menu fileperform18 
            Caption         =   "Performance"
         End
         Begin VB.Menu fileaerodin18 
            Caption         =   "Aerodinâmica"
         End
         Begin VB.Menu filedimensiona18 
            Caption         =   "Dimensionamento"
         End
         Begin VB.Menu fileprincipia18 
            Caption         =   "Página inicial"
         End
      End
   End
   Begin VB.Menu editmenu 
      Caption         =   "E&ditar"
      Begin VB.Menu retmenu 
         Caption         =   "Retirar menús"
      End
      Begin VB.Menu filelinhaux 
         Caption         =   "Retirar linhas auxiliares"
         Shortcut        =   {F2}
      End
      Begin VB.Menu editcopy 
         Caption         =   "Copiar"
      End
      Begin VB.Menu editcut 
         Caption         =   "Recortar"
         Shortcut        =   {F1}
      End
      Begin VB.Menu editpaste 
         Caption         =   "Colar"
      End
      Begin VB.Menu editclear 
         Caption         =   "Limpar tudo"
      End
   End
   Begin VB.Menu barrafer 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "&Calculadora && Conversor"
         Shortcut        =   {F5}
      End
      Begin VB.Menu filecofigura 
         Caption         =   "Configuração"
         Shortcut        =   {F6}
      End
      Begin VB.Menu filepadrao 
         Caption         =   "Padronização"
         Shortcut        =   {F7}
      End
      Begin VB.Menu insereqip 
         Caption         =   "Projeto"
         Shortcut        =   {F8}
      End
      Begin VB.Menu inseridesenho 
         Caption         =   "Desenho"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu shapemenu 
      Caption         =   "F&ormas"
      Begin VB.Menu drawline 
         Caption         =   "&Linhas elásticas "
      End
      Begin VB.Menu drawcircle 
         Caption         =   "Círculo "
      End
      Begin VB.Menu drawbox 
         Caption         =   "Retangulos "
      End
      Begin VB.Menu drawelipse 
         Caption         =   "Elipse "
      End
   End
   Begin VB.Menu widthmenu 
      Caption         =   "&Espessuras"
      Begin VB.Menu width1 
         Caption         =   "Nº 1"
      End
      Begin VB.Menu width2 
         Caption         =   "Nº 2"
      End
      Begin VB.Menu width3 
         Caption         =   "Nº 3"
      End
   End
   Begin VB.Menu stylemenu 
      Caption         =   "Est&ilo"
      Begin VB.Menu stylesolid 
         Caption         =   "Sólido"
      End
      Begin VB.Menu styledash 
         Caption         =   "Tracejado"
      End
      Begin VB.Menu styledot 
         Caption         =   "Pontilhado"
      End
      Begin VB.Menu menuseparator 
         Caption         =   "-"
      End
      Begin VB.Menu stylefilled 
         Caption         =   "Forma sólida"
      End
   End
   Begin VB.Menu colormenu 
      Caption         =   "&Côres"
      Begin VB.Menu colorpage 
         Caption         =   "Da página"
      End
      Begin VB.Menu colorpen 
         Caption         =   "Da pena"
      End
      Begin VB.Menu colorfill 
         Caption         =   "Do preenchimento"
      End
   End
   Begin VB.Menu fileinserir 
      Caption         =   "&Inserir"
      Begin VB.Menu insereasa 
         Caption         =   "Asa e estabilizadores"
      End
      Begin VB.Menu filehelice 
         Caption         =   "Hélice"
      End
      Begin VB.Menu fileauxilio 
         Caption         =   "Linhas auxiliares da fuselagem"
      End
      Begin VB.Menu filedistancia 
         Caption         =   "Dimensão X"
      End
      Begin VB.Menu filedistanciaY 
         Caption         =   "Dimensão Y"
      End
      Begin VB.Menu drawtext 
         Caption         =   "Texto"
      End
   End
   Begin VB.Menu filelocaliza 
      Caption         =   "&Localizar"
      Begin VB.Menu filecoord 
         Caption         =   "Coordenadas"
      End
   End
   Begin VB.Menu interasa 
      Caption         =   "&Downwash"
      Begin VB.Menu donestol 
         Caption         =   "Em Veloc. estol"
         Enabled         =   0   'False
      End
      Begin VB.Menu dondecol 
         Caption         =   "Em Veloc. decolagem"
         Enabled         =   0   'False
      End
      Begin VB.Menu donsub 
         Caption         =   "Em Veloc. subida"
         Enabled         =   0   'False
      End
      Begin VB.Menu donvmax 
         Caption         =   "Em Veloc. máxima"
         Enabled         =   0   'False
      End
      Begin VB.Menu filelimpa 
         Caption         =   "Limpar todas as linhas"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu saidaparaf 
      Caption         =   "Paraf&uso"
      Begin VB.Menu filelinha 
         Caption         =   "Traçar linhas"
         Enabled         =   0   'False
      End
      Begin VB.Menu fileapagalin 
         Caption         =   "Apagar linhas"
         Enabled         =   0   'False
      End
      Begin VB.Menu filerecuper 
         Caption         =   "Calcular recuperação"
      End
   End
   Begin VB.Menu fileimprim 
      Caption         =   "Im&primir"
      Begin VB.Menu fileimprimi18 
         Caption         =   "&Projeto"
      End
      Begin VB.Menu fileimpdes 
         Caption         =   "&Desenho"
      End
   End
   Begin VB.Menu filesai18 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Shape As String
Dim XStart, YStart, XPrevious, YPrevious As Single
Dim CopyBMP, PasteBMP, CutBMP, PrintText As Integer
Dim PDrawWidth, PDrawStyle, PFillStyle As Integer
Dim CopyWidth, CopyHeight As Integer
Dim XLabel, YLabel As Integer

Dim OpenFile As String

Private Sub UnCheckStyles()
    
    stylesolid.Checked = False
    styledash.Checked = False
    styledot.Checked = False
    
End Sub
Private Sub alfa_GotFocus()
alfa.SelStart = 0
alfa.SelLength = Len(alfa)
End Sub
Private Sub alfa_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command19.SetFocus
End Sub

Private Sub altura_GotFocus()
altura.SelStart = 0
altura.SelLength = Len(altura)
End Sub
Private Sub altura_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command31.SetFocus
End Sub

Private Sub apaga_Click()
Shape = "apaga"
Form18.MousePointer = 99
End Sub

Private Sub base_GotFocus()
base.SelStart = 0
base.SelLength = Len(base)
End Sub
Private Sub base_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then altura.SetFocus
End Sub

Private Sub ColorFill_Click()
    
    CommonDialog1.Color = Form18.FillColor
    CommonDialog1.Flags = cdlCCRGBInit
    CommonDialog1.ShowColor
    Form18.FillColor = CommonDialog1.Color

End Sub

Private Sub ColorPage_Click()
Dim resp As String

resp = MsgBox("Se voce alterar a côr da página os desenhos já executados serão apagados", vbOKCancel)

If resp = vbOK Then
    CommonDialog1.Color = Form18.BackColor
    CommonDialog1.Flags = cdlCCRGBInit
    CommonDialog1.ShowColor
    Form18.BackColor = CommonDialog1.Color
    
        Label3.BackColor = Form18.BackColor
        Label2.BackColor = Form18.BackColor
        med2.BackColor = Form18.BackColor
        med1.BackColor = Form18.BackColor
 End If
 
 If resp = vbCancel Then Exit Sub

End Sub

Private Sub ColorPen_Click()
    
    CommonDialog1.Color = Form18.ForeColor
    CommonDialog1.Flags = cdlCCRGBInit
    CommonDialog1.ShowColor
    Form18.ForeColor = CommonDialog1.Color
    
End Sub
Private Sub Command1_Click()
Call DrawLine_Click
Call StyleSolid_Click
Form18.MousePointer = 2

End Sub


Private Sub Command10_Click()

Dim escalax, ytp, yhb, fatx, yprinc, rprinc, rbeq, Ybeq, yterra, estprinc, estbeq As Double

'xxxxxxxxxx DESENHANDO RODAS xxxxxxxxxxxxxxxxxxxx

'Fator de escala
fatx = Form18.esc18
escalax = 800 * fatx

estbeq = Form14.d13 / 100
estprinc = Form14.d14 / 100

'.............. Fim

'Trem principal
ytp = Form1.y1
yprinc = linhaif.y1 + ((ytp * Form1.diamtp / 100) - ((Form1.diamtp / 100) / 2)) * escalax
rprinc = ((Form1.diamtp / 2) / 100) * escalax
Form1.ForeColor = &H80000012

'Desenho da roda

If estprinc = Empty Then
MsgBox "Entre com a estação do trem principal", vbInformation
Exit Sub
End If

If estbeq = Empty Then
MsgBox "Entre com a estação da bequilha", vbInformation
Exit Sub
End If
'Desenho da roda do T. principal
Circle (linhaY.x1 + (estprinc * escalax), yprinc), rprinc, &H80000012
Circle (linhaY.x1 + (estprinc * escalax), yprinc), 0.5 * rprinc, &H80000012
Circle (linhaY.x1 + (estprinc * escalax), yprinc), 0.4 * rprinc, &H80000012
Circle (linhaY.x1 + (estprinc * escalax), yprinc), 0.05 * rprinc, &H80000012

'...... FIM Principal ............

'Bequilha

If estprinc = Empty Then
MsgBox "Entre com a estação"
Exit Sub
End If

If Form1.tremconv Then
yhb = Form1.y2 ^ 2.5
Ybeq = (linhaif.y1 - (((Form1.diambeq / 100) * escalax) / 2)) + (((yhb * Form1.diambeq / 100) / 8) * escalax)
End If

If Form1.tremtric Then
yhb = Form1.y2
Ybeq = linhaif.y1 + ((yhb * Form1.diamtp / 100) - ((Form1.diambeq / 100) / 2)) * escalax
End If

rbeq = ((Form1.diambeq / 2) / 100) * escalax
Form1.ForeColor = &H80000012

If estprinc = Empty Then
MsgBox "Entre com a estação"
Exit Sub
End If
'Desenho da roda da bequilha
Circle (linhaY.x1 + (estbeq * escalax), Ybeq), rbeq, &H80000012
Circle (linhaY.x1 + (estbeq * escalax), Ybeq), 0.5 * rbeq, &H80000012
Circle (linhaY.x1 + (estbeq * escalax), Ybeq), 0.4 * rbeq, &H80000012
Circle (linhaY.x1 + (estbeq * escalax), Ybeq), 0.1 * rbeq, &H80000012

'Desenho da linha de terra

yterra = linhaif + (yprinc + rprinc)

Line (linhaY.x1, yterra)-(linhaY.x1 + (Form2.lfs * escalax), yterra), &H80000012

End Sub

Private Sub Command14_Click()
Call drawelipse_Click
Form18.MousePointer = 2
End Sub






Private Sub Command19_Click()
If alfa < 0 Or alfa = Empty Then
MsgBox " O raio e o angulo devem ser maiores que ""zero""", vbCritical
Exit Sub
End If

Dim escalax, fatx, kalfa As Double
Call ColorPen_Click
'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

kalfa = 6.28
Form18.Circle (1200 + (xx * escalax), yy * escalax), Sqr((((raio / (2 ^ 0.5))) * escalax) ^ 2 + (((raio / (2 ^ 0.5))) * escalax) ^ 2), , (kalfa - (3.1415 / 180) * (alfa))

End Sub

Private Sub Command2_Click()
Call DrawLine_Click
Call StyleDash_Click
Form18.MousePointer = 2
End Sub

Private Sub Command20_Click()
Call ColorPen_Click
Form18.MousePointer = 2
    Shape = "LINE2"
End Sub

Private Sub Command21_Click()
Call DrawText_Click
Form18.MousePointer = 2
End Sub

Private Sub Command22_Click()
Call ColorPen_Click
Form18.MousePointer = 2
    Shape = "LINE3"
End Sub

Private Sub Command23_Click()
Call filedistancia_Click
End Sub

Private Sub Command24_Click()
Call filedistanciaY_Click
End Sub

Private Sub Command25_Click()
Frame1.Visible = False
Call inseridesenho_Click
Label3.Top = 120
med2.Top = 360
End Sub

Private Sub Command26_Click()
Frame12.Visible = False
Frame11.Visible = False
Label3.Left = 120
med2.Left = 120
End Sub

Private Sub Command27_Click()
Call filelinhaux_Click
Frame22.Visible = False
Form18.Frame6.Visible = True
Form18.Frame1.Visible = True
Form18.Frame12.Visible = True
Form18.Option1.Value = True
End Sub


Private Sub Command28_Click()
Dim kfilasa, kfilaht As Double

kfilasa = 1 - Form1.Afilasa
kfilaht = 1 - Form1.AfilaHt

If kfilasa = 0 Then Form18.HScroll1.Enabled = False
If kfilaht = 0 Then Form18.HScroll2.Enabled = False

If kfilasa <> 0 Then Form18.HScroll1.Enabled = True
If kfilaht <> 0 Then Form18.HScroll2.Enabled = True

If Form1.Afilasa <> 0 Then Call fileplanta_Click
If Form1.AfilaHt <> 0 Then Call fileplanta_Click

Form18.afilaASA = Form1.Afilasa
Form18.AfilaHt = Form1.AfilaHt
End Sub

Private Sub Command29_Click()
Frame6.Visible = False
Command30.Enabled = True
vsescala.Enabled = False
vslinha.Enabled = False
'Inserindo ferramentas de equipamentos
Frame1.Visible = True
Label3.Top = 1440
med2.Top = 1680
End Sub

Private Sub Command3_Click()
Call DrawLine_Click
Call StyleDot_Click
Form18.MousePointer = 2
End Sub

Private Sub Command30_Click()
Dim resp1 As String

resp1 = MsgBox("Os desenhos serão apagados, deseja continuar?", vbOKCancel)

If resp1 = vbOK Then
    vsescala.Enabled = True
    vslinha.Enabled = True
    Command30.Enabled = False
 End If
 
 If resp1 = vbCancel Then Exit Sub
End Sub

Private Sub Command31_Click()

Call ColorPen_Click

Dim fatx, escalax As Double

fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

Line ((1200 + (xx * escalax)), (yy * escalax))-((1200 + ((Abs(xx) + Abs(base)) * escalax)), yy * escalax)

Line ((1200 + ((Abs(xx) + Abs(base)) * escalax)), yy * escalax)-((1200 + ((Abs(xx) + Abs(base)) * escalax)), (Abs(yy) - Abs(altura)) * escalax)

Line ((1200 + ((Abs(xx) + Abs(base)) * escalax)), (Abs(yy) - Abs(altura)) * escalax)-((1200 + (xx * escalax)), (Abs(yy) - Abs(altura)) * escalax)

Line ((1200 + (xx * escalax)), (Abs(yy) - Abs(altura)) * escalax)-((1200 + (xx * escalax)), (yy * escalax))

End Sub

Private Sub Command32_Click()
Call filelateral_Click
End Sub

Private Sub Command33_Click()
Call fileplanta_Click
End Sub

Private Sub Command34_Click()


Call filelateral_Click

If Form20.p1 <> Form21.Data1.Recordset("Projeto") Then
    Form18.HScroll4.Enabled = True
    Form18.HScroll4.Value = 0
    Form18.HScroll3.Enabled = True
End If

    If Form20.p1 = Form21.Data1.Recordset("Projeto") Then

 Form18.HScroll3.Enabled = True '(21/05/2002)
 
        Form18.afilaASA = Form21.Data1.Recordset("AfilamW")
        Form18.AfilaHt = Form21.Data1.Recordset("Afilamht")
     
        Form18.hsasaval = Form21.Data1.Recordset("f18hs1")
        Form18.HScroll1.Value = Form18.hsasaval
    
        Form18.hshtval = Form21.Data1.Recordset("f18hs2")
        Form18.HScroll2.Value = Form18.hshtval
    
        Form18.gradvt = Form21.Data1.Recordset("f18hs3")
        Form18.HScroll3.Value = Form18.gradvt
    
        Form18.vt = Form21.Data1.Recordset("f18hs4")
        
        Form18.hs4 = Form21.Data1.Recordset("Fo18hs4")
        
        Form18.HScroll4.Enabled = True
        Form18.HScroll4.Value = Form18.hs4
        
    End If

End Sub
Private Sub Command4_Click()
Form18.MousePointer = 2
Call DrawCircle_Click
End Sub
Private Sub Command5_Click()
'Form18.Cls 'limpando

Call insereasa_Click

End Sub

Private Sub Command6_Click()
 stylefilled.Checked = Not stylefilled.Checked
    If stylefilled.Checked Then
        Form18.FillStyle = 0
    Else
        Form18.FillStyle = 1
    End If
    
Dim escalax, fatx As Double
'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx
 Form18.DrawWidth = 2

'Motor
Circle ((1680 + (Form14.d1 / 100) * escalax), linhaX.y1), 0.05 * escalax, &HFF&
'Piloto e Copiloto
Circle ((1680 + (Form14.d5 / 100) * escalax), linhaX.y1), 0.05 * escalax, &HFF00FF
'Ocupantes
Circle ((1680 + (Form14.d6 / 100) * escalax), linhaX.y1), 0.05 * escalax, &H80FF&
'Combustivel
Circle ((1680 + (Form14.d2 / 100) * escalax), linhaX.y1), 0.05 * escalax, &HC000&
'Bagagem
Circle ((1680 + (Form14.d7 / 100) * escalax), linhaX.y1), 0.05 * escalax, &HFF0000
'Instrumentos
Circle ((1680 + (Form14.d3 / 100) * escalax), linhaX.y1), 0.05 * escalax, &HFFFF&
'Outros
Circle ((1680 + (Form14.d15 / 100) * escalax), linhaX.y1), 0.05 * escalax, &HFFFF00
'Fuselagem
Circle ((1680 + (Form14.d12 / 100) * escalax), linhaX.y1), 0.05 * escalax, &H0&

    Form18.FillStyle = 1
    Form18.DrawWidth = 1

End Sub

Private Sub Command8_Click()
Call DrawBox_Click
Call StyleSolid_Click
Form18.MousePointer = 2
End Sub


Private Sub dondecol_Click()
Dim fatx, escalax, cordaw, distw, distcam, dist, ed, k, l9, ya8, cordaht, ktail, dwx1, dwy, dwy1, dwx2, dwy2n, dwy2p, deflex, EpsonD, cd, Fa, Fb1, Fb2, Fb, Ew, cdmax, Eps As Double

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

'Efeito Diedro - ed tangente de 1º = 0.017453
ed = ((Form10.diedro * 0.017453) * Form2.bw / 2) * escalax

If Form1.posialta.Value = True Then k = linhasf.y1 - (ed / 2)
If Form1.posimedia.Value = True Then k = Form18.valasa - (ed / 2)
If Form1.posibaixa.Value = True Then k = linhaif.y1 - (ed / 2)

cordaw = Form7.CMaerodw * escalax
distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
ya8 = 0.078 * cordaw
l9 = 0.3 * cordaw

'Dados de geometria das superficies transformadas
'Corda da asa
cordaht = Form7.CMaerodht * escalax
distcam = distw + (0.25 * cordaw)
dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)

' Deflexão da linha devido a angulo downwash
EpsonD = 35 * Form7.cldecol / Form1.arw
EpsonD = Format(EpsonD, "0.000")
deflex = (Tan(EpsonD * (3.141516 / 180))) * (distcam + (Form10.LT * escalax))

Form18.DrawStyle = 2 ' Tracejado
'Linha média
Line (distcam, k - (0.25 * ya8))-(distcam + (Form10.LT * escalax), k - (0.25 * ya8) + deflex), &H800000

DrawStyle = 2 ' Pontilhado

dwx1 = distcam
dwy = k - (0.25 * ya8)
dwy1 = k - (0.25 * ya8) + deflex

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax

'Espessura do Downwash
cdmax = Form4.Data1.Recordset("cdmax")

'METODO STINTON
'cd = Abs((Form7.cldecol ^ 2) / (3.1415 * Form1.arw * 0.9) + Abs(cdmax))
'Fa = ((cd ^ 0.5) / (5 * 3.141516))
'Fb1 = Form10.LT / 0.305
'Fb2 = Form2.cmw / 0.305
'Fb = (Form1.arw * (Fb1 / Fb2)) ' A fórmula correta seria: Fb = (Form1.arw * (Fb1 / Fb2)) +7
'Ew = ((Fa * Fb) * (Form2.cmw / 0.305)) * 0.305
'Ew = Format(Ew, "0.000")

'METODO DOMMASCH
Eps = 20 * Form7.cldecol * (((1 / Form1.Afilasa) ^ 0.3) / (Form1.arw ^ 0.725)) * ((3 * Form7.CMaerodw / Form10.LT) ^ 0.25)
Ew = (Tan(Eps * 3.1415 / 180)) * Form10.LT
Ew = Format(Ew, "0.000")

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax


' Para decolagem
    Line (dwx1, dwy)-(dwx2, dwy2n), &HC0C000
    Line (dwx1, dwy)-(dwx2, dwy2p), &HC0C000

    Form18.MousePointer = 2
    Shape = "LINE"
    Form18.DrawStyle = 0
End Sub

Private Sub donestol_Click()
Dim Eps, fatx, escalax, cordaw, distw, distcam, dist, ed, k, l9, ya8, cordaht, ktail, dwx1, dwy, dwy1, dwx2, dwy2n, dwy2p, deflex, EpsonS, cd, Fa, Fb1, Fb2, Fb, Ew, cdmax As Double

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

'Efeito Diedro - ed tangente de 1º = 0.017453
ed = ((Form10.diedro * 0.017453) * Form2.bw / 2) * escalax

If Form1.posialta.Value = True Then k = linhasf.y1 - (ed / 2)
If Form1.posimedia.Value = True Then k = Form18.valasa - (ed / 2)
If Form1.posibaixa.Value = True Then k = linhaif.y1 - (ed / 2)

cordaw = Form7.CMaerodw * escalax
distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
ya8 = 0.078 * cordaw
l9 = 0.3 * cordaw

'Dados de geometria das superficies transformadas
'Corda da asa
cordaht = Form7.CMaerodht * escalax
distcam = distw + (0.25 * cordaw)
dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)

' Deflexão da linha devido a angulo downwash
EpsonS = 35 * Form7.clk1 / Form1.arw
EpsonS = Format(EpsonS, "0.000")

deflex = (Tan(EpsonS * (3.141516 / 180))) * (distcam + (Form10.LT * escalax))

Form18.DrawStyle = 2
'Linha média
Line (distcam, k - (0.25 * ya8))-(distcam + (Form10.LT * escalax), k - (0.25 * ya8) + deflex), &H800000

DrawStyle = 2 ' Pontilhado

dwx1 = distcam
dwy = k - (0.25 * ya8)
dwy1 = k - (0.25 * ya8) + deflex

'Espessura do Downwash
cdmax = Form4.Data1.Recordset("cdmax")

' METODO STINTON
'cd = Abs((Form7.clk1 ^ 2) / (3.1415 * Form1.arw * 0.9) + Abs(cdmax))
'Fa = ((cd ^ 0.5) / (5 * 3.141516))
'Fb1 = Form10.LT / 0.305
'Fb2 = Form2.cmw / 0.305
'Fb = (Form1.arw * (Fb1 / Fb2)) ' A fórmula correta seria: Fb = (Form1.arw * (Fb1 / Fb2)) +7
'Ew = ((Fa * Fb) * (Form2.cmw / 0.305)) * 0.305
'Ew = Format(Ew, "0.000")

'METODO DOMMASCH
Eps = 20 * Form7.clk1 * (((1 / Form1.Afilasa) ^ 0.3) / (Form1.arw ^ 0.725)) * ((3 * Form7.CMaerodw / Form10.LT) ^ 0.25)
Ew = (Tan(Eps * 3.1415 / 180)) * Form10.LT
Ew = Format(Ew, "0.000")

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax

' Para stall
    Line (dwx1, dwy)-(dwx2, dwy2n), &HC0&
    Line (dwx1, dwy)-(dwx2, dwy2p), &HC0&


    Form18.MousePointer = 2
    Shape = "LINE"
    Form18.DrawStyle = 0
End Sub

Private Sub donsub_Click()
Dim fatx, escalax, cordaw, distw, distcam, dist, ed, k, l9, ya8, cordaht, ktail, dwx1, dwy, dwy1, dwx2, dwy2n, dwy2p, deflex, EpsonSB, cd, Fa, Fb1, Fb2, Fb, Ew, cdmax, Eps As Double

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

'Efeito Diedro - ed tangente de 1º = 0.017453
ed = ((Form10.diedro * 0.017453) * Form2.bw / 2) * escalax

If Form1.posialta.Value = True Then k = linhasf.y1 - (ed / 2)
If Form1.posimedia.Value = True Then k = Form18.valasa - (ed / 2)
If Form1.posibaixa.Value = True Then k = linhaif.y1 - (ed / 2)


cordaw = Form7.CMaerodw * escalax
distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
ya8 = 0.078 * cordaw
l9 = 0.3 * cordaw

'Dados de geometria das superficies transformadas
'Corda da asa
cordaht = Form7.CMaerodht * escalax
distcam = distw + (0.25 * cordaw)
dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)

' Deflexão da linha devido a angulo downwash
EpsonSB = 35 * Form7.clk2 / Form1.arw
EpsonSB = Format(EpsonSB, "0.000")
deflex = (Tan(EpsonSB * (3.141516 / 180))) * (distcam + (Form10.LT * escalax))

Form18.DrawStyle = 2 ' Tracejado
'Linha média
Line (distcam, k - (0.25 * ya8))-(distcam + (Form10.LT * escalax), k - (0.25 * ya8) + deflex), &H800000

DrawStyle = 2 ' Pontilhado

dwx1 = distcam
dwy = k - (0.25 * ya8)
dwy1 = k - (0.25 * ya8) + deflex

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax

'Espessura do Downwash
cdmax = Form4.Data1.Recordset("cdmax")

'METODO STINTON
'cd = Abs((Form7.clk2 ^ 2) / (3.1415 * Form1.arw * 0.9) + Abs(cdmax))
'Fa = ((cd ^ 0.5) / (5 * 3.141516))
'Fb1 = Form10.LT / 0.305
'Fb2 = Form2.cmw / 0.305
'Fb = (Form1.arw * (Fb1 / Fb2)) ' A fórmula correta seria: Fb = (Form1.arw * (Fb1 / Fb2)) +7
'Ew = ((Fa * Fb) * (Form2.cmw / 0.305)) * 0.305
'Ew = Format(Ew, "0.000")

'METODO DOMMASCH
Eps = 20 * Form7.clk2 * (((1 / Form1.Afilasa) ^ 0.3) / (Form1.arw ^ 0.725)) * ((3 * Form7.CMaerodw / Form10.LT) ^ 0.25)
Ew = (Tan(Eps * 3.1415 / 180)) * Form10.LT
Ew = Format(Ew, "0.000")

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax


' Para subida
    Line (dwx1, dwy)-(dwx2, dwy2n), &HC000&
    Line (dwx1, dwy)-(dwx2, dwy2p), &HC000&

    Form18.MousePointer = 2
    Shape = "LINE"
    Form18.DrawStyle = 0
End Sub

Private Sub donvmax_Click()
Dim fatx, escalax, cordaw, distw, distcam, dist, ed, k, l9, ya8, cordaht, ktail, dwx1, dwy, dwy1, dwx2, dwy2n, dwy2p, deflex, EpsonVM, cd, Fa, Fb1, Fb2, Fb, Ew, cdmax, Eps As Double

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

'Efeito Diedro - ed tangente de 1º = 0.017453
ed = ((Form10.diedro * 0.017453) * Form2.bw / 2) * escalax

If Form1.posialta.Value = True Then k = linhasf.y1 - (ed / 2)
If Form1.posimedia.Value = True Then k = Form18.valasa - (ed / 2)
If Form1.posibaixa.Value = True Then k = linhaif.y1 - (ed / 2)


cordaw = Form7.CMaerodw * escalax
distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
ya8 = 0.078 * cordaw
l9 = 0.3 * cordaw

'Dados de geometria das superficies transformadas
'Corda da asa
cordaht = Form7.CMaerodht * escalax
distcam = distw + (0.25 * cordaw)
dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)

' Deflexão da linha devido a angulo downwash
EpsonVM = 35 * Form7.clk4 / Form1.arw
EpsonVM = Format(EpsonVM, "0.000")
deflex = (Tan(EpsonVM * (3.141516 / 180))) * (distcam + (Form10.LT * escalax))

Form18.DrawStyle = 2 ' Tracejado
'Linha média
Line (distcam, k - (0.25 * ya8))-(distcam + (Form10.LT * escalax), k - (0.25 * ya8) + deflex), &H800000

DrawStyle = 2 ' Pontilhado

dwx1 = distcam
dwy = k - (0.25 * ya8)
dwy1 = k - (0.25 * ya8) + deflex

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax

'Espessura do Downwash
cdmax = Form4.Data1.Recordset("cdmax")

'METODO STINTON
'cd = Abs((Form7.clk4 ^ 2) / (3.1415 * Form1.arw * 0.9) + Abs(cdmax))
'Fa = ((cd ^ 0.5) / (5 * 3.141516))
'Fb1 = Form10.LT / 0.305
'Fb2 = Form2.cmw / 0.305
'Fb = (Form1.arw * (Fb1 / Fb2)) ' A fórmula correta seria: Fb = (Form1.arw * (Fb1 / Fb2)) +7
'Ew = ((Fa * Fb) * (Form2.cmw / 0.305)) * 0.305
'Ew = Format(Ew, "0.000")

'METODO DOMMASCH
Eps = 20 * Form7.clk4 * (((1 / Form1.Afilasa) ^ 0.3) / (Form1.arw ^ 0.725)) * ((3 * Form7.CMaerodw / Form10.LT) ^ 0.25)
Ew = (Tan(Eps * 3.1415 / 180)) * Form10.LT
Ew = Format(Ew, "0.000")

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax


' Para cruzeiro
    Line (dwx1, dwy)-(dwx2, dwy2n), &HC000C0
    Line (dwx1, dwy)-(dwx2, dwy2p), &HC000C0


    Form18.MousePointer = 2
    Shape = "LINE"
    Form18.DrawStyle = 0
End Sub

Private Sub DrawBox_Click()
Form18.MousePointer = 2
Call ColorPen_Click
    Shape = "BOX"
End Sub

Private Sub DrawCircle_Click()
Form18.MousePointer = 2
Call ColorPen_Click
    Shape = "CIRCLE"
End Sub

Private Sub drawelipse_Click()
Form18.MousePointer = 2
Call ColorPen_Click
  Shape = "ELIPSE"

End Sub

Private Sub DrawLine_Click()
Call ColorPen_Click
Form18.MousePointer = 2
    Shape = "LINE"
End Sub


Private Sub DrawText_Click()

    Dim DrawString As String
    DrawString = InputBox("Entre com o texto")
    Label1.Caption = DrawString
    Form18.ForeColor = CommonDialog1.Color
    PrintText = True
        
End Sub

Private Sub EditClear_Click()

    Form18.Cls
    Form18.MousePointer = 1
    
End Sub

Private Sub EditCopy_Click()
    Form18.MousePointer = 2
    CopyBMP = True
      Form18.ForeColor = &HC0C0C0
      
End Sub

Private Sub EditCut_Click()
    Form18.MousePointer = 2
    CutBMP = True
    Form18.ForeColor = &H80000012
End Sub

Private Sub EditPaste_Click()
    Form18.MousePointer = 2
    PasteBMP = True
    
End Sub


Private Sub fileaerodin18_Click()
Form7.Visible = True
Form18.Visible = False
Form7.Refresh
End Sub

Private Sub fileapagalin_Click()
Dim fatx, escalax, tango, tango1, cordaht, distcam, dist, L1, ktail, distw, cordaw As Double
'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx
'..................AEROFOLIO DA EMPENAGEM
cordaht = Form7.CMaerodht * escalax
cordaw = Form7.CMaerodw * escalax

distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
distcam = distw + (0.25 * cordaw)
dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

cordaht = Form7.CMaerodht * escalax

dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)

tango = 1.732 * ((1 + dist) - (1 + dist + (escalax))) ' Para 60º
tango1 = 0.57735 * ((1 + dist) - (1 + dist + (escalax))) ' Para 30º
'TRAÇADO DA LINHA para apagar
Form18.Line (L1 + dist, ktail)-(1 + dist + (escalax), ktail + tango), &HFFFFFF

'Bordo de fuga
Form18.Line (cordaht + dist, ktail)-(cordaht + dist + (escalax), ktail + tango1), &HFFFFFF
End Sub

Private Sub fileasa18_Click()
Dim respdes As Double
respdes = MsgBox("Salvar vista lateral e planta?", vbYesNo)

If respdes = vbYes Then
Call FileSaveAs_Click
Exit Sub
End If

If respdes = vbNo Then GoTo 10

10:
Form18.Cls
Form18.MousePointer = 1

fileplanta.Enabled = False
'filefrontal.Enabled = False

Dim fatx, escalax As Double

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

Form18.linhasf.Visible = True
Form18.linhafinal.Visible = True
linhaasa1.Visible = True
Form18.linhaY.Visible = True
Form18.linhaif.Visible = True
Form18.linhaleme.Visible = False

'Escala
'escalasa = (0.9 * Form18.Width) / Form2.bw 'fator de escala

Form18.linhaY.x1 = Form18.Width / 2
Form18.linhaY.X2 = Form18.Width / 2

linhaX.y1 = 2000
linhaX.y2 = 2000

Form18.linhasf.y1 = linhaX.y1
Form18.linhasf.y2 = linhaX.y1

Form18.linhaif.y1 = linhaX.y1 + Form2.crw * escalax
Form18.linhaif.y2 = linhaX.y1 + Form2.crw * escalax

'Linha final para asa
linhafinal.y2 = linhaif.y2
linhafinal.y1 = linhasf.y1

Form18.linhafinal.x1 = linhaY.x1 + ((Form2.bw / 2) * escalax)
Form18.linhafinal.X2 = linhaY.X2 + ((Form2.bw / 2) * escalax)

linhaasa1.y2 = linhaif.y2
linhaasa1.y1 = linhasf.y1

Form18.linhaasa1.x1 = linhaY.x1 - ((Form2.bw / 2) * escalax)
Form18.linhaasa1.X2 = linhaY.X2 - ((Form2.bw / 2) * escalax)

End Sub

Private Sub filecalcon_Click()
Form29.Show
End Sub

Private Sub filecofigura_Click()
Frame22.Visible = True
End Sub

Private Sub filecoord_Click()
MsgBox "Dê um click com o mouse sobre o ponto que deseja conhecer as coordenadas"
Form18.ForeColor = &HFFFFFF
End Sub

Private Sub filedimensiona18_Click()
Form18.Visible = False
Form2.Visible = True
Form2.Refresh
End Sub


Private Sub filedistancia_Click()
Call ColorPen_Click
Form18.ForeColor = &H0&
Form18.MousePointer = 2
Shape = "dimensionX"
End Sub

Private Sub filedistanciaY_Click()
Call ColorPen_Click
Form18.ForeColor = &H0&
Form18.MousePointer = 2
Shape = "dimensionY"
End Sub

Private Sub fileEH18_Click()
Dim respdes2 As Double
respdes2 = MsgBox("Salvar vista lateral e planta?", vbYesNo)

If respdes2 = vbYes Then
Call FileSaveAs_Click
Exit Sub
End If

If respdes2 = vbNo Then GoTo 10

10:
Form18.Cls
Form18.MousePointer = 1

fileplanta.Enabled = False
'filefrontal.Enabled = False

Dim fatx, escalax As Double

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

Form18.linhasf.Visible = True
Form18.linhafinal.Visible = True
linhaasa1.Visible = True
Form18.linhaY.Visible = True
Form18.linhaif.Visible = True
Form18.linhaleme.Visible = False

'Escala
'escalasa = (0.9 * Form18.Width) / Form2.bw 'fator de escala

Form18.linhaY.x1 = Form18.Width / 2
Form18.linhaY.X2 = Form18.Width / 2

linhaX.y1 = 5040
linhaX.y2 = 5040

Form18.linhasf.y1 = linhaX.y1
Form18.linhasf.y2 = linhaX.y1

Form18.linhaif.y1 = linhaX.y1 + Form2.crht * escalax
Form18.linhaif.y2 = linhaX.y1 + Form2.crht * escalax

linhafinal.y2 = linhaif.y2
linhafinal.y1 = linhasf.y1

Form18.linhafinal.x1 = linhaY.x1 + ((Form2.bht / 2) * escalax)
Form18.linhafinal.X2 = linhaY.X2 + ((Form2.bht / 2) * escalax)

linhaasa1.y2 = linhaif.y2
linhaasa1.y1 = linhasf.y1

Form18.linhaasa1.x1 = linhaY.x1 - ((Form2.bht / 2) * escalax)
Form18.linhaasa1.X2 = linhaY.X2 - ((Form2.bht / 2) * escalax)

End Sub

Private Sub fileestab18_Click()
Form10.Visible = True
Form18.Visible = False
Form10.Refresh
End Sub

Private Sub fileestrut18_Click()
Form11.Visible = True
Form18.Visible = False
Form11.Refresh
End Sub



Private Sub fileimpdes_Click()
Form31.Show
End Sub


Private Sub fileimprimi18_Click()

If LIBER <> "KFAR0852" Then
    'MsgBox "A rotina de impressão sómente estará disponível após a liberação permanente", 16, "Cancelamento de impressão": Exit Sub
End If

Dim Tipo, asa, trem, sistema, regime, We, Wf, Wcu, Wtot As String

We = Requisitos.weR / 2.2
We = Format(We, "0.00")
Wf = Requisitos.wfR / 2.2
Wf = Format(Wf, "0.00")
Wcu = Requisitos.CPR / 2.2
Wcu = Format(Wcu, "0.00")
Wtot = Requisitos.woestimadoR / 2.2
Wtot = Format(Wtot, "0.00")

'REQUERIMENTO

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
Printer.Print Tab(20); "Alongamento da empenagem horizontal:"; Tab(85); Form1.ARht
Printer.Print Tab(20); "Alongamento da empenagem vertical:"; Tab(85); Form1.ARvt
Printer.Print Tab(20); "Afilamento da asa :"; Tab(85); Form1.Afilasa
Printer.Print Tab(20); "Afilamento da empenagem horizontal:"; Tab(85); Form1.AfilaHt
Printer.Print Tab(20); "Afilamento da empenagem vertical:"; Tab(85); Form1.Afilavt
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
Printer.Print Tab(20); "Fim da página"
Printer.EndDoc

'DIMENSIONAMENTO
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
'AERODINÂMICA
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "II - AERODINÂMICA"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "1 - ASA"
Printer.FontSize = 10
Printer.Print Tab(20); "Angulo alfa para Clmáx.s/flap:"; Tab(70); Form7.alfasamax;
Printer.Print Tab(20); "Angulo alfa para Clmáx.c/flap:"; Tab(70); Form7.alfaFLAPm;
Printer.Print Tab(20); "Angulo alfa para Clo.s/flap:"; Tab(70); Form7.alfasazero;
Printer.Print Tab(20); "Angulo alfa para Clo.c/flap:"; Tab(70); Form7.alfaFLAPo;
Printer.Print Tab(20); "Coeficiente de arrasto do perfil:"; Tab(70); Form7.cdperfasa;
Printer.Print Tab(20); "Coeficiente Cl3D.s/flap:"; Tab(70); Form7.cl3dasa;
Printer.Print Tab(20); "Coeficiente Cl3D.c/flap:"; Tab(70); Form7.cl3dFLAP;
Printer.Print Tab(20); "dCl/dalfa - s/flap:"; Tab(70); Form7.slopeasa; " /rad"
Printer.Print Tab(20); "dCl/dalfa - c/flap:"; Tab(70); Form7.slopeFLAP; " /rad"
Printer.Print Tab(20); "Área molhada da asa:"; Tab(70); Form7.swetasa; " m2"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "2 - ESTABILIZADOR HORIZONTAL"
Printer.FontSize = 10
Printer.Print Tab(20); "Angulo alfa para Clmáximo:"; Tab(70); Form7.alfahtmax;
Printer.Print Tab(20); "Angulo alfa para Clzero:"; Tab(70); Form7.alfahtzero;
Printer.Print Tab(20); "Coeficiente de arrasto do perfil:"; Tab(70); Form7.cdperfht;
Printer.Print Tab(20); "Coeficiente de sustentação 3D:"; Tab(70); Form7.cl3dht;
Printer.Print Tab(20); "Inclinação da curva de sustentação:"; Tab(70); Form7.slopeht; " /rad"
Printer.Print Tab(20); "Área molhada do Estabilizador horizontal:"; Tab(70); Form7.swethorz; " m2"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "3 - ESTABILIZADOR VERTICAL"
Printer.FontSize = 10
Printer.Print Tab(20); "Coeficiente de arrasto do perfil:"; Tab(70); Form7.cdperfvt;
Printer.Print Tab(20); "Área molhada do Estabilizador Vertical:"; Tab(70); Form7.swetvert; " m2"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "4 - FUSELAGEM"
Printer.FontSize = 10
Printer.Print Tab(20); "Área molhada da Fuselagem:"; Tab(70); Form7.swetfuse; " m2"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "5 - SUSTENTAÇÃO"
Printer.FontSize = 10
Printer.Print Tab(20); "Carga alar de projeto:"; Tab(70); Form7.wsproj; " kgf/m2"
Printer.Print Tab(20); "Cl em estol:"; Tab(70); Form7.clk1;
Printer.Print Tab(20); "Cl em decolagem:"; Tab(70); Form7.cldecol;
Printer.Print Tab(20); "Cl em subida:"; Tab(70); Form7.clk2;
Printer.Print Tab(20); "Cl em velocidade máxima:"; Tab(70); Form7.clk4;
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print ""
Printer.Print Tab(25); "6 - ARRASTO"
Printer.FontSize = 10
Printer.Print Tab(20); "Fator de placa plana:"; Tab(70); Form7.placaplan; " m2"
Printer.Print Tab(20); "Cdtotal em estol:"; Tab(70); Form7.cdtotstall;
Printer.Print Tab(20); "Cdtotal em decolagem:"; Tab(70); Form7.cddecol;
Printer.Print Tab(20); "Cdtotal em subida:"; Tab(70); Form7.cdtotsubida;
Printer.Print Tab(20); "Cdtotal em velocidade máxima:"; Tab(70); Form7.cdtotvmax;
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print ""
Printer.Print Tab(25); "7 - EFICIENCIA AERODINÂMICA"
Printer.FontSize = 10
Printer.Print Tab(20); "Melhor angulo de planeio:"; Tab(70); Form7.anplan; " º"
Printer.Print Tab(20); "L/D em estol:"; Tab(70); Form7.ldk1;
Printer.Print Tab(20); "L/D em decolagem:"; Tab(70); Form7.lddecol;
Printer.Print Tab(20); "L/D em subida:"; Tab(70); Form7.ldk2;
Printer.Print Tab(20); "L/D em velocidade máxima:"; Tab(70); Form7.ldk4;
Printer.Print ""
Printer.Print Tab(20); "Fim da página";
Printer.EndDoc


'PERFORMANCE
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
Printer.Print Tab(20); "Especificação do motor:"; Tab(70); Form8.mtr;
Printer.Print Tab(20); "Peso do motor:"; Tab(70); Form8.pesomot; " kgf"
Printer.Print Tab(20); "HP nominal do motor - original:"; Tab(70); Form8.hpnominal; " HP"
Printer.Print Tab(20); "RPM nominal do motor - original:"; Tab(70); Form8.rpmf1; " RPM"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "2 - HÉLICE"
Printer.FontSize = 10
Printer.Print Tab(20); "Especificada por:"; Tab(70); Form8.cahel;
Printer.Print Tab(20); "Diametro da hélice:"; Tab(70); Form8.diahel; " pol."
Printer.Print Tab(20); "Passo da hélice:"; Tab(70); Form8.pashel; " pol."
Printer.Print Tab(20); "Avanço da hélice:"; Tab(70); Form8.avahel; " pol."
Printer.Print Tab(20); "Eficiencia da hélice:"; Tab(70); Form8.NI; " %"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "3 - VELOCIDADE"
Printer.FontSize = 10
Printer.Print Tab(20); "Velocidade de stall sem flap:"; Tab(70); Form8.vst2; " Km/h"; "     com flap:"; Form8.vstol; " Km/h"
Printer.Print Tab(20); "Velocidade de decolagem sem flap:"; Tab(70); Form8.vdecolsf; " Km/h"; "     com flap:"; Form8.vdecol; " Km/h"
Printer.Print Tab(20); "Velocidade de subida (ideal):"; Tab(70); Form8.vsubida; " Km/h"
Printer.Print Tab(20); "Velocidade máxima:"; Tab(70); Form8.vmax; " Km/h"
Printer.Print Tab(20); "Velocidade de toque no pouso sem flap:"; Tab(70); Form8.toque2; " Km/h"; "      com flap:"; Form8.toque; " Km/h"
Printer.Print Tab(20); "Velocidade de planeio:"; Tab(70); Form8.velplan; " Km/h"
Printer.Print Tab(20); "Razão de subida:"; Tab(70); Form8.climb; " ft/min"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "4 - POTENCIA"
Printer.FontSize = 10
Printer.Print Tab(20); "Potencia disponivel real:"; Tab(70); Form8.hpreal; " HP"
Printer.Print Tab(20); "Potencia requerida em Vstall:"; Tab(70); Form8.potreqstol; " HP"
Printer.Print Tab(20); "Potencia requerida em Vdecolagem:"; Tab(70); Form8.potreqdec; " HP"
Printer.Print Tab(20); "Potencia requerida em Vsubida:"; Tab(70); Form8.potreqsub; " HP"
Printer.Print Tab(20); "Potencia requerida em Vmáxima:"; Tab(70); Form8.potreq; " HP"
Printer.Print Tab(20); "Sobra de potencia - Vstall/Vmáxima:"; Tab(70); Form8.hpreal; " HP"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "5 - FORÇA E TRAÇÃO"
Printer.FontSize = 10
Printer.Print Tab(20); "Força de arrasto em Vstall:"; Tab(70); Form8.FDstol; " kgf"
Printer.Print Tab(20); "Força de arrasto em Vdecolagem:"; Tab(70); Form8.FDdecol; " kgf"
Printer.Print Tab(20); "Força de arrasto em Vsubida:"; Tab(70); Form8.Fdsubida; " kgf"
Printer.Print Tab(20); "Força de arrasto em Vmáxima:"; Tab(70); Form8.FD; " kgf"
Printer.Print Tab(20); "Tração estática:"; Tab(70); Form8.tracestatica; " kgf"
Printer.Print Tab(20); "Tração dinâmica:"; Tab(70); Form8.tradin; " kgf"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "6 - GERAIS"
Printer.FontSize = 10
Printer.Print Tab(20); "Angulo de subida(ideal):"; Tab(70); Form8.angclimb; " º"
Printer.Print Tab(20); "Distancia de decolagem s/flap:"; Tab(70); Form8.ldecsf; " m"; "    com flap:"; Form8.Ldecol; " m"
Printer.Print Tab(20); "Distancia de pouso s/flap:"; Tab(70); Form8.lpousf; " m"; "    com flap:"; Form8.Lpouso; " m"
Printer.Print Tab(20); "Relação L/D e W/T:"; Tab(70); Form8.ldtw;
Printer.Print Tab(20); "Alcance:"; Tab(70); Requisitos.alcanceBD; " nm"
Printer.Print Tab(20); "Razão de subida:"; Tab(70); Form8.climb; " ft/min"
Printer.Print Tab(20); "CAFE challenge:"; Tab(70); Form8.cafechal;
Printer.Print Tab(20); "CAFE Triviathon:"; Tab(70); Form8.cafetria;
Printer.Print ""
Printer.Print Tab(20); "Fim da página";
Printer.EndDoc


'ESTABILIDADE
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
Printer.Print Tab(20); "Localização da corda média da asa:"; Tab(70); Form10.yasa; " m"
Printer.Print Tab(20); "Localização da corda média do Etb.Horz.:"; Tab(70); Form10.yht; " m"
Printer.Print Tab(20); "Volume da empenagem horizontal:"; Tab(70); Form10.volht;
Printer.Print Tab(20); "Volume da empenagem vertical:"; Tab(70); Form10.volvt;
Printer.Print Tab(20); "Fator de eficiencia do Etb.Horz.:"; Tab(70); Form10.eficiht;
Printer.Print Tab(20); "Variação do angulo de downwash da asa:"; Tab(70); Form10.DEDA;
Printer.Print Tab(20); "Força de sustentação da asa:"; Tab(70); Form10.sustentasa; " kgf"
Printer.Print Tab(20); "Força de sustentação do Etb.Horz.:"; Tab(70); Form10.sustentaht; " kgf"
Printer.Print Tab(20); "Angulo de incidencia da asa:"; Tab(70); Form10.incidasa; " º"
Printer.Print Tab(20); "Angulo de incidencia do Etb.Horz.:"; Tab(70); Form10.incidaht; " º"
Printer.Print Tab(20); "Localização do Etb.Horz. referente a L.A:"; Tab(70); Form10.htail; " m"
Printer.Print Tab(20); "Localização do C.A-asa na corda média:"; Tab(70); Form10.CAasa; " m"
Printer.Print Tab(20); "Localizaçãodo C.A-ht na corda média:"; Tab(70); Form10.CAht; " m"
Printer.Print Tab(20); "Distancia C.A da asa ao C.A-ht (LT):"; Tab(70); Form10.LT; " m"
Printer.Print Tab(20); "Localização do Ponto neutro P.N:"; Tab(70); Form10.PN; " %"
Printer.Print Tab(20); "Localização do Centro de gravidade C.G:"; Tab(70); Form10.cg; " %"
Printer.Print Tab(20); "Margem estática em porcentagem:"; Tab(70); Form10.MARGEM; " %"
Printer.Print Tab(20); "Distancia entre P.N e C.G:"; Tab(70); Form10.PNCG; " m"
Printer.Print Tab(20); "Distancia do L.R ao P.N:"; Tab(70); Form10.LRPN; " m"
Printer.Print Tab(20); "Distancia do L.R ao C.A da asa:"; Tab(70); Form10.LRCAasa; " m"
Printer.Print Tab(20); "Distancia do L.R ao C.A do Etb.Horz.:"; Tab(70); Form10.LRCAht; " m"
Printer.Print Tab(20); "Distancia do L.R ao C.G:"; Tab(70); Form10.LRCG; " m"
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


'ESTRUTURA DA ASA
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(15); "V - ESTRUTURA DA ASA"
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 10
Printer.Print ""
Printer.Print Tab(25); "ABREVIATURAS"
Printer.Print Tab(20); "CL = Coeficiente de sustentação do perfil"""
Printer.Print Tab(20); "Q = Carga aerodinâmica "
Printer.Print Tab(20); "M = Momento fletor"
Printer.Print Tab(20); "V = Esforço cortante"
Printer.Print Tab(20); "L = Largura da viga"
Printer.Print Tab(20); "H = Altura da viga"
Printer.Print Tab(20); "e = Espessura da viga"
Printer.Print ""
Printer.Print ""
Printer.Print Tab(25); "TABELA"
Printer.Print Tab(20); "ESTAÇÃO:"; Tab(40); "LOCAL"; Tab(55); "CL"; Tab(70); "Q (kgf)"; Tab(85); "M (kgf x m)"; Tab(100); "V (kgf)"
Printer.Print ""
Printer.Print Tab(20); "Estação 01:"; Tab(40); Form11.e0; " m"; Tab(55); Form11.cl0; Tab(70); Form11.q0; Tab(85); Form11.m0; Tab(100); Form11.v0
Printer.Print Tab(20); "Estação 02:"; Tab(40); Form11.e1; " m"; Tab(55); Form11.cl1; Tab(70); Form11.q1; Tab(85); Form11.m1; Tab(100); Form11.v1
Printer.Print Tab(20); "Estação 03:"; Tab(40); Form11.e2; " m"; Tab(55); Form11.cl2; Tab(70); Form11.q2; Tab(85); Form11.m2; Tab(100); Form11.v2
Printer.Print Tab(20); "Estação 04:"; Tab(40); Form11.e3; " m"; Tab(55); Form11.cl3; Tab(70); Form11.q3; Tab(85); Form11.m3; Tab(100); Form11.v3
Printer.Print Tab(20); "Estação 05:"; Tab(40); Form11.e4; " m"; Tab(55); Form11.cl4; Tab(70); Form11.q4; Tab(85); Form11.m4; Tab(100); Form11.v4
Printer.Print Tab(20); "Estação 06:"; Tab(40); Form11.e5; " m"; Tab(55); Form11.cl5; Tab(70); Form11.q5; Tab(85); Form11.m5; Tab(100); Form11.v5
Printer.Print Tab(20); "Estação 07:"; Tab(40); Form11.e6; " m"; Tab(55); Form11.cl6; Tab(70); Form11.q6; Tab(85); Form11.m6; Tab(100); Form11.v6
Printer.Print Tab(20); "Estação 08:"; Tab(40); Form11.e7; " m"; Tab(55); Form11.cl7; Tab(70); Form11.q7; Tab(85); Form11.m7; Tab(100); Form11.v7
Printer.Print Tab(20); "Estação 09:"; Tab(40); Form11.e8; " m"; Tab(55); Form11.cl8; Tab(70); Form11.q8; Tab(85); Form11.m8; Tab(100); Form11.v8
Printer.Print Tab(20); "Estação 10:"; Tab(40); Form11.e9; " m"; Tab(55); Form11.cl9; Tab(70); Form11.q9; Tab(85); Form11.m9; Tab(100); Form11.v9
Printer.Print Tab(20); "Estação 11:"; Tab(40); Form11.e10; " m"; Tab(55); Form11.cl10; Tab(70); Form11.q10; Tab(85); Form11.m10; Tab(100); Form11.v10
Printer.Print ""
Printer.Print ""
Printer.Print Tab(25); "DIMENSIONAMENTO"
Printer.Print ""
Printer.Print Tab(20); "Secção transversal da longarina tipo:"; Tab(70); Form11.section;
Printer.Print Tab(20); "Mesas da longarina em madeira de:"; Tab(70); Form11.materialM;
Printer.Print Tab(20); "Alma da longarina em madeira de:"; Tab(70); Form11.materialA;
Printer.Print Tab(20); "Tensão admissível a tração - mesa:"; Tab(70); Form11.sigtrac; " kgf/cm2"
Printer.Print Tab(20); "Tensão admissível a compressão - mesa:"; Tab(70); Form11.sigcomp; " kgf/cm2"
Printer.Print Tab(20); "Tensão admissível ao cizalhamento:"; Tab(70); Form11.sigcizal; " %"
Printer.Print Tab(20); "Altura da viga:"; Tab(70); Form11.altmax; " cm"
Printer.Print Tab(20); "Largura da viga:"; Tab(70); Form11.largamax; " cm"
Printer.Print Tab(20); "Mesa superior (L x H):"; Tab(70); Form11.mesacp; " cm"
Printer.Print Tab(20); "Mesa inferior (L x H):"; Tab(70); Form11.mesatc; " cm"
Printer.Print Tab(20); "Alma (e x H):"; Tab(70); Form11.alma; " cm"
Printer.Print Tab(20); "Nº de Gs positivos:"; Tab(70); Form13.gsatual;
Printer.Print Tab(20); "Nº de Gs negativos:"; Tab(70); Form13.gsnegatual;
Printer.Print ""

Printer.Print ""
Printer.Print ""
Printer.Print Tab(20); "Fim da página";
Printer.EndDoc

'PESO E BALANCEAMENTO
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
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

' Trimagem - form27
If Form27.Option3.Value = True Then regime = "Subida"
If Form27.Option4.Value = True Then regime = "Veloc. máxima"
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

Private Sub filelateral_Click()

filelinha.Enabled = True
fileapagalin.Enabled = True

Form18.dondecol.Enabled = True
Form18.donestol.Enabled = True
Form18.donsub.Enabled = True
Form18.donvmax.Enabled = True
Form18.filelimpa.Enabled = True

fileplanta.Enabled = True
linhaasa1.Visible = False
linhaleme.Visible = True

basa.Visible = False
bfasa.Visible = False
baht.Visible = False
bfht.Visible = False
pontasa.Visible = False
pontht.Visible = False
medasa.Visible = False
raizasa.Visible = False
raizht.Visible = False
medht.Visible = False

    Form18.linhaX.Visible = True
    Form18.linhaif.Visible = True
    Form18.linhasf.Visible = True
    Form18.linhaY.Visible = True
    Form18.linhaleme.Visible = True
    linhaasa1.Visible = False
    
'Capacitando os comandos de equipamentos
Command5.Enabled = True

'Call Command5_Click

Command10.Enabled = True
linhaasa1.Visible = False

Form18.linhaY.x1 = 1680
Form18.linhaY.X2 = 1680

vslinha.Enabled = True

Dim fatx, escalax As Double

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

Form18.linhaX.y1 = posilinha
Form18.linhaX.y2 = posilinha


'......... Linha do leme
linhaleme.y1 = linhaX.y1 - (Form2.bvt * escalax)
linhaleme.y2 = linhaX.y2 - (Form2.bvt * escalax)

'......... Localização das lihas da fuselagem

    If Form1.posimedia.Value = True Then

        linhasf.y1 = linhaX.y1 - ((Form2.hfs * 0.6) * escalax)
        linhasf.y2 = linhaX.y1 - ((Form2.hfs * 0.6) * escalax)

        linhaif.y1 = linhaX.y1 + ((Form2.hfs * 0.4) * escalax)
        linhaif.y2 = linhaX.y1 + ((Form2.hfs * 0.4) * escalax)

    End If

        If Form1.posibaixa.Value = True Or Form1.posialta.Value = True Then
            linhasf.y1 = linhaX.y1 - ((Form2.hfs * 0.4) * escalax)
            linhasf.y2 = linhaX.y1 - ((Form2.hfs * 0.4) * escalax)

            linhaif.y1 = linhaX.y1 + ((Form2.hfs * 0.6) * escalax)
            linhaif.y2 = linhaX.y1 + ((Form2.hfs * 0.6) * escalax)
        End If
        
        linhafinal.Visible = True

linhafinal.x1 = linhaY.x1 + (Form2.lfs * escalax)
linhafinal.X2 = linhaY.x1 + (Form2.lfs * escalax)

linhafinal.y1 = linhaX.y1 - ((Form2.lfs * 0.025) * escalax)
linhafinal.y2 = linhaX.y1 + ((Form2.lfs * 0.025) * escalax)


' ESTABILIZADOR VERTICAL -

bavt.Visible = True
bfvt.Visible = True

'Fator de escala
fatx = Form18.esc18
escalax = 800 * fatx

'Posição X

Dim kvt As Double
gradvt = HScroll3.Value

kvt = vt

bavt.x1 = linhafinal.x1 - (kvt * Form2.crvt * escalax)
bavt.X2 = bavt.x1 + (gradvt * 15 * fatx)

bfvt.x1 = bavt.x1 + (Form2.crvt * escalax)
bfvt.X2 = bavt.X2 + (Form2.cpvt * escalax)




' Posição Y

bavt.y1 = linhaX.y1
bfvt.y1 = linhaX.y1

bavt.y2 = linhaleme.y1
bfvt.y2 = linhaleme.y1


'Local da asa média
If Form21.Data1.Recordset.RecordCount < 1 Then GoTo 10

If Form1.posimedia.Value = True And Form21.Data1.Recordset("Projeto") = Form20.p1 Then

    If Form18.Option3.Value = True Then Form18.VScroll3.Value = Form18.VScroll3.Value
            If Form18.Option3.Value = False Then
                Form18.VSasamed = Form21.Data1.Recordset("pmed")
                If Form21.Data1.Recordset("pmed") = 0 Then Form18.VSasamed = 6
                Form18.VScroll3 = Form18.VSasamed
            End If
End If

If Form1.posimedia.Value = True And Form21.Data1.Recordset("Projeto") <> Form20.p1 Then
    If Form18.Option3.Value = True Then Form18.VScroll3.Value = Form18.VScroll3.Value
        If Form18.Option3.Value = False Then
10:
            Form18.VScroll3.Value = 4
            Form18.VScroll3.Value = 6
        End If
End If

End Sub

Private Sub filelimpa_Click()
Dim Eps, fatx, escalax, cordaw, distw, distcam, dist, k, l9, ed, EpsonD, EpsonSB, EpsonVM, ya8, cordaht, ktail, dwx1, dwy, dwy1, dwx2, dwy2n, dwy2p, deflex, EpsonS, cd, Fa, Fb1, Fb2, Fb, Ew, cdmax As Double

                ' STALL
'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

'Efeito Diedro - ed tangente de 1º = 0.017453
ed = ((Form10.diedro * 0.017453) * Form2.bw / 2) * escalax

If Form1.posialta.Value = True Then k = linhasf.y1 - (ed / 2)
If Form1.posimedia.Value = True Then k = Form18.valasa - (ed / 2)
If Form1.posibaixa.Value = True Then k = linhaif.y1 - (ed / 2)

cordaw = Form7.CMaerodw * escalax
distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
ya8 = 0.078 * cordaw
l9 = 0.3 * cordaw

'Dados de geometria das superficies transformadas
'Corda da asa
cordaht = Form7.CMaerodht * escalax
distcam = distw + (0.25 * cordaw)
dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)

' Deflexão da linha devido a angulo downwash
EpsonS = 35 * Form7.clk1 / Form1.arw
EpsonS = Format(EpsonS, "0.000")

deflex = (Tan(EpsonS * (3.141516 / 180))) * (distcam + (Form10.LT * escalax))

stylesolid.Checked = True
    Form18.DrawStyle = 0
'Linha média
Line (distcam, k - (0.25 * ya8))-(distcam + (Form10.LT * escalax), k - (0.25 * ya8) + deflex), &HFFFFFF

stylesolid.Checked = True
    Form18.DrawStyle = 0

dwx1 = distcam
dwy = k - (0.25 * ya8)
dwy1 = k - (0.25 * ya8) + deflex

'Espessura do Downwash
cdmax = Form4.Data1.Recordset("cdmax")

' METODO STINTON
'cd = Abs((Form7.clk1 ^ 2) / (3.1415 * Form1.arw * 0.9) + Abs(cdmax))
'Fa = ((cd ^ 0.5) / (5 * 3.141516))
'Fb1 = Form10.LT / 0.305
'Fb2 = Form2.cmw / 0.305
'Fb = (Form1.arw * (Fb1 / Fb2)) ' A fórmula correta seria: Fb = (Form1.arw * (Fb1 / Fb2)) +7
'Ew = ((Fa * Fb) * (Form2.cmw / 0.305)) * 0.305
'Ew = Format(Ew, "0.000")

'METODO DOMMASCH
Eps = 20 * Form7.clk1 * (((1 / Form1.Afilasa) ^ 0.3) / (Form1.arw ^ 0.725)) * ((3 * Form7.CMaerodw / Form10.LT) ^ 0.25)
Ew = (Tan(Eps * 3.1415 / 180)) * Form10.LT
Ew = Format(Ew, "0.000")

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax

' Para stall
    Line (dwx1, dwy)-(dwx2, dwy2n), &HFFFFFF
    Line (dwx1, dwy)-(dwx2, dwy2p), &HFFFFFF
    
                     'DECOLAGEM

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

ed = ((Form10.diedro * 0.017453) * Form2.bw / 2) * escalax

If Form1.posialta.Value = True Then k = linhasf.y1 - (ed / 2)
If Form1.posimedia.Value = True Then k = Form18.valasa - (ed / 2)
If Form1.posibaixa.Value = True Then k = linhaif.y1 - (ed / 2)

cordaw = Form7.CMaerodw * escalax
distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
ya8 = 0.078 * cordaw
l9 = 0.3 * cordaw

'Dados de geometria das superficies transformadas
'Corda da asa
cordaht = Form7.CMaerodht * escalax
distcam = distw + (0.25 * cordaw)
dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)

' Deflexão da linha devido a angulo downwash
EpsonD = 35 * Form7.cldecol / Form1.arw
EpsonD = Format(EpsonD, "0.000")
deflex = (Tan(EpsonD * (3.141516 / 180))) * (distcam + (Form10.LT * escalax))

stylesolid.Checked = True
    Form18.DrawStyle = 0
'Linha média
Line (distcam, k - (0.25 * ya8))-(distcam + (Form10.LT * escalax), k - (0.25 * ya8) + deflex), &HFFFFFF

stylesolid.Checked = True
    Form18.DrawStyle = 0

dwx1 = distcam
dwy = k - (0.25 * ya8)
dwy1 = k - (0.25 * ya8) + deflex

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax

'Espessura do Downwash
cdmax = Form4.Data1.Recordset("cdmax")

'METODO STINTON
'cd = Abs((Form7.cldecol ^ 2) / (3.1415 * Form1.arw * 0.9) + Abs(cdmax))
'Fa = ((cd ^ 0.5) / (5 * 3.141516))
'Fb1 = Form10.LT / 0.305
'Fb2 = Form2.cmw / 0.305
'Fb = (Form1.arw * (Fb1 / Fb2)) ' A fórmula correta seria: Fb = (Form1.arw * (Fb1 / Fb2)) +7
'Ew = ((Fa * Fb) * (Form2.cmw / 0.305)) * 0.305
'Ew = Format(Ew, "0.000")

'METODO DOMMASCH
Eps = 20 * Form7.cldecol * (((1 / Form1.Afilasa) ^ 0.3) / (Form1.arw ^ 0.725)) * ((3 * Form7.CMaerodw / Form10.LT) ^ 0.25)
Ew = (Tan(Eps * 3.1415 / 180)) * Form10.LT
Ew = Format(Ew, "0.000")

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax


' Para decolagem
    Line (dwx1, dwy)-(dwx2, dwy2n), &HFFFFFF
    Line (dwx1, dwy)-(dwx2, dwy2p), &HFFFFFF

                    ' SUBIDA
'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

ed = ((Form10.diedro * 0.017453) * Form2.bw / 2) * escalax

If Form1.posialta.Value = True Then k = linhasf.y1 - (ed / 2)
If Form1.posimedia.Value = True Then k = Form18.valasa - (ed / 2)
If Form1.posibaixa.Value = True Then k = linhaif.y1 - (ed / 2)

cordaw = Form7.CMaerodw * escalax
distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
ya8 = 0.078 * cordaw
l9 = 0.3 * cordaw

'Dados de geometria das superficies transformadas
'Corda da asa
cordaht = Form7.CMaerodht * escalax
distcam = distw + (0.25 * cordaw)
dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)

' Deflexão da linha devido a angulo downwash
EpsonSB = 35 * Form7.clk2 / Form1.arw
EpsonSB = Format(EpsonSB, "0.000")
deflex = (Tan(EpsonSB * (3.141516 / 180))) * (distcam + (Form10.LT * escalax))

stylesolid.Checked = True
    Form18.DrawStyle = 0

'Linha média
Line (distcam, k - (0.25 * ya8))-(distcam + (Form10.LT * escalax), k - (0.25 * ya8) + deflex), &HFFFFFF

stylesolid.Checked = True
    Form18.DrawStyle = 0

dwx1 = distcam
dwy = k - (0.25 * ya8)
dwy1 = k - (0.25 * ya8) + deflex

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax

'Espessura do Downwash
cdmax = Form4.Data1.Recordset("cdmax")

'METODO STINTON
'cd = Abs((Form7.clk2 ^ 2) / (3.1415 * Form1.arw * 0.9) + Abs(cdmax))
'Fa = ((cd ^ 0.5) / (5 * 3.141516))
'Fb1 = Form10.LT / 0.305
'Fb2 = Form2.cmw / 0.305
'Fb = (Form1.arw * (Fb1 / Fb2)) ' A fórmula correta seria: Fb = (Form1.arw * (Fb1 / Fb2)) +7
'Ew = ((Fa * Fb) * (Form2.cmw / 0.305)) * 0.305
'Ew = Format(Ew, "0.000")

'METODO DOMMASCH
Eps = 20 * Form7.clk2 * (((1 / Form1.Afilasa) ^ 0.3) / (Form1.arw ^ 0.725)) * ((3 * Form7.CMaerodw / Form10.LT) ^ 0.25)
Ew = (Tan(Eps * 3.1415 / 180)) * Form10.LT
Ew = Format(Ew, "0.000")

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax


' Para subida
    Line (dwx1, dwy)-(dwx2, dwy2n), &HFFFFFF
    Line (dwx1, dwy)-(dwx2, dwy2p), &HFFFFFF
    
                'VELOCIDADE MÁXIMA
                

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

ed = ((Form10.diedro * 0.017453) * Form2.bw / 2) * escalax

If Form1.posialta.Value = True Then k = linhasf.y1 - (ed / 2)
If Form1.posimedia.Value = True Then k = Form18.valasa - (ed / 2)
If Form1.posibaixa.Value = True Then k = linhaif.y1 - (ed / 2)

cordaw = Form7.CMaerodw * escalax
distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
ya8 = 0.078 * cordaw
l9 = 0.3 * cordaw

'Dados de geometria das superficies transformadas
'Corda da asa
cordaht = Form7.CMaerodht * escalax
distcam = distw + (0.25 * cordaw)
dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)

' Deflexão da linha devido a angulo downwash
EpsonVM = 35 * Form7.clk4 / Form1.arw
EpsonVM = Format(EpsonVM, "0.000")
deflex = (Tan(EpsonVM * (3.141516 / 180))) * (distcam + (Form10.LT * escalax))

stylesolid.Checked = True
    Form18.DrawStyle = 0

'Linha média
Line (distcam, k - (0.25 * ya8))-(distcam + (Form10.LT * escalax), k - (0.25 * ya8) + deflex), &HFFFFFF

stylesolid.Checked = True
    Form18.DrawStyle = 0


dwx1 = distcam
dwy = k - (0.25 * ya8)
dwy1 = k - (0.25 * ya8) + deflex

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax

'Espessura do Downwash
cdmax = Form4.Data1.Recordset("cdmax")

'METODO STINTON
'cd = Abs((Form7.clk4 ^ 2) / (3.1415 * Form1.arw * 0.9) + Abs(cdmax))
'Fa = ((cd ^ 0.5) / (5 * 3.141516))
'Fb1 = Form10.LT / 0.305
'Fb2 = Form2.cmw / 0.305
'Fb = (Form1.arw * (Fb1 / Fb2)) ' A fórmula correta seria: Fb = (Form1.arw * (Fb1 / Fb2)) +7
'Ew = ((Fa * Fb) * (Form2.cmw / 0.305)) * 0.305
'Ew = Format(Ew, "0.000")

'METODO DOMMASCH
Eps = 20 * Form7.clk4 * (((1 / Form1.Afilasa) ^ 0.3) / (Form1.arw ^ 0.725)) * ((3 * Form7.CMaerodw / Form10.LT) ^ 0.25)
Ew = (Tan(Eps * 3.1415 / 180)) * Form10.LT
Ew = Format(Ew, "0.000")

dwx2 = (distcam + (Form10.LT * escalax))
dwy2n = dwy1 - (Ew / 2) * escalax
dwy2p = dwy1 + (Ew / 2) * escalax


' Para cruzeiro
    Line (dwx1, dwy)-(dwx2, dwy2n), &HFFFFFF
    Line (dwx1, dwy)-(dwx2, dwy2p), &HFFFFFF


    Form18.MousePointer = 2
    Shape = "LINE"
    Form18.DrawStyle = 0
    
    ' PARA RESTAURAR AS SUPERFÍCIES
Call Command5_Click
Call Command10_Click

End Sub

Private Sub filelinha_Click()
Dim fatx, escalax, tango, tango1, cordaht, distcam, dist, L1, l18, ktail, distw, cordaw As Double
L1 = 0
'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx
'..................AEROFOLIO DA EMPENAGEM
cordaht = Form7.CMaerodht * escalax
cordaw = Form7.CMaerodw * escalax

distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
distcam = distw + (0.25 * cordaw)
dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

cordaht = Form7.CMaerodht * escalax

dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)

tango = 1.732 * ((1 + dist) - (1 + dist + (escalax)))  ' Para 60º
tango1 = 0.57735 * ((1 + dist) - (1 + dist + (escalax)))  ' Para 30º

'TRAÇADO DA LINHA
' Bordo de ataque
Form18.Line (L1 + dist, ktail)-(1 + dist + (escalax), ktail + tango), &HFF00FF

'Bordo de fuga

Form18.Line (cordaht + dist, ktail)-(cordaht + dist + (escalax), ktail + tango1), &HFF00FF



End Sub

Private Sub filelinhaux_Click()
linhaleme.Visible = False
linhasf.Visible = False
linhaX.Visible = False
linhaY.Visible = False
linhaif.Visible = False
linhaasa1.Visible = False
linhafinal.Visible = False
basa.Visible = False
bfasa.Visible = False
baht.Visible = False
bfht.Visible = False
pontasa.Visible = False
pontht.Visible = False
medasa.Visible = False
raizasa.Visible = False
raizht.Visible = False
medht.Visible = False
bavt.Visible = False
bfvt.Visible = False
End Sub

Private Sub FileNew_Click()
outros.Value = Checked
    Form18.Picture = LoadPicture()
    OpenFile = ""
      
End Sub

Private Sub FileOpen_Click()
outros.Value = Unchecked

Dim respdes As Double
Dim LocalSave As String
    LocalSave = App.Path + "\" + Form20.p1 + "\" + Trim(Trim(Form20.p1) + Trim(".bmp"))
    If Dir(LocalSave) = "" Then
        MsgBox "Não existe desenho salvo para este projeto", 16, "Erro na abertura do desenho"
        Exit Sub
    End If
    CommonDialog1.FileName = LocalSave
    Form18.Picture = LoadPicture(CommonDialog1.FileName)
    
    Picture1.Picture = Form18.Picture
    
    Form18.linhaX.Visible = False
    Form18.linhaif.Visible = False
    Form18.linhasf.Visible = False
    Form18.linhaY.Visible = False
    Form18.linhaleme.Visible = False
    
    'Abrindo desenho
    
respdes = MsgBox("Deseja manter a escala original", vbYesNo)

If respdes = vbYes Then

Call Command34_Click

    Form18.esc18 = Form21.Data1.Recordset("escala18")
    Form18.vsescala.Value = 10 * Form18.esc18
    
    'FORM18 desenho se não existir salvo com o nome do projeto
    'If Form21.Data1.Recordset("AfilamW") = Empty And Form21.Data1.Recordset("Afilamht") = Empty And Form21.Data1.Recordset("f18hs1") = Empty And Form21.Data1.Recordset("f18hs2") = Empty And Form21.Data1.Recordset("f18hs3") = Empty And Form21.Data1.Recordset("f18hs4") = Empty Then
        'MsgBox " O desenho que será aberto não pertence ao projeto"
        'Exit Sub
    'End If
    
  ' FORM18 desenho se existir salvo
    Form18.afilaASA = Form21.Data1.Recordset("AfilamW")
    Form18.AfilaHt = Form21.Data1.Recordset("Afilamht")
     
    Form18.hsasaval = Form21.Data1.Recordset("f18hs1")
    Form18.HScroll1.Value = Form18.hsasaval
    
    Form18.hshtval = Form21.Data1.Recordset("f18hs2")
    Form18.HScroll2.Value = Form18.hshtval
    
    Form18.gradvt = Form21.Data1.Recordset("f18hs3")
    Form18.HScroll3.Value = Form18.gradvt
    
    Form18.vt = Form21.Data1.Recordset("f18hs4")
  
    Form18.hs4 = Form21.Data1.Recordset("Fo18hs4")
    Form18.HScroll4.Value = Form18.hs4
    
    
End If

If respdes = vbNo Then
    Form18.esc18 = 1
    Form18.vsescala.Value = 10 * Form18.esc18
End If



End Sub

Private Sub fileopout_Click()
outros.Value = Checked
    CommonDialog1.Filter = "Images|*.bmp;*.gif;*.jpg"
    CommonDialog1.DefaultExt = "BMP"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then Exit Sub
    Form18.Picture = LoadPicture(CommonDialog1.FileName)
    OpenFile = CommonDialog1.FileName
    Picture1.Picture = Form18.Picture
    
End Sub

Private Sub filepadrao_Click()
Frame22.Visible = False
Frame6.Visible = True
Label3.Top = 1440
med2.Top = 1680
End Sub

Private Sub filePB_Click()
Form14.Visible = True
Form18.Visible = False
Form14.Refresh
End Sub

Private Sub fileperform18_Click()
Form8.Visible = True
Form18.Visible = False
Form8.Refresh
End Sub

Private Sub fileplanta_Click()
Dim Dxht, fatx, escalax, cordahtR, cordahtP, flexht, porcentoht As Double
Dim cordaw, cordawP, cordawR, cordaht, distw, dist, l2, l18, distca As Double
Dim dxasa, porcentoa, flexasa As Double

Form18.bavt.Visible = False
Form18.bfvt.Visible = False

filelinha.Enabled = False
fileapagalin.Enabled = False

Form18.dondecol.Enabled = False
Form18.donestol.Enabled = False
Form18.donsub.Enabled = False
Form18.donvmax.Enabled = False
Form18.filelimpa.Enabled = False

'Form18.Cls

linhaleme.Visible = True
linhasf.Visible = True
linhaX.Visible = True
linhaY.Visible = True
linhaif.Visible = True
linhafinal.Visible = True

basa.Visible = True
bfasa.Visible = True
baht.Visible = True
bfht.Visible = True
pontasa.Visible = True
pontht.Visible = True
medasa.Visible = True
raizasa.Visible = True
raizht.Visible = True
medht.Visible = True

'Incapacitando os comandos de equipamentos
Command5.Enabled = False
Command10.Enabled = False

linhaleme.Visible = False

Form18.linhaY.x1 = 1680
Form18.linhaY.X2 = 1680

fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

vslinha.Enabled = False


'Dim escalax As Double

linhafinal.Visible = True

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx
Form18.linhaX.y1 = 7000
Form18.linhaX.y2 = 7000

Form18.linhaif.y1 = Form18.linhaX.y1 + (Form2.wfs / 2) * escalax
Form18.linhaif.y2 = Form18.linhaX.y1 + (Form2.wfs / 2) * escalax

Form18.linhasf.y1 = Form18.linhaX.y1 - (Form2.wfs / 2) * escalax
Form18.linhasf.y2 = Form18.linhaX.y1 - (Form2.wfs / 2) * escalax

linhafinal.x1 = linhaY.x1 + (Form2.lfs * escalax)
linhafinal.X2 = linhaY.x1 + (Form2.lfs * escalax)

'Linha fina em planta para fuselagem

linhafinal.y1 = linhasf.y1
linhafinal.y2 = linhaif.y1

'ASA

cordaw = Form7.CMaerodw * escalax
cordawR = Form2.crw * escalax
cordawP = Form2.cpw * escalax


l2 = 0.0125 * cordaw
l18 = 1 * cordaw

distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
 
basa.x1 = l2 + distw
basa.X2 = l2 + distw

basa.y1 = linhaX.y1
basa.y2 = linhaX.y1 - ((Form2.bw / 2) * escalax)


bfasa.x1 = cordawR + distw
bfasa.X2 = cordawP + distw

bfasa.y1 = linhaX.y1
bfasa.y2 = linhaX.y1 - ((Form2.bw / 2) * escalax)


'Corda de ponta
pontasa.x1 = basa.x1
pontasa.X2 = bfasa.X2

pontasa.y1 = linhaX.y1 - ((Form2.bw / 2) * escalax)
pontasa.y2 = linhaX.y1 - ((Form2.bw / 2) * escalax)

'Corda média
medasa.y1 = linhaX.y1 - ((Form2.bw / 4) * escalax)
medasa.y2 = linhaX.y1 - ((Form2.bw / 4) * escalax)

medasa.x1 = l2 + distw
medasa.X2 = l18 + distw

'Corda da raiz

raizasa.x1 = basa.x1
raizasa.X2 = bfasa.x1

raizasa.y1 = linhaX.y1
raizasa.y2 = linhaX.y1


'EMPENAGEM
cordaht = Form7.CMaerodht * escalax
cordahtR = Form2.crht * escalax
cordahtP = Form2.cpht * escalax


l2 = 0.00125 * cordaht
l18 = cordaht

distca = distw + (0.25 * cordaw)
dist = distca + ((Form10.LT * escalax) - (0.25 * cordaht))

baht.x1 = l2 + dist
baht.X2 = l2 + dist

baht.y1 = linhaX.y1
baht.y2 = linhaX.y1 - ((Form2.bht / 2) * escalax)

bfht.x1 = cordahtR + dist
bfht.X2 = cordahtP + dist

bfht.y1 = linhaX.y1
bfht.y2 = linhaX.y1 - ((Form2.bht / 2) * escalax)

'CORDA DE PONTA HT
pontht.x1 = baht.x1
pontht.X2 = bfht.X2

pontht.y1 = linhaX.y1 - ((Form2.bht / 2) * escalax)
pontht.y2 = linhaX.y1 - ((Form2.bht / 2) * escalax)

'CORDA MÉDIA HT
medht.y1 = linhaX.y1 - ((Form2.bht / 4) * escalax)
medht.y2 = linhaX.y1 - ((Form2.bht / 4) * escalax)

medht.x1 = l2 + dist
medht.X2 = l18 + dist

'CORDA DA RAIZ HT

raizht.x1 = baht.x1
raizht.X2 = bfht.x1

raizht.y1 = linhaX.y1
raizht.y2 = linhaX.y1

If HScroll1 > 1 Then Call HScroll1_Change
If HScroll2 > 1 Then Call HScroll2_Change

End Sub

Private Sub fileprincipia18_Click()
    Form18.AutoRedraw = True
    Form1.Visible = True
    Form18.Visible = False
    
    Form1.Refresh
End Sub


Private Sub fileroda_Click()

End Sub

Private Sub filerecuper_Click()
Form30.Show
End Sub

Private Sub filesai18_Click()
Form1.Visible = True
Form18.Visible = False

Form1.Refresh
End Sub

Private Sub FileSave_Click()
If outros.Value = Checked Then
    Call FileSaveAs_Click
    Exit Sub
End If

'Salvando dados do Est.Vert.
Form21.Data1.Recordset.Edit
Form21.Data1.Recordset("Fo18hs4") = Form18.hs4
Form21.Data1.Recordset("f18hs4") = Form18.vt
Form21.Data1.Recordset.Update

Dim LocalSave As String
'Acrescentei aqui
Dim variavel As String
variavel = "Select * from Salvar where Projeto=Projeto"
            Form21.Data1.RecordSource = variavel
            Form21.Data1.Refresh
'Até aqui para posicionar o nome do projeto

    LocalSave = App.Path + "\" + Form20.p1 + "\" + Trim(Trim(Form20.p1) + Trim(".bmp"))
 On Error Resume Next
        MkDir App.Path + "\" + Form20.p1
    CommonDialog1.FileName = LocalSave
    SavePicture Form18.Image, CommonDialog1.FileName
   
    MsgBox "Os desenhos foram salvos. Não esqueça de salvar ou atualizar o seu projeto.", vbInformation
    
End Sub


Private Sub FileSaveAs_Click()
MsgBox "Crie uma pasta com o nome de projeto e salve nesta os seus outros desenhos", vbInformation
    CommonDialog1.Filter = "Images|*.bmp"
    CommonDialog1.DefaultExt = "BMP"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then Exit Sub
    SavePicture Form18.Image, CommonDialog1.FileName
    OpenFile = CommonDialog1.FileName
End Sub



Private Sub filetrim_Click()
Form18.Visible = False
Form27.Visible = True
End Sub

Private Sub Form_Load()

Label3.BackColor = Form18.BackColor
Label2.BackColor = Form18.BackColor
med1.BackColor = Form18.BackColor
med1.BackColor = Form18.BackColor

    CopyBMP = False
    PasteBMP = False
    PrintText = False
    
    posilinha = Form18.vslinha.Value
    
'Barra de ferramentas
Frame1.Visible = False
Label3.Top = 120
med2.Top = 360

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim escalax, fatx As Double

fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

    
    If Button = 2 Then Shape = ""
    If Button = 1 Then
        XStart = X
        YStart = Y
        XPrevious = XStart
        YPrevious = YStart
        Form18.AutoRedraw = False
    End If
    
    If CopyBMP Or CutBMP Then
        PDrawWidth = Form18.DrawWidth
        PDrawStyle = Form18.DrawStyle
        PFillStyle = Form18.FillStyle
        Form18.DrawWidth = 1
        Form18.DrawStyle = 0
        Form18.FillStyle = 1
    End If
    
    If PasteBMP Then
    If CopyWidth = 0 Or CopyHeight = 0 Then
    Form18.Refresh
    Exit Sub
    End If
    
        Form18.PaintPicture Picture1.Image, X, Y, CopyWidth, CopyHeight, 0, 0, CopyWidth, CopyHeight, &HCC0020
        XPrevious = X
        YPrevious = Y
        Exit Sub
    End If
    
If PrintText Then
    Label1.Visible = True
    Label1.Left = X
    Label1.Top = Y
    Exit Sub
End If

'Para iniciar desenhos
    Form18.Label2 = (X - 1200) / escalax
    Form18.Label2 = Format(Form18.Label2, "0.000")
    Form18.Label3 = Y / escalax
    Form18.Label3 = Format(Form18.Label3, "0.000")
    
    xx = Label2
    yy = Label3
    
    med1.Visible = False
    med2.Visible = False
    
    'CURSOR
    If Button <> 1 Then Exit Sub
    Form18.Label2 = (X - 1200) / escalax
    Form18.Label2 = Format(Form18.Label2, "0.000")
    Form18.Label3 = Y / escalax
    Form18.Label3 = Format(Form18.Label3, "0.000")
    'FIM CURSOR
                 
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim escalax, fatx, raio As Double
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

    If Button <> 1 Then Exit Sub
    If CopyBMP Or CutBMP Then
        Form18.Line (XStart, YStart)-(XPrevious, YPrevious), , B
        Form18.Refresh
        Form18.Line (XStart, YStart)-(X, Y), , B
        XPrevious = X
        YPrevious = Y
        Exit Sub
    End If
    
    If PasteBMP Then
        Form18.PaintPicture Picture1.Image, XPrevious, YPrevious, CopyWidth, CopyHeight, 0, 0, CopyWidth, CopyHeight, &HFFFFFF
        Form18.Refresh
        Form18.PaintPicture Picture1.Image, X, Y, CopyWidth, CopyHeight, 0, 0, CopyWidth, CopyHeight, &HCC0020
        Exit Sub
    End If
    
    If PrintText Then
    Label1.Left = X
    Label1.Top = Y
    Exit Sub
    End If

    Select Case Shape
        Case "LINE":
        Form18.Refresh
            Form18.Line (XStart, YStart)-(X, Y)
            Form18.Line (X, Form18.ScaleTop)-(X, Form18.ScaleTop + Form18.ScaleHeight), &H0&
            Form18.Line (Form18.ScaleLeft, Y)-(Form18.ScaleLeft + Form18.ScaleWidth, Y), &H0&
                  
        Case "LINE2":
        Form18.Refresh
            Form18.Line (XStart, YStart)-(X, YStart)
            Form18.Line (X, Form18.ScaleTop)-(X, Form18.ScaleTop + Form18.ScaleHeight), &H0&
            Form18.Line (Form18.ScaleLeft, Y)-(Form18.ScaleLeft + Form18.ScaleWidth, Y), &H0&
            
        Case "LINE3":
        Form18.Refresh
            Form18.Line (XStart, YStart)-(XStart, Y)
            Form18.Line (X, Form18.ScaleTop)-(X, Form18.ScaleTop + Form18.ScaleHeight), &H0&
            Form18.Line (Form18.ScaleLeft, Y)-(Form18.ScaleLeft + Form18.ScaleWidth, Y), &H0&
            
        Case "CIRCLE":
        Form18.Refresh
            Form18.Circle (XStart, YStart), Sqr((X - XStart) ^ 2 + (Y - YStart) ^ 2), 1
            Form18.Line (X, Form18.ScaleTop)-(X, Form18.ScaleTop + Form18.ScaleHeight), &H0&
            Form18.Line (Form18.ScaleLeft, Y)-(Form18.ScaleLeft + Form18.ScaleWidth, Y), &H0&
                        
        Case "BOX":
        Form18.Refresh
            Form18.Line (XStart, YStart)-(X, Y), 1, B
            Form18.Line (X, Form18.ScaleTop)-(X, Form18.ScaleTop + Form18.ScaleHeight), &H0&
            Form18.Line (Form18.ScaleLeft, Y)-(Form18.ScaleLeft + Form18.ScaleWidth, Y), &H0&
            
            
        Case "ELIPSE":
        Form18.Refresh
        If Y = 0 Then Y = 0.001
        If X = 0 Then X = 0.001
        Circle (XStart, YStart), Abs((Abs(Label3 * escalax) - Abs(Label2 * escalax))), , , , X / Y
        Form18.Line (X, Form18.ScaleTop)-(X, Form18.ScaleTop + Form18.ScaleHeight), &H0&
        Form18.Line (Form18.ScaleLeft, Y)-(Form18.ScaleLeft + Form18.ScaleWidth, Y), &H0&
        
        
        Case "apaga"
        Form18.Refresh
            Form18.DrawWidth = 2
            Form18.PSet (X, Y), BackColor
            Form18.DrawWidth = 1
            Form18.Line (X, Form18.ScaleTop)-(X, Form18.ScaleTop + Form18.ScaleHeight), &H0&
            Form18.Line (Form18.ScaleLeft, Y)-(Form18.ScaleLeft + Form18.ScaleWidth, Y), &H0&
            
            
        Case "dimensionX"
        Form18.Refresh
        Form18.Line (XStart, YStart)-(X, YStart)
        
        Case "dimensionY"
        Form18.Refresh
        Form18.Line (XStart, YStart)-(XStart, Y)
        
    End Select
    
    'CURSOR
    If Button <> 1 Then Exit Sub
    Form18.Label2 = (X - 1200) / escalax
    Form18.Label2 = Format(Form18.Label2, "0.000")
    Form18.Label3 = Y / escalax
    Form18.Label3 = Format(Form18.Label3, "0.000")
    'FIM CURSOR
    

        
End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim x1, y1, fatx, escalax, raio As Double

fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx
                   
    If CopyBMP Then
        Form18.Line (XStart, YStart)-(XPrevious, YPrevious), , B
        Form18.Refresh
        If X > XStart Then x1 = XStart Else x1 = X
        If Y > YStart Then y1 = YStart Else y1 = Y
        
        'Para copiar sem quadro
        If Abs(X - XStart) = 0 Or Abs(Y - YStart) = 0 Then
        Exit Sub
        End If
        
        Picture1.PaintPicture Form18.Image, 0, 0, Abs(X - XStart), Abs(Y - YStart), x1, y1, Abs(X - XStart), Abs(Y - YStart), &HCC0020
        CopyBMP = False
        Form18.DrawWidth = PDrawWidth
        Form18.DrawStyle = PDrawStyle
        Form18.FillStyle = PFillStyle
        CopyWidth = Abs(X - XStart)
        CopyHeight = Abs(Y - YStart)
        Exit Sub
    End If
    If CutBMP Then
        Form18.AutoRedraw = True
        CopyWidth = XStart - X
        CopyHeight = YStart - Y
        If X > XStart Then x1 = XStart Else x1 = X
        If Y > YStart Then y1 = YStart Else y1 = Y
        
        'Para colar sem quadro
        If Abs(X - XStart) = 0 Or Abs(Y - YStart) = 0 Then
        Exit Sub
        End If
        
        Picture1.PaintPicture Form18.Image, 0, 0, Abs(X - XStart), Abs(Y - YStart), x1, y1, Abs(X - XStart), Abs(Y - YStart), &HCC0020
        Form18.Line (X, Y)-Step(CopyWidth, CopyHeight), Form18.BackColor, BF
        CutBMP = False
        Form18.DrawWidth = PDrawWidth
        Form18.DrawStyle = PDrawStyle
        Form18.FillStyle = PFillStyle
        CopyWidth = Abs(X - XStart)
        CopyHeight = Abs(Y - YStart)
       
        Exit Sub
    End If
    
    If PasteBMP Then
    
    'Para colar sem quadro
    If Abs(X - XStart) = 0 Or Abs(Y - YStart) = 0 Then
        Exit Sub
        End If
    
        Form18.AutoRedraw = True
        Form18.PaintPicture Picture1.Image, X, Y, CopyWidth, CopyHeight, 0, 0, CopyWidth, CopyHeight, &HCC0020
        PasteBMP = False
        Exit Sub
    End If
    
    If PrintText Then
        Form18.AutoRedraw = True
        Form18.CurrentX = X
        Form18.CurrentY = Y
        Form18.Print Label1.Caption
        Label1.Visible = False
        PrintText = False
        Exit Sub
    End If

    Form18.Refresh
    Form18.AutoRedraw = True
    Select Case Shape
        Case "LINE":
            Form18.Line (XStart, YStart)-(X, Y)
            
        Case "LINE2":
            Form18.Line (XStart, YStart)-(X, YStart)
            
        Case "LINE3":
            Form18.Line (XStart, YStart)-(XStart, Y)
            
        Case "CIRCLE":
            Form18.Circle (XStart, YStart), Sqr((X - XStart) ^ 2 + (Y - YStart) ^ 2)
            
        Case "BOX":
            Form18.Line (XStart, YStart)-(X, Y), , B
            
        Case "ELIPSE":
        Circle (XStart, YStart), Abs(Abs(Label3 * escalax) - Abs(Label2 * escalax)), , , , X / Y
            
        Case "apaga"
                    Form18.DrawWidth = 2
                    Form18.PSet (X, Y), BackColor
                    Form18.DrawWidth = 1
                    
        Case "dimensionX"
                    Form18.ForeColor = CommonDialog1.Color
                    Form18.Line (XStart, YStart)-(X, YStart)
                    Dim DrawString As String
                    Label1.Caption = ((Abs(Abs(Label2) - Abs(xx))) * 100) / 100
                    Label1.Caption = Format(Label1.Caption, "0.000")
                    If Label1.Caption = 0 Then Label1.Caption = ""
                    PrintText = True
                    
        Case "dimensionY"
                    Form18.ForeColor = CommonDialog1.Color
                    Form18.Line (XStart, YStart)-(XStart, Y)
                    Label1.Caption = ((Abs(Abs(Label3) - Abs(yy))) * 100) / 100
                    Label1.Caption = Format(Label1.Caption, "0.000")
                    If Label1.Caption = 0 Then Label1.Caption = ""
                    PrintText = True
                    
                                                                              
        End Select
        
                             
    med1.Visible = True
    med2.Visible = True
    
    med1 = Abs(Label2) - Abs(xx)
    med1 = Format(med1, "0.000")
    med2 = Abs(yy) - Abs(Label3)
    med2 = Format(med2, "0.000")
    
    
    'Usando o botão Direito do Mouse
    If Button = 2 Then
    Form18.MousePointer = 0
    Form18.MousePointer = Form18.MousePointer
    Form18.ForeColor = &HFFFFFF
    Exit Sub
    End If
        
End Sub


Private Sub Form_Resize()

    Picture1.Width = Form18.Width
    Picture1.Height = Form18.Height
    
End Sub



Private Sub HScroll1_Change()
Dim dxasa, fatx, escalax, cordawR, cordawP, flexasa, porcentoa As Double

hsasaval = Form18.HScroll1.Value

'Fator de escala
fatx = Form18.esc18
escalax = 800 * fatx

dxasa = (Abs(Form2.crw * escalax) - Abs(Form2.cpw * escalax))

If dxasa = 0 Then Exit Sub

If dxasa > 0 Then
porcentoa = HScroll1.Value / 100

flexasa = dxasa * porcentoa

'Mudança do bordo de ataque da asa
basa.x1 = medasa.x1 - (flexasa / 2)
basa.X2 = medasa.x1 + (flexasa / 2)
pontasa.x1 = basa.X2

'Mudança do bordo de fuga da asa

cordawR = Form2.crw * escalax
cordawP = Form2.cpw * escalax

bfasa.X2 = medasa.x1 + cordawP + (flexasa / 2)
pontasa.X2 = bfasa.X2
bfasa.x1 = basa.x1 + cordawR
raizasa.x1 = basa.x1
raizasa.X2 = bfasa.x1

grau = (((basa.X2 - basa.x1) / (raizasa.y1 - pontasa.y1)) * 180 / 3.1415162)
grau = Format(grau, "0.00")
End If



End Sub

Private Sub HScroll2_Change()
Dim Dxht, fatx, escalax, cordahtR, cordahtP, flexht, porcentoht As Double

hshtval = Form18.HScroll2.Value

'Fator de escala
fatx = Form18.esc18
escalax = 800 * fatx

Dxht = (Abs(Form2.crht) - Abs(Form2.cpht)) * escalax

If Dxht = 0 Then Exit Sub

If Dxht > 0 Then
porcentoht = HScroll2.Value / 100

flexht = Dxht * porcentoht

'Mudança do bordo de ataque do HT
baht.x1 = medht.x1 - (flexht / 2)
baht.X2 = medht.x1 + (flexht / 2)
pontht.x1 = baht.X2

'Mudança do bordo de fuga do HT

cordahtR = Form2.crht * escalax
cordahtP = Form2.cpht * escalax

bfht.X2 = medht.x1 + cordahtP + (flexht / 2)
pontht.X2 = bfht.X2
bfht.x1 = baht.x1 + cordahtR
raizht.x1 = baht.x1
raizht.X2 = bfht.x1

grauht = ((baht.X2 - baht.x1) / (raizht.y1 - pontht.y1) * 180 / 3.1415162)
grauht = Format(grauht, "0.00")
End If
End Sub



Private Sub HScroll3_Change()

Call filelateral_Click

gradvt = HScroll3.Value
Dim fatx, escalax As Double

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

bavt.X2 = bavt.x1 + (gradvt * 15 * fatx)
bfvt.X2 = bavt.X2 + (Form2.cpvt * escalax)


End Sub



Private Sub HScroll4_Change()
Dim kvt, fatx, escalax As Double
vt = (100 - (HScroll4.Value * 5)) / 100
kvt = vt

Form18.hs4 = Form18.HScroll4.Value

' ESTABILIZADOR VERTICAL -

bavt.Visible = True
bfvt.Visible = True

'Fator de escala
fatx = Form18.esc18
escalax = 800 * fatx

'Posição X

gradvt = HScroll3.Value

kvt = vt

bavt.x1 = linhafinal.x1 - (kvt * Form2.crvt * escalax)
bavt.X2 = bavt.x1 + (gradvt * 15 * fatx)

bfvt.x1 = bavt.x1 + (Form2.crvt * escalax)
bfvt.X2 = bavt.X2 + (Form2.cpvt * escalax)




' Posição Y

bavt.y1 = linhaX.y1
bfvt.y1 = linhaX.y1

bavt.y2 = linhaleme.y1
bfvt.y2 = linhaleme.y1

End Sub

Private Sub insereasa_Click()

'Form18.Cls 'limpando

    Form18.linhaX.Visible = True
    Form18.linhaif.Visible = True
    Form18.linhasf.Visible = True
    Form18.linhaY.Visible = True
    Form18.linhaleme.Visible = True
    linhaasa1.Visible = False
    
    Form18.linhaY.x1 = 1680
    Form18.linhaY.X2 = 1680
    
    
Form18.AutoRedraw = True

Dim distw, distcam, dist, k, ktail, L1, l2, l3, l4, l5, l6, l7, l8, l9, l10, l11, l12, l13, l14, l15, l16, l17, l18, ya1, ya2, ya3, ya4, ya5, ya6, ya7, ya8, ya9, ya10, ya11, ya12, ya13, ya14, ya15, ya16, ya17, ya18 As Double
Dim yb1, yb2, yb3, yb4, yb5, yb6, yb7, yb8, yb9, yb10, yb11, yb12, yb13, yb14, yb15, yb16, yb17, yb18 As Double
Dim cordaw, diam, cordaht, fatx, escalax, ed As Double

'Dim escalax As Double

'Posição da linha de centro
linhaX.y1 = Form18.posilinha
linhaX.y2 = Form18.posilinha

'Escala
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

'......... Linha do leme
linhaleme.y1 = linhaX.y1 - (Form2.bvt * escalax)
linhaleme.y2 = linhaX.y2 - (Form2.bvt * escalax)

'......... Localização das lihas da fuselagem
        If Form1.posimedia.Value = True Then

            linhasf.y1 = linhaX.y1 - ((Form2.hfs * 0.6) * escalax)
            linhasf.y2 = linhaX.y1 - ((Form2.hfs * 0.6) * escalax)

            linhaif.y1 = linhaX.y1 + ((Form2.hfs * 0.4) * escalax)
            linhaif.y2 = linhaX.y1 + ((Form2.hfs * 0.4) * escalax)

        End If

            If Form1.posibaixa.Value = True Or Form1.posialta.Value = True Then
                linhasf.y1 = linhaX.y1 - ((Form2.hfs * 0.4) * escalax)
                linhasf.y2 = linhaX.y1 - ((Form2.hfs * 0.4) * escalax)

                linhaif.y1 = linhaX.y1 + ((Form2.hfs * 0.6) * escalax)
                linhaif.y2 = linhaX.y1 + ((Form2.hfs * 0.6) * escalax)
            End If

'Determinando a altura do final da fuselagem
linhafinal.Visible = True

linhafinal.x1 = linhaY.x1 + (Form2.lfs * escalax)
linhafinal.X2 = linhaY.x1 + (Form2.lfs * escalax)

linhafinal.y1 = linhaX.y1 - ((Form2.lfs * 0.025) * escalax)
linhafinal.y2 = linhaX.y1 + ((Form2.lfs * 0.025) * escalax)



'......... FIM localização das linhas ....................

'Dados de geometria das superficies transformadas
'Corda da asa
cordaw = Form7.CMaerodw * escalax

'Diametro do desenho do CA
diam = 0.025 * cordaw


'................. AEROFOLIO DA ASA - corda média

'........... Valor do eixo X = 11880
'.................... escalax = 10680 / (Form2.lfs * 100)

'Efeito Diedro - ed tangente de 1º = 0.017453
ed = ((Form10.diedro * 0.017453) * Form2.bw / 2) * escalax

If Form1.posialta.Value = True Then k = linhasf.y1 - (ed / 2)
If Form1.posimedia.Value = True Then k = Form18.valasa - (ed / 2)
If Form1.posibaixa.Value = True Then k = linhaif.y1 - (ed / 2)

distw = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))
distcam = distw + (0.25 * cordaw)

'Line (1200, k)-((distw), k), &HFF0000 'Linha azul LR ao BA

'Cordenadas em X
L1 = 0
l2 = 0.0125 * cordaw
l3 = 0.025 * cordaw
l4 = 0.05 * cordaw
l5 = 0.075 * cordaw
l6 = 0.1 * cordaw
l7 = 0.15 * cordaw
l8 = 0.2 * cordaw
l9 = 0.3 * cordaw
l10 = 0.4 * cordaw
l11 = 0.5 * cordaw
l12 = 0.6 * cordaw
l13 = 0.7 * cordaw
l14 = 0.8 * cordaw
l15 = 0.9 * cordaw
l16 = 0.95 * cordaw
l17 = 0.975 * cordaw
l18 = 1 * cordaw

'Cordenadas em Y

ya1 = 0
ya2 = 0.0185 * cordaw
ya3 = 0.0285 * cordaw
ya4 = 0.0385 * cordaw
ya5 = 0.048 * cordaw
ya6 = 0.0653 * cordaw
ya7 = 0.0723 * cordaw
ya8 = 0.078 * cordaw
ya9 = 0.0757 * cordaw
ya10 = 0.0675 * cordaw
ya11 = 0.055 * cordaw
ya12 = 0.0404 * cordaw
ya13 = 0.025 * cordaw
ya14 = 0.0118 * cordaw
ya15 = 0.006 * cordaw
ya16 = 0.003 * cordaw
ya17 = 0

yb1 = 0
yb2 = -0.015 * cordaw
yb3 = -0.02 * cordaw
yb4 = -0.025 * cordaw
yb5 = -0.03 * cordaw
yb6 = -0.033 * cordaw
yb7 = -0.0366 * cordaw
yb8 = -0.039 * cordaw
yb9 = -0.04 * cordaw
yb10 = -0.04 * cordaw
yb11 = -0.038 * cordaw
yb12 = -0.0346 * cordaw
yb13 = -0.0289 * cordaw
yb14 = -0.0215 * cordaw
yb15 = -0.0125 * cordaw
yb16 = -0.007 * cordaw
yb17 = 0

Form18.DrawWidth = 1
'Desenho do CA
Circle ((0.25 * cordaw) + distw + (diam / 2), k - ((0.25 * ya8))), diam, &HFF&

Form18.DrawWidth = 1
'Linhas superiores
'Form1.Line (l1 + dist, k)-(l2 + dist, k - ya1), &H0&
Form18.Line (l2 + distw, k - ya1)-(l3 + distw, k - ya2), &HC0C0C0
Form18.Line (l3 + distw, k - ya2)-(l4 + distw, k - ya3), &HC0C0C0
Form18.Line (l4 + distw, k - ya3)-(l5 + distw, k - ya4), &HC0C0C0
Form18.Line (l5 + distw, k - ya4)-(l6 + distw, k - ya5), &HC0C0C0
Form18.Line (l6 + distw, k - ya5)-(l7 + distw, k - ya6), &HC0C0C0
Form18.Line (l7 + distw, k - ya6)-(l8 + distw, k - ya7), &HC0C0C0
Form18.Line (l8 + distw, k - ya7)-(l9 + distw, k - ya8), &HC0C0C0
Form18.Line (l9 + distw, k - ya8)-(l10 + distw, k - ya9), &HC0C0C0
Form18.Line (l10 + distw, k - ya9)-(l11 + distw, k - ya10), &HC0C0C0
Form18.Line (l11 + distw, k - ya10)-(l12 + distw, k - ya11), &HC0C0C0
Form18.Line (l12 + distw, k - ya11)-(l13 + distw, k - ya12), &HC0C0C0
Form18.Line (l13 + distw, k - ya12)-(l14 + distw, k - ya13), &HC0C0C0
Form18.Line (l14 + distw, k - ya13)-(l15 + distw, k - ya14), &HC0C0C0
Form18.Line (l15 + distw, k - ya14)-(l16 + distw, k - ya15), &HC0C0C0
Form18.Line (l16 + distw, k - ya15)-(l17 + distw, k - ya16), &HC0C0C0
Form18.Line (l17 + distw, k - ya16)-(l18 + distw, k - ya17), &HC0C0C0
Form18.Line (l18 + distw, k - ya17)-(l18 + distw, k), &HC0C0C0

'Linhas inferiores
'Form1.Line (l1 + dist, k)-(l2 + dist, k - yb1), &H0&
Form18.Line (l2 + distw, k - yb1)-(l3 + distw, k - yb2), &HC0C0C0
Form18.Line (l3 + distw, k - yb2)-(l4 + distw, k - yb3), &HC0C0C0
Form18.Line (l4 + distw, k - yb3)-(l5 + distw, k - yb4), &HC0C0C0
Form18.Line (l5 + distw, k - yb4)-(l6 + distw, k - yb5), &HC0C0C0
Form18.Line (l6 + distw, k - yb5)-(l7 + distw, k - yb6), &HC0C0C0
Form18.Line (l7 + distw, k - yb6)-(l8 + distw, k - yb7), &HC0C0C0
Form18.Line (l8 + distw, k - yb7)-(l9 + distw, k - yb8), &HC0C0C0
Form18.Line (l9 + distw, k - yb8)-(l10 + distw, k - yb9), &HC0C0C0
Form18.Line (l10 + distw, k - yb9)-(l11 + distw, k - yb10), &HC0C0C0
Form18.Line (l11 + distw, k - yb10)-(l12 + distw, k - yb11), &HC0C0C0
Form18.Line (l12 + distw, k - yb11)-(l13 + distw, k - yb12), &HC0C0C0
Form18.Line (l13 + distw, k - yb12)-(l14 + distw, k - yb13), &HC0C0C0
Form18.Line (l14 + distw, k - yb13)-(l15 + distw, k - yb14), &HC0C0C0
Form18.Line (l15 + distw, k - yb14)-(l16 + distw, k - yb15), &HC0C0C0
Form18.Line (l16 + distw, k - yb15)-(l17 + distw, k - yb16), &HC0C0C0
Form18.Line (l17 + distw, k - yb16)-(l18 + distw, k - yb17), &HC0C0C0
Form18.Line (l18 + distw, k - yb17)-(l18 + distw, k), &HC0C0C0

' XXXXXXXXXX Corda da Ponta/15/08/2001 XXXXXXXXXXXXX
Dim distwP, distcaP, cordawP, locap As Double

cordawP = Form2.cpw * escalax

'Diametro do desenho do CA
diam = 0.025 * cordawP
'Efeito Diedro - ed tangente de 1º = 0.017453
ed = ((Form10.diedro * 0.017453) * Form2.bw / 2) * escalax
If Form1.posialta.Value = True Then k = linhasf.y1 - ed
If Form1.posimedia.Value = True Then k = Form18.valasa - ed
If Form1.posibaixa.Value = True Then k = linhaif.y1 - ed

'Fator de escala
fatx = Form18.esc18
escalax = 800 * fatx

'daqui
locap = ((Form2.bw / 2) * Tan(grau * 3.1415162 / 180)) * escalax
distwP = Abs(distw) + Abs(locap / 2)

'd1 = Tan(grau * 3.1415162 / 180)
'd2 = distw
'd3 = locap / 2
'd4 = distwP
'aqui

distcaP = distwP + (0.25 * cordawP)

'Line (1200, k)-((distw), k), &HFF0000 'Linha azul LR ao BA

'Cordenadas em X
L1 = 0
l2 = 0.0125 * cordawP
l3 = 0.025 * cordawP
l4 = 0.05 * cordawP
l5 = 0.075 * cordawP
l6 = 0.1 * cordawP
l7 = 0.15 * cordawP
l8 = 0.2 * cordawP
l9 = 0.3 * cordawP
l10 = 0.4 * cordawP
l11 = 0.5 * cordawP
l12 = 0.6 * cordawP
l13 = 0.7 * cordawP
l14 = 0.8 * cordawP
l15 = 0.9 * cordawP
l16 = 0.95 * cordawP
l17 = 0.975 * cordawP
l18 = 1 * cordawP

'Cordenadas em Y

ya1 = 0
ya2 = 0.0185 * cordawP
ya3 = 0.0285 * cordawP
ya4 = 0.0385 * cordawP
ya5 = 0.048 * cordawP
ya6 = 0.0653 * cordawP
ya7 = 0.0723 * cordawP
ya8 = 0.078 * cordawP
ya9 = 0.0757 * cordawP
ya10 = 0.0675 * cordawP
ya11 = 0.055 * cordawP
ya12 = 0.0404 * cordawP
ya13 = 0.025 * cordawP
ya14 = 0.0118 * cordawP
ya15 = 0.006 * cordawP
ya16 = 0.003 * cordawP
ya17 = 0

yb1 = 0
yb2 = -0.015 * cordawP
yb3 = -0.02 * cordawP
yb4 = -0.025 * cordawP
yb5 = -0.03 * cordawP
yb6 = -0.033 * cordawP
yb7 = -0.0366 * cordawP
yb8 = -0.039 * cordawP
yb9 = -0.04 * cordawP
yb10 = -0.04 * cordawP
yb11 = -0.038 * cordawP
yb12 = -0.0346 * cordawP
yb13 = -0.0289 * cordawP
yb14 = -0.0215 * cordawP
yb15 = -0.0125 * cordawP
yb16 = -0.007 * cordawP
yb17 = 0

Form18.DrawWidth = 1
'Desenho do CA
Circle ((0.25 * cordawP) + distwP + (diam / 2), k - ((0.25 * ya8))), diam, &HFF&

'Linhas superiores
'Form1.Line (l1 + dist, k)-(l2 + dist, k - ya1), &H0&
Form18.Line (l2 + distwP, k - ya1)-(l3 + distwP, k - ya2), &H0&
Form18.Line (l3 + distwP, k - ya2)-(l4 + distwP, k - ya3), &H0&
Form18.Line (l4 + distwP, k - ya3)-(l5 + distwP, k - ya4), &H0&
Form18.Line (l5 + distwP, k - ya4)-(l6 + distwP, k - ya5), &H0&
Form18.Line (l6 + distwP, k - ya5)-(l7 + distwP, k - ya6), &H0&
Form18.Line (l7 + distwP, k - ya6)-(l8 + distwP, k - ya7), &H0&
Form18.Line (l8 + distwP, k - ya7)-(l9 + distwP, k - ya8), &H0&
Form18.Line (l9 + distwP, k - ya8)-(l10 + distwP, k - ya9), &H0&
Form18.Line (l10 + distwP, k - ya9)-(l11 + distwP, k - ya10), &H0&
Form18.Line (l11 + distwP, k - ya10)-(l12 + distwP, k - ya11), &H0&
Form18.Line (l12 + distwP, k - ya11)-(l13 + distwP, k - ya12), &H0&
Form18.Line (l13 + distwP, k - ya12)-(l14 + distwP, k - ya13), &H0&
Form18.Line (l14 + distwP, k - ya13)-(l15 + distwP, k - ya14), &H0&
Form18.Line (l15 + distwP, k - ya14)-(l16 + distwP, k - ya15), &H0&
Form18.Line (l16 + distwP, k - ya15)-(l17 + distwP, k - ya16), &H0&
Form18.Line (l17 + distwP, k - ya16)-(l18 + distwP, k - ya17), &H0&
Form18.Line (l18 + distwP, k - ya17)-(l18 + distwP, k), &H0&

'Linhas inferiores
'Form1.Line (l1 + dist, k)-(l2 + dist, k - yb1), &H0&
Form18.Line (l2 + distwP, k - yb1)-(l3 + distwP, k - yb2), &H0&
Form18.Line (l3 + distwP, k - yb2)-(l4 + distwP, k - yb3), &H0&
Form18.Line (l4 + distwP, k - yb3)-(l5 + distwP, k - yb4), &H0&
Form18.Line (l5 + distwP, k - yb4)-(l6 + distwP, k - yb5), &H0&
Form18.Line (l6 + distwP, k - yb5)-(l7 + distwP, k - yb6), &H0&
Form18.Line (l7 + distwP, k - yb6)-(l8 + distwP, k - yb7), &H0&
Form18.Line (l8 + distwP, k - yb7)-(l9 + distwP, k - yb8), &H0&
Form18.Line (l9 + distwP, k - yb8)-(l10 + distwP, k - yb9), &H0&
Form18.Line (l10 + distwP, k - yb9)-(l11 + distwP, k - yb10), &H0&
Form18.Line (l11 + distwP, k - yb10)-(l12 + distwP, k - yb11), &H0&
Form18.Line (l12 + distwP, k - yb11)-(l13 + distwP, k - yb12), &H0&
Form18.Line (l13 + distwP, k - yb12)-(l14 + distwP, k - yb13), &H0&
Form18.Line (l14 + distwP, k - yb13)-(l15 + distwP, k - yb14), &H0&
Form18.Line (l15 + distwP, k - yb14)-(l16 + distwP, k - yb15), &H0&
Form18.Line (l16 + distwP, k - yb15)-(l17 + distwP, k - yb16), &H0&
Form18.Line (l17 + distwP, k - yb16)-(l18 + distwP, k - yb17), &H0&
Form18.Line (l18 + distwP, k - yb17)-(l18 + distwP, k), &H0&

' XXXXXXXXXX Corda da Raiz /15/08/2001 XXXXXXXXXXXXX

Dim distwR, distcaR, cordawR, locaR As Double

cordawR = Form2.crw * escalax

'Diametro do desenho do CA
diam = 0.025 * cordawR

If Form1.posialta.Value = True Then k = linhasf.y1
If Form1.posimedia.Value = True Then k = Form18.valasa
If Form1.posibaixa.Value = True Then k = linhaif.y1

'distwR = linhaY.x1 + (Form10.LRCAasa * escalax - (0.25 * cordaw))

'daqui
locaR = ((Form2.bw / 2) * Tan(grau * 3.1415162 / 180)) * escalax
distwR = Abs(distw) - Abs(locaR / 2)

'd1 = Tan(grau * 3.1415162 / 180)
'd2 = distw
'd3 = locaR / 2
'd4 = distwR
'aqui

distcaR = distwR + (0.25 * cordawR)

'Line (1200, k)-((distw), k), &HFF0000 'Linha azul LR ao BA

'Cordenadas em X
L1 = 0
l2 = 0.0125 * cordawR
l3 = 0.025 * cordawR
l4 = 0.05 * cordawR
l5 = 0.075 * cordawR
l6 = 0.1 * cordawR
l7 = 0.15 * cordawR
l8 = 0.2 * cordawR
l9 = 0.3 * cordawR
l10 = 0.4 * cordawR
l11 = 0.5 * cordawR
l12 = 0.6 * cordawR
l13 = 0.7 * cordawR
l14 = 0.8 * cordawR
l15 = 0.9 * cordawR
l16 = 0.95 * cordawR
l17 = 0.975 * cordawR
l18 = 1 * cordawR

'Cordenadas em Y

ya1 = 0
ya2 = 0.0185 * cordawR
ya3 = 0.0285 * cordawR
ya4 = 0.0385 * cordawR
ya5 = 0.048 * cordawR
ya6 = 0.0653 * cordawR
ya7 = 0.0723 * cordawR
ya8 = 0.078 * cordawR
ya9 = 0.0757 * cordawR
ya10 = 0.0675 * cordawR
ya11 = 0.055 * cordawR
ya12 = 0.0404 * cordawR
ya13 = 0.025 * cordawR
ya14 = 0.0118 * cordawR
ya15 = 0.006 * cordawR
ya16 = 0.003 * cordawR
ya17 = 0

yb1 = 0
yb2 = -0.015 * cordawR
yb3 = -0.02 * cordawR
yb4 = -0.025 * cordawR
yb5 = -0.03 * cordawR
yb6 = -0.033 * cordawR
yb7 = -0.0366 * cordawR
yb8 = -0.039 * cordawR
yb9 = -0.04 * cordawR
yb10 = -0.04 * cordawR
yb11 = -0.038 * cordawR
yb12 = -0.0346 * cordawR
yb13 = -0.0289 * cordawR
yb14 = -0.0215 * cordawR
yb15 = -0.0125 * cordawR
yb16 = -0.007 * cordawR
yb17 = 0

Form18.DrawWidth = 1
'Desenho do CA
Circle ((0.25 * cordawR) + distwR + (diam / 2), k - ((0.25 * ya8))), diam, &HFF&

'Linhas superiores
'Form1.Line (l1 + dist, k)-(l2 + dist, k - ya1), &H0&
Form18.Line (l2 + distwR, k - ya1)-(l3 + distwR, k - ya2), &H0&
Form18.Line (l3 + distwR, k - ya2)-(l4 + distwR, k - ya3), &H0&
Form18.Line (l4 + distwR, k - ya3)-(l5 + distwR, k - ya4), &H0&
Form18.Line (l5 + distwR, k - ya4)-(l6 + distwR, k - ya5), &H0&
Form18.Line (l6 + distwR, k - ya5)-(l7 + distwR, k - ya6), &H0&
Form18.Line (l7 + distwR, k - ya6)-(l8 + distwR, k - ya7), &H0&
Form18.Line (l8 + distwR, k - ya7)-(l9 + distwR, k - ya8), &H0&
Form18.Line (l9 + distwR, k - ya8)-(l10 + distwR, k - ya9), &H0&
Form18.Line (l10 + distwR, k - ya9)-(l11 + distwR, k - ya10), &H0&
Form18.Line (l11 + distwR, k - ya10)-(l12 + distwR, k - ya11), &H0&
Form18.Line (l12 + distwR, k - ya11)-(l13 + distwR, k - ya12), &H0&
Form18.Line (l13 + distwR, k - ya12)-(l14 + distwR, k - ya13), &H0&
Form18.Line (l14 + distwR, k - ya13)-(l15 + distwR, k - ya14), &H0&
Form18.Line (l15 + distwR, k - ya14)-(l16 + distwR, k - ya15), &H0&
Form18.Line (l16 + distwR, k - ya15)-(l17 + distwR, k - ya16), &H0&
Form18.Line (l17 + distwR, k - ya16)-(l18 + distwR, k - ya17), &H0&
Form18.Line (l18 + distwR, k - ya17)-(l18 + distwR, k), &H0&

'Linhas inferiores
'Form1.Line (l1 + dist, k)-(l2 + dist, k - yb1), &H0&
Form18.Line (l2 + distwR, k - yb1)-(l3 + distwR, k - yb2), &H0&
Form18.Line (l3 + distwR, k - yb2)-(l4 + distwR, k - yb3), &H0&
Form18.Line (l4 + distwR, k - yb3)-(l5 + distwR, k - yb4), &H0&
Form18.Line (l5 + distwR, k - yb4)-(l6 + distwR, k - yb5), &H0&
Form18.Line (l6 + distwR, k - yb5)-(l7 + distwR, k - yb6), &H0&
Form18.Line (l7 + distwR, k - yb6)-(l8 + distwR, k - yb7), &H0&
Form18.Line (l8 + distwR, k - yb7)-(l9 + distwR, k - yb8), &H0&
Form18.Line (l9 + distwR, k - yb8)-(l10 + distwR, k - yb9), &H0&
Form18.Line (l10 + distwR, k - yb9)-(l11 + distwR, k - yb10), &H0&
Form18.Line (l11 + distwR, k - yb10)-(l12 + distwR, k - yb11), &H0&
Form18.Line (l12 + distwR, k - yb11)-(l13 + distwR, k - yb12), &H0&
Form18.Line (l13 + distwR, k - yb12)-(l14 + distwR, k - yb13), &H0&
Form18.Line (l14 + distwR, k - yb13)-(l15 + distwR, k - yb14), &H0&
Form18.Line (l15 + distwR, k - yb14)-(l16 + distwR, k - yb15), &H0&
Form18.Line (l16 + distwR, k - yb15)-(l17 + distwR, k - yb16), &H0&
Form18.Line (l17 + distwR, k - yb16)-(l18 + distwR, k - yb17), &H0&
Form18.Line (l18 + distwR, k - yb17)-(l18 + distwR, k), &H0&

'..................AEROFOLIO DA EMPENAGEM

escalax = 800 * fatx


'Dados de geometria das superficies transformadas
'Corda da asa
cordaht = Form7.CMaerodht * escalax

dist = distcam + ((Form10.LT * escalax) - (0.25 * cordaht))

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)
'......... Fim .........
'Cordenadas em X
L1 = 0
l2 = 0.0125 * cordaht
l3 = 0.025 * cordaht
l4 = 0.05 * cordaht
l5 = 0.075 * cordaht
l6 = 0.1 * cordaht
l7 = 0.15 * cordaht
l8 = 0.2 * cordaht
l9 = 0.3 * cordaht
l10 = 0.4 * cordaht
l11 = 0.5 * cordaht
l12 = 0.6 * cordaht
l13 = 0.7 * cordaht
l14 = 0.8 * cordaht
l15 = 0.9 * cordaht
l16 = 0.95 * cordaht
l17 = 0.975 * cordaht
l18 = 1 * cordaht

'Cordenadas em Y

ya1 = 0
ya2 = 0.022 * cordaht
ya3 = 0.03 * cordaht
ya4 = 0.035 * cordaht
ya5 = 0.04 * cordaht
ya6 = 0.045 * cordaht
ya7 = 0.048 * cordaht
ya8 = 0.05 * cordaht
ya9 = 0.051 * cordaht
ya10 = 0.0488 * cordaht
ya11 = 0.044 * cordaht
ya12 = 0.0384 * cordaht
ya13 = 0.031 * cordaht
ya14 = 0.0221 * cordaht
ya15 = 0.0122 * cordaht
ya16 = 0.01 * cordaht
ya17 = 0

yb1 = 0
yb2 = -0.022 * cordaht
yb3 = -0.03 * cordaht
yb4 = -0.035 * cordaht
yb5 = -0.04 * cordaht
yb6 = -0.045 * cordaht
yb7 = -0.048 * cordaht
yb8 = -0.05 * cordaht
yb9 = -0.051 * cordaht
yb10 = -0.0488 * cordaht
yb11 = -0.044 * cordaht
yb12 = -0.0384 * cordaht
yb13 = -0.031 * cordaht
yb14 = -0.0221 * cordaht
yb15 = -0.0122 * cordaht
yb16 = -0.01 * cordaht
yb17 = 0

Form18.DrawWidth = 1
Circle (l9 + dist + (diam / 2), ktail - ((0.25 * ya8))), 0.5 * diam, &HFF&

Form18.DrawWidth = 1

'Linhas superiores
Form18.Line (L1 + dist, ktail)-(l2 + dist, ktail - ya1), &HC0C0C0
Form18.Line (l2 + dist, ktail - ya1)-(l3 + dist, ktail - ya2), &HC0C0C0
Form18.Line (l3 + dist, ktail - ya2)-(l4 + dist, ktail - ya3), &HC0C0C0
Form18.Line (l4 + dist, ktail - ya3)-(l5 + dist, ktail - ya4), &HC0C0C0
Form18.Line (l5 + dist, ktail - ya4)-(l6 + dist, ktail - ya5), &HC0C0C0
Form18.Line (l6 + dist, ktail - ya5)-(l7 + dist, ktail - ya6), &HC0C0C0
Form18.Line (l7 + dist, ktail - ya6)-(l8 + dist, ktail - ya7), &HC0C0C0
Form18.Line (l8 + dist, ktail - ya7)-(l9 + dist, ktail - ya8), &HC0C0C0
Form18.Line (l9 + dist, ktail - ya8)-(l10 + dist, ktail - ya9), &HC0C0C0
Form18.Line (l10 + dist, ktail - ya9)-(l11 + dist, ktail - ya10), &HC0C0C0
Form18.Line (l11 + dist, ktail - ya10)-(l12 + dist, ktail - ya11), &HC0C0C0
Form18.Line (l12 + dist, ktail - ya11)-(l13 + dist, ktail - ya12), &HC0C0C0
Form18.Line (l13 + dist, ktail - ya12)-(l14 + dist, ktail - ya13), &HC0C0C0
Form18.Line (l14 + dist, ktail - ya13)-(l15 + dist, ktail - ya14), &HC0C0C0
Form18.Line (l15 + dist, ktail - ya14)-(l16 + dist, ktail - ya15), &HC0C0C0
Form18.Line (l16 + dist, ktail - ya15)-(l17 + dist, ktail - ya16), &HC0C0C0
Form18.Line (l17 + dist, ktail - ya16)-(l18 + dist, ktail - ya17), &HC0C0C0
Form18.Line (l18 + dist, ktail - ya17)-(l18 + dist, ktail), &HC0C0C0

'Linhas inferiores
Form18.Line (L1 + dist, ktail)-(l2 + dist, ktail - yb1), &HC0C0C0
Form18.Line (l2 + dist, ktail - yb1)-(l3 + dist, ktail - yb2), &HC0C0C0
Form18.Line (l3 + dist, ktail - yb2)-(l4 + dist, ktail - yb3), &HC0C0C0
Form18.Line (l4 + dist, ktail - yb3)-(l5 + dist, ktail - yb4), &HC0C0C0
Form18.Line (l5 + dist, ktail - yb4)-(l6 + dist, ktail - yb5), &HC0C0C0
Form18.Line (l6 + dist, ktail - yb5)-(l7 + dist, ktail - yb6), &HC0C0C0
Form18.Line (l7 + dist, ktail - yb6)-(l8 + dist, ktail - yb7), &HC0C0C0
Form18.Line (l8 + dist, ktail - yb7)-(l9 + dist, ktail - yb8), &HC0C0C0
Form18.Line (l9 + dist, ktail - yb8)-(l10 + dist, ktail - yb9), &HC0C0C0
Form18.Line (l10 + dist, ktail - yb9)-(l11 + dist, ktail - yb10), &HC0C0C0
Form18.Line (l11 + dist, ktail - yb10)-(l12 + dist, ktail - yb11), &HC0C0C0
Form18.Line (l12 + dist, ktail - yb11)-(l13 + dist, ktail - yb12), &HC0C0C0
Form18.Line (l13 + dist, ktail - yb12)-(l14 + dist, ktail - yb13), &HC0C0C0
Form18.Line (l14 + dist, ktail - yb13)-(l15 + dist, ktail - yb14), &HC0C0C0
Form18.Line (l15 + dist, ktail - yb14)-(l16 + dist, ktail - yb15), &HC0C0C0
Form18.Line (l16 + dist, ktail - yb15)-(l17 + dist, ktail - yb16), &HC0C0C0
Form18.Line (l17 + dist, ktail - yb16)-(l18 + dist, ktail - yb17), &HC0C0C0
Form18.Line (l18 + dist, ktail - yb17)-(l18 + dist, ktail), &HC0C0C0

'............... AEROFOLIO DA PONTA DA EMPENAGEM ................

'Dados de geometria das superficies transformadas
'Corda da asa
Dim cordahtP, locahtp, disthtP As Double
cordahtP = Form2.cpht * escalax

locahtp = ((Form2.bht / 2) * Tan(grauht * 3.1415162 / 180)) * escalax
disthtP = Abs(dist) + Abs(locahtp / 2)

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If
'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)
'......... Fim .........
'Cordenadas em X
L1 = 0
l2 = 0.0125 * cordahtP
l3 = 0.025 * cordahtP
l4 = 0.05 * cordahtP
l5 = 0.075 * cordahtP
l6 = 0.1 * cordahtP
l7 = 0.15 * cordahtP
l8 = 0.2 * cordahtP
l9 = 0.3 * cordahtP
l10 = 0.4 * cordahtP
l11 = 0.5 * cordahtP
l12 = 0.6 * cordahtP
l13 = 0.7 * cordahtP
l14 = 0.8 * cordahtP
l15 = 0.9 * cordahtP
l16 = 0.95 * cordahtP
l17 = 0.975 * cordahtP
l18 = 1 * cordahtP

'Cordenadas em Y

ya1 = 0
ya2 = 0.022 * cordahtP
ya3 = 0.03 * cordahtP
ya4 = 0.035 * cordahtP
ya5 = 0.04 * cordahtP
ya6 = 0.045 * cordahtP
ya7 = 0.048 * cordahtP
ya8 = 0.05 * cordahtP
ya9 = 0.051 * cordahtP
ya10 = 0.0488 * cordahtP
ya11 = 0.044 * cordahtP
ya12 = 0.0384 * cordahtP
ya13 = 0.031 * cordahtP
ya14 = 0.0221 * cordahtP
ya15 = 0.0122 * cordahtP
ya16 = 0.01 * cordahtP
ya17 = 0

yb1 = 0
yb2 = -0.022 * cordahtP
yb3 = -0.03 * cordahtP
yb4 = -0.035 * cordahtP
yb5 = -0.04 * cordahtP
yb6 = -0.045 * cordahtP
yb7 = -0.048 * cordahtP
yb8 = -0.05 * cordahtP
yb9 = -0.051 * cordahtP
yb10 = -0.0488 * cordahtP
yb11 = -0.044 * cordahtP
yb12 = -0.0384 * cordahtP
yb13 = -0.031 * cordahtP
yb14 = -0.0221 * cordahtP
yb15 = -0.0122 * cordahtP
yb16 = -0.01 * cordahtP
yb17 = 0

Form18.DrawWidth = 1
Circle (l9 + disthtP + (diam / 2), ktail - ((0.25 * ya8))), 0.5 * diam, &HFF&

Form18.DrawWidth = 1

'Linhas superiores
Form18.Line (L1 + disthtP, ktail)-(l2 + disthtP, ktail - ya1), &H0&
Form18.Line (l2 + disthtP, ktail - ya1)-(l3 + disthtP, ktail - ya2), &H0&
Form18.Line (l3 + disthtP, ktail - ya2)-(l4 + disthtP, ktail - ya3), &H0&
Form18.Line (l4 + disthtP, ktail - ya3)-(l5 + disthtP, ktail - ya4), &H0&
Form18.Line (l5 + disthtP, ktail - ya4)-(l6 + disthtP, ktail - ya5), &H0&
Form18.Line (l6 + disthtP, ktail - ya5)-(l7 + disthtP, ktail - ya6), &H0&
Form18.Line (l7 + disthtP, ktail - ya6)-(l8 + disthtP, ktail - ya7), &H0&
Form18.Line (l8 + disthtP, ktail - ya7)-(l9 + disthtP, ktail - ya8), &H0&
Form18.Line (l9 + disthtP, ktail - ya8)-(l10 + disthtP, ktail - ya9), &H0&
Form18.Line (l10 + disthtP, ktail - ya9)-(l11 + disthtP, ktail - ya10), &H0&
Form18.Line (l11 + disthtP, ktail - ya10)-(l12 + disthtP, ktail - ya11), &H0&
Form18.Line (l12 + disthtP, ktail - ya11)-(l13 + disthtP, ktail - ya12), &H0&
Form18.Line (l13 + disthtP, ktail - ya12)-(l14 + disthtP, ktail - ya13), &H0&
Form18.Line (l14 + disthtP, ktail - ya13)-(l15 + disthtP, ktail - ya14), &H0&
Form18.Line (l15 + disthtP, ktail - ya14)-(l16 + disthtP, ktail - ya15), &H0&
Form18.Line (l16 + disthtP, ktail - ya15)-(l17 + disthtP, ktail - ya16), &H0&
Form18.Line (l17 + disthtP, ktail - ya16)-(l18 + disthtP, ktail - ya17), &H0&
Form18.Line (l18 + disthtP, ktail - ya17)-(l18 + disthtP, ktail), &H0&

'Linhas inferiores
Form18.Line (L1 + disthtP, ktail)-(l2 + disthtP, ktail - yb1), &H0&
Form18.Line (l2 + disthtP, ktail - yb1)-(l3 + disthtP, ktail - yb2), &H0&
Form18.Line (l3 + disthtP, ktail - yb2)-(l4 + disthtP, ktail - yb3), &H0&
Form18.Line (l4 + disthtP, ktail - yb3)-(l5 + disthtP, ktail - yb4), &H0&
Form18.Line (l5 + disthtP, ktail - yb4)-(l6 + disthtP, ktail - yb5), &H0&
Form18.Line (l6 + disthtP, ktail - yb5)-(l7 + disthtP, ktail - yb6), &H0&
Form18.Line (l7 + disthtP, ktail - yb6)-(l8 + disthtP, ktail - yb7), &H0&
Form18.Line (l8 + disthtP, ktail - yb7)-(l9 + disthtP, ktail - yb8), &H0&
Form18.Line (l9 + disthtP, ktail - yb8)-(l10 + disthtP, ktail - yb9), &H0&
Form18.Line (l10 + disthtP, ktail - yb9)-(l11 + disthtP, ktail - yb10), &H0&
Form18.Line (l11 + disthtP, ktail - yb10)-(l12 + disthtP, ktail - yb11), &H0&
Form18.Line (l12 + disthtP, ktail - yb11)-(l13 + disthtP, ktail - yb12), &H0&
Form18.Line (l13 + disthtP, ktail - yb12)-(l14 + disthtP, ktail - yb13), &H0&
Form18.Line (l14 + disthtP, ktail - yb13)-(l15 + disthtP, ktail - yb14), &H0&
Form18.Line (l15 + disthtP, ktail - yb14)-(l16 + disthtP, ktail - yb15), &H0&
Form18.Line (l16 + disthtP, ktail - yb15)-(l17 + disthtP, ktail - yb16), &H0&
Form18.Line (l17 + disthtP, ktail - yb16)-(l18 + disthtP, ktail - yb17), &H0&
Form18.Line (l18 + disthtP, ktail - yb17)-(l18 + disthtP, ktail), &H0&

'............... AEROFOLIO DA RAIZ DA EMPENAGEM ................

'Dados de geometria das superficies transformadas
'Corda da asa
Dim cordahtR, locahtr, disthtR As Double
cordahtR = Form2.crht * escalax

locahtr = ((Form2.bht / 2) * Tan(grauht * 3.1415162 / 180)) * escalax
disthtR = Abs(dist) - Abs(locahtr / 2)

If Form10.htail <> 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhaX.y1
If Form1.posibaixa Then ktail = linhaX.y1
End If

If Form10.htail = 0 Then
If Form1.posimedia Then ktail = linhaX.y1
If Form1.posialta Then ktail = linhasf.y1
If Form1.posibaixa Then ktail = linhaif.y1
End If

'......... Localização de htail na vertical..........
ktail = ktail - (Form10.htail * escalax)
'......... Fim .........
'Cordenadas em X
L1 = 0
l2 = 0.0125 * cordahtR
l3 = 0.025 * cordahtR
l4 = 0.05 * cordahtR
l5 = 0.075 * cordahtR
l6 = 0.1 * cordahtR
l7 = 0.15 * cordahtR
l8 = 0.2 * cordahtR
l9 = 0.3 * cordahtR
l10 = 0.4 * cordahtR
l11 = 0.5 * cordahtR
l12 = 0.6 * cordahtR
l13 = 0.7 * cordahtR
l14 = 0.8 * cordahtR
l15 = 0.9 * cordahtR
l16 = 0.95 * cordahtR
l17 = 0.975 * cordahtR
l18 = 1 * cordahtR

'Cordenadas em Y

ya1 = 0
ya2 = 0.022 * cordahtR
ya3 = 0.03 * cordahtR
ya4 = 0.035 * cordahtR
ya5 = 0.04 * cordahtR
ya6 = 0.045 * cordahtR
ya7 = 0.048 * cordahtR
ya8 = 0.05 * cordahtR
ya9 = 0.051 * cordahtR
ya10 = 0.0488 * cordahtR
ya11 = 0.044 * cordahtR
ya12 = 0.0384 * cordahtR
ya13 = 0.031 * cordahtR
ya14 = 0.0221 * cordahtR
ya15 = 0.0122 * cordahtR
ya16 = 0.01 * cordahtR
ya17 = 0

yb1 = 0
yb2 = -0.022 * cordahtR
yb3 = -0.03 * cordahtR
yb4 = -0.035 * cordahtR
yb5 = -0.04 * cordahtR
yb6 = -0.045 * cordahtR
yb7 = -0.048 * cordahtR
yb8 = -0.05 * cordahtR
yb9 = -0.051 * cordahtR
yb10 = -0.0488 * cordahtR
yb11 = -0.044 * cordahtR
yb12 = -0.0384 * cordahtR
yb13 = -0.031 * cordahtR
yb14 = -0.0221 * cordahtR
yb15 = -0.0122 * cordahtR
yb16 = -0.01 * cordahtR
yb17 = 0

Form18.DrawWidth = 1
Circle (l9 + disthtR + (diam / 2), ktail - ((0.25 * ya8))), 0.5 * diam, &HFF&

Form18.DrawWidth = 1

'Linhas superiores
Form18.Line (L1 + disthtR, ktail)-(l2 + disthtR, ktail - ya1), &H0&
Form18.Line (l2 + disthtR, ktail - ya1)-(l3 + disthtR, ktail - ya2), &H0&
Form18.Line (l3 + disthtR, ktail - ya2)-(l4 + disthtR, ktail - ya3), &H0&
Form18.Line (l4 + disthtR, ktail - ya3)-(l5 + disthtR, ktail - ya4), &H0&
Form18.Line (l5 + disthtR, ktail - ya4)-(l6 + disthtR, ktail - ya5), &H0&
Form18.Line (l6 + disthtR, ktail - ya5)-(l7 + disthtR, ktail - ya6), &H0&
Form18.Line (l7 + disthtR, ktail - ya6)-(l8 + disthtR, ktail - ya7), &H0&
Form18.Line (l8 + disthtR, ktail - ya7)-(l9 + disthtR, ktail - ya8), &H0&
Form18.Line (l9 + disthtR, ktail - ya8)-(l10 + disthtR, ktail - ya9), &H0&
Form18.Line (l10 + disthtR, ktail - ya9)-(l11 + disthtR, ktail - ya10), &H0&
Form18.Line (l11 + disthtR, ktail - ya10)-(l12 + disthtR, ktail - ya11), &H0&
Form18.Line (l12 + disthtR, ktail - ya11)-(l13 + disthtR, ktail - ya12), &H0&
Form18.Line (l13 + disthtR, ktail - ya12)-(l14 + disthtR, ktail - ya13), &H0&
Form18.Line (l14 + disthtR, ktail - ya13)-(l15 + disthtR, ktail - ya14), &H0&
Form18.Line (l15 + disthtR, ktail - ya14)-(l16 + disthtR, ktail - ya15), &H0&
Form18.Line (l16 + disthtR, ktail - ya15)-(l17 + disthtR, ktail - ya16), &H0&
Form18.Line (l17 + disthtR, ktail - ya16)-(l18 + disthtR, ktail - ya17), &H0&
Form18.Line (l18 + disthtR, ktail - ya17)-(l18 + disthtR, ktail), &H0&

'Linhas inferiores
Form18.Line (L1 + disthtR, ktail)-(l2 + disthtR, ktail - yb1), &H0&
Form18.Line (l2 + disthtR, ktail - yb1)-(l3 + disthtR, ktail - yb2), &H0&
Form18.Line (l3 + disthtR, ktail - yb2)-(l4 + disthtR, ktail - yb3), &H0&
Form18.Line (l4 + disthtR, ktail - yb3)-(l5 + disthtR, ktail - yb4), &H0&
Form18.Line (l5 + disthtR, ktail - yb4)-(l6 + disthtR, ktail - yb5), &H0&
Form18.Line (l6 + disthtR, ktail - yb5)-(l7 + disthtR, ktail - yb6), &H0&
Form18.Line (l7 + disthtR, ktail - yb6)-(l8 + disthtR, ktail - yb7), &H0&
Form18.Line (l8 + disthtR, ktail - yb7)-(l9 + disthtR, ktail - yb8), &H0&
Form18.Line (l9 + disthtR, ktail - yb8)-(l10 + disthtR, ktail - yb9), &H0&
Form18.Line (l10 + disthtR, ktail - yb9)-(l11 + disthtR, ktail - yb10), &H0&
Form18.Line (l11 + disthtR, ktail - yb10)-(l12 + disthtR, ktail - yb11), &H0&
Form18.Line (l12 + disthtR, ktail - yb11)-(l13 + disthtR, ktail - yb12), &H0&
Form18.Line (l13 + disthtR, ktail - yb12)-(l14 + disthtR, ktail - yb13), &H0&
Form18.Line (l14 + disthtR, ktail - yb13)-(l15 + disthtR, ktail - yb14), &H0&
Form18.Line (l15 + disthtR, ktail - yb14)-(l16 + disthtR, ktail - yb15), &H0&
Form18.Line (l16 + disthtR, ktail - yb15)-(l17 + disthtR, ktail - yb16), &H0&
Form18.Line (l17 + disthtR, ktail - yb16)-(l18 + disthtR, ktail - yb17), &H0&
Form18.Line (l18 + disthtR, ktail - yb17)-(l18 + disthtR, ktail), &H0&


'INSERINDO O CG NO DESENHO
Circle ((1680 + (Form14.cg / 100) * escalax), linhaX.y1), 0.1 * escalax, &H800000
Line ((1680 + (Form14.cg / 100) * escalax), linhaX.y1 - (0.2 * escalax))-((1680 + (Form14.cg / 100) * escalax), linhaX.y1 + (0.2 * escalax)), &H800000
Line ((1680 + (Form14.cg / 100) * escalax) - (0.2 * escalax), linhaX.y1)-((1680 + (Form14.cg / 100) * escalax) + (0.2 * escalax), linhaX.y1), &H800000


End Sub

Private Sub insereqip_Click()
Frame22.Visible = False
Frame1.Visible = True
Label3.Top = 1440
med2.Top = 1680
End Sub

Private Sub inseridesenho_Click()
Frame11.Visible = True
Frame12.Visible = True
Label3.Left = 840
med2.Left = 840
End Sub

Private Sub Option1_Click()
Command33.Visible = False
Command32.Visible = True
End Sub

Private Sub Option2_Click()
Command33.Visible = True
Command32.Visible = False
End Sub

Private Sub Option3_Click()
VScroll3.Enabled = True
End Sub

Private Sub raio_GotFocus()
raio.SelStart = 0
raio.SelLength = Len(raio)
End Sub
Private Sub raio_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then alfa.SetFocus
End Sub

Private Sub retmenu_Click()
Call filelinhaux_Click
Frame22.Visible = False

Frame6.Visible = False
Command30.Enabled = True
vsescala.Enabled = False
vslinha.Enabled = False

Frame1.Visible = False

Frame12.Visible = False
Frame11.Visible = False
End Sub

Private Sub salvproj18_Click()
Dim variavslav As String
If Form20.p1 = Empty Then
    Form20.Visible = True
    Exit Sub
End If

If Form20.p1 <> Empty Then
 Form21.Visible = True
 Form21.Command1.Visible = False
 Form21.Command4.Visible = True
 Form21.Command2.Visible = False
End If
End Sub

Private Sub StyleDash_Click()

    UnCheckStyles
    styledash.Checked = True
    Form18.DrawStyle = 1

End Sub

Private Sub StyleDot_Click()

    UnCheckStyles
    styledot.Checked = True
    Form18.DrawStyle = 2

End Sub

Private Sub StyleFilled_Click()

    stylefilled.Checked = Not stylefilled.Checked
    If stylefilled.Checked Then
        Form18.FillStyle = 0
    Else
        Form18.FillStyle = 1
    End If
    
End Sub

Private Sub StyleSolid_Click()

    UnCheckStyles
    stylesolid.Checked = True
    Form18.DrawStyle = 0
    
End Sub


Private Sub VScroll3_Change()
Dim escalax, fatx As Double
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

Form18.valasa = linhasf.y1 + ((Form2.hfs * (Form18.VScroll3.Value / 10)) * escalax)
Form18.hasamed = (Form2.hfs * (1 - Form18.VScroll3 / 10))
Form18.hasamed = Format(Form18.hasamed, "0.00")

Form18.VSasamed = Form18.VScroll3

If Form18.Option3.Value = True Then
Form18.Cls
Form18.MousePointer = 1
Call Command5_Click
End If

End Sub

Private Sub vsescala_Change()
Form18.Cls

insereasa.Enabled = True
esc18 = vsescala.Value / 10

Call Command32_Click

End Sub

Private Sub vslinha_Change()


Dim escalax, fatx, raio As Double
fatx = Form18.esc18   'fator de escala
escalax = 800 * fatx

posilinha = vslinha.Value
linhaX.y1 = posilinha
linhaX.y2 = posilinha

'......... Linha do leme
linhaleme.y1 = linhaX.y1 - (Form2.bvt * escalax)
linhaleme.y2 = linhaX.y2 - (Form2.bvt * escalax)

'......... Localização das lihas da fuselagem

linhasf.y1 = linhaX.y1 - ((Form2.hfs * 0.4) * escalax)
linhasf.y2 = linhaX.y1 - ((Form2.hfs * 0.4) * escalax)

linhaif.y1 = linhaX.y1 + ((Form2.hfs * 0.6) * escalax)
linhaif.y2 = linhaX.y1 + ((Form2.hfs * 0.6) * escalax)

'Determinando a altura do final da fuselagem
linhafinal.Visible = True
linhafinal.x1 = linhaY.x1 + (Form2.lfs * escalax)
linhafinal.X2 = linhaY.x1 + (Form2.lfs * escalax)

linhafinal.y1 = linhaX.y1 - ((Form2.lfs * 0.025) * escalax)
linhafinal.y2 = linhaX.y1 + ((Form2.lfs * 0.025) * escalax)

linhaX.Visible = True
linhaY.Visible = True

' Posição Y

bavt.y1 = linhaX.y1
bfvt.y1 = linhaX.y1

bavt.y2 = linhaleme.y1
bfvt.y2 = linhaleme.y1

Form18.valasa = linhasf.y1 + ((Form2.hfs * (Form18.VScroll3.Value / 10)) * escalax)
Form18.VSasamed = Form18.VScroll3
End Sub



Private Sub width1_Click()

    Form18.DrawWidth = 1
    
End Sub

Private Sub Width2_Click()

    Form18.DrawWidth = 2

End Sub

Private Sub Width3_Click()

    Form18.DrawWidth = 3
    
End Sub

Private Sub word_Click()
Dim Tipo, asa, trem, sistema, regime, We, Wf, Wcu, Wtot As String

We = Requisitos.weR / 2.2
We = Format(We, "0.00")
Wf = Requisitos.wfR / 2.2
Wf = Format(Wf, "0.00")
Wcu = Requisitos.CPR / 2.2
Wcu = Format(Wcu, "0.00")
Wtot = Requisitos.woestimadoR / 2.2
Wtot = Format(Wtot, "0.00")

'REQUERIMENTO

If Form1.posialta.Value = True Then asa = "Asa alta"
If Form1.posimedia.Value = True Then asa = "Asa média"
If Form1.posibaixa.Value = True Then asa = "Asa baixa"

If Form1.tremconv.Value = True Then trem = "Convencional -"
If Form1.tremtric.Value = True Then trem = "Triciclo -"

If Form1.Checkretratil = Checked Then sistema = "Trem retrátil"
If Form1.Checkretratil = Unchecked Then sistema = " Trem fixo"

''''INICIO DO TEXTO WORD
On Error Resume Next
MkDir App.Path + "\" + Form20.p1
Open App.Path + "\" + Form20.p1 + "\" + Trim(Trim(Form20.p1) & ".DOC") For Append As #1

Print #1, ""
Print #1, ""
Print #1, Tab(25); "PROJETO CONCEITUAL DE AERONAVES"
Print #1, ""
Print #1, ""
Print #1, Tab(20); "REQUISITOS OPERACIONAIS"

Print #1, ""
Print #1, Tab(1); "Nome do projeto:"; Tab(60); Form20.p1
Print #1, Tab(1); "Nome do projetista:"; Tab(60); Form20.p2
Print #1, Tab(1); "Data do projeto:"; Tab(60); Form20.datatual
Print #1, ""
Print #1, Tab(1); "Posição da asa:"; Tab(60); asa
Print #1, Tab(1); "Trem de pouso:"; Tab(60); trem + sistema
Print #1, Tab(1); "Motor:"; Tab(60); Form1.motor; " Hp"
Print #1, Tab(1); "Peso do avião:"; Tab(60); Form1.wo; " kgf"
Print #1, Tab(1); "Coeficiente de sustentação máx. Asa Cl:"; Tab(60); Form1.clmaximo
Print #1, Tab(1); "Coeficiente de sustentação máx. Asa/flap Clflap:"; Tab(60); Form1.clmaxFLAP
Print #1, Tab(1); "Coeficiente de sustentação máx. Empenagem Clht:"; Tab(60); Form1.clmaxht
Print #1, Tab(1); "Alongamento da asa :"; Tab(60); Form1.arw
Print #1, Tab(1); "Alongamento da empenagem horizontal:"; Tab(60); Form1.ARht
Print #1, Tab(1); "Alongamento da empenagem vertical:"; Tab(60); Form1.ARvt
Print #1, Tab(1); "Afilamento da asa :"; Tab(60); Form1.Afilasa
Print #1, Tab(1); "Afilamento da empenagem horizontal:"; Tab(60); Form1.AfilaHt
Print #1, Tab(1); "Afilamento da empenagem vertical:"; Tab(60); Form1.Afilavt
Print #1, Tab(1); "Área adicional"; Tab(60); Form1.adicionarea; " m2"
Print #1, Tab(1); "Diametro da bequilha"; Tab(60); Form1.diambeq; " cm"
Print #1, Tab(1); "Espessura da bequilha"; Tab(60); Form1.espbeq; " cm"
Print #1, Tab(1); "Diametro do trem principal"; Tab(60); Form1.diamtp; " cm"
Print #1, Tab(1); "Espessura do trem principal"; Tab(60); Form1.esptp; " cm"
Print #1, Tab(1); "Velocidade de Stall"; Tab(60); Form1.vstol; "  km/h"
Print #1, ""
Print #1, ""
Print #1, ""; Tab(20); "ALTERAÇÕES"
Print #1, ""
Print #1, Tab(1); "Alteração - Volume de Ht"; Tab(60); Form1.volumeht; " %"
Print #1, Tab(1); "Alteração - Volume de Vt"; Tab(60); Form1.volumevt; " %"
Print #1, Tab(1); "Alteração - Comprimento da fuselagem"; Tab(60); Form1.fuse; " %"
Print #1, Tab(1); "Alteração - Largura da fuselagem"; Tab(60); Form1.largo; " %"
Print #1, Tab(1); "Alteração - Altura da fuselagem"; Tab(60); Form1.alto; " %"
Print #1, Tab(1); "Alteração - Nariz (Hélice-CAasa)"; Tab(60); Form1.nariz; " %"
Print #1, Tab(1); "Alteração - Margem estática "; Tab(60); Form1.margestat; " %"
Print #1, ""
Print #1, ""
Print #1, ""; Tab(20); "CÁLCULOS INICIAIS"
Print #1, ""
Print #1, Tab(1); "Peso Vazio"; Tab(60); We; " kgf"
Print #1, Tab(1); "Peso de Combustivel"; Tab(60); Wf; " kgf"
Print #1, Tab(1); "Peso de Carga Util"; Tab(60); Wcu; " kgf"
Print #1, Tab(1); "Peso Total"; Tab(60); Wtot; " kgf"
Print #1, ""
Print #1, Tab(1); "Fim da página"

'DIMENSIONAMENTO
'Pagina 2
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, Tab(25); "I - DIMENSIONAMENTO"
Print #1, ""
Print #1, Tab(25); "1 - ASA"
Print #1, Tab(1); "Aerofólio utilizado:"; Tab(60); Form2.perfasa
Print #1, Tab(1); "Espessura do aerofólio da asa:"; Tab(60); Form2.espessura; " %"
Print #1, Tab(1); "Área da Asa:"; Tab(60); Form2.Sw; " m2"
Print #1, Tab(1); "Envergadura da Asa:"; Tab(60); Form2.bw; " m"
Print #1, Tab(1); "Corda da Raiz da Asa:"; Tab(60); Form2.crw; " m"
Print #1, Tab(1); "Corda Média da Asa:"; Tab(60); Form2.cmw; " m"
Print #1, Tab(1); "Corda da Ponta da Asa:"; Tab(60); Form2.cpw; " m"
Print #1, Tab(1); "Envergadura do Aileron:"; Tab(60); Form2.bail; " m"
Print #1, Tab(1); "Corda da Raiz do Aileron:"; Tab(60); Form2.crail; " m"
Print #1, Tab(1); "Corda da Ponta do Aileron:"; Tab(60); Form2.cpail; " m"
Print #1, ""
Print #1, Tab(25); "2 - EMPENAGEM HORIZONTAL"
Print #1, Tab(1); "Aerofólio utilizado:"; Tab(60); Form2.perfht
Print #1, Tab(1); "Espessura do aerofólio:"; Tab(60); Form2.espesht; " %"
Print #1, Tab(1); "Área da Emp.Horizontal:"; Tab(60); Form2.sht; " m2"
Print #1, Tab(1); "Envergadurada Emp.Horizontal:"; Tab(60); Form2.bht; " m"
Print #1, Tab(1); "Corda da Raiz da Emp. Horizontal:"; Tab(60); Form2.crht; " m"
Print #1, Tab(1); "Corda Média da Emp. Horizontal:"; Tab(60); Form2.cmht; " m"
Print #1, Tab(1); "Corda da Ponta da Emp. Horizontal:"; Tab(60); Form2.cpht; " m"
Print #1, Tab(1); "Área do Profundor:"; Tab(60); Form2.sprofund; " m2"
Print #1, Tab(1); "Envergadura do Profundor:"; Tab(60); Form2.bprofund; " m"
Print #1, Tab(1); "Corda da Raiz do Profundor:"; ; Tab(60); Form2.crprofund; " m"
Print #1, Tab(1); "Corda da Ponta do Profundor:"; Tab(60); Form2.cpprofund; " m"
Print #1, ""
Print #1, Tab(25); "3 - EMPENAGEM VERTICAL"
Print #1, ""
Print #1, Tab(1); "Aerofólio utilizado:"; Tab(60); Form2.perfvt
Print #1, Tab(1); "Espessura do aerofólio:"; Tab(60); Form2.espesvt; " %"
Print #1, Tab(1); "Área da Emp. Vertical:"; Tab(60); Form2.svt; " m2"
Print #1, Tab(1); "Envegadura da Emp. Vertical:"; Tab(60); Form2.bvt; " m"
Print #1, Tab(1); "Corda da Raiz da Emp. Vertical:"; Tab(60); Form2.crvt; " m"
Print #1, Tab(1); "Corda Média da Emp. Vertical:"; Tab(60); Form2.cmvt; " m"
Print #1, Tab(1); "Corda da Ponta da Emp. Vertical:"; Tab(60); Form2.cpvt; " m"
Print #1, Tab(1); "Área do Leme de direção:"; Tab(60); Form2.slemdir; " m2"
Print #1, Tab(1); "Envergadura do Leme de direção:"; Tab(60); Form2.blemdir; " m2"
Print #1, Tab(1); "Corda da Raiz do Leme de direção:"; Tab(60); Form2.crlemdir; " m"
Print #1, Tab(1); "Corda da Ponta do Leme de direção:"; Tab(60); Form2.cplemdir; " m"
Print #1, ""

Print #1, Tab(25); "4 - FUSELAGEM"
Print #1, ""
Print #1, Tab(1); "Comprimento da Fuselagem:"; Tab(60); Form2.lfs; " m"
Print #1, Tab(1); "Largura da Fuselagem:"; Tab(60); Form2.wfs; " m"
Print #1, Tab(1); "Altura da Fuselagem:"; Tab(60); Form2.hfs; " m"
Print #1, Tab(1); "Volume da Fuselagem:"; Tab(60); Form2.volcab; " m3"

Print #1, ""

Print #1, Tab(25); "5 - FLAP"
Print #1, ""
Print #1, Tab(1); "Tipo de Flap:"; Tab(60); Form2.tipoflap
Print #1, Tab(1); "Envegadura do Flap:"; Tab(60); Form2.bflap; " m"
Print #1, Tab(1); "Corda Média do Flap:"; Tab(60); Form2.cmflap; " m"
Print #1, ""
Print #1, Tab(1); "Fim da página";
'AERODINÂMICA
Print #1, ""
Print #1, ""
Print #1, Tab(25); "II - AERODINÂMICA"
Print #1, ""
Print #1, Tab(25); "1 - ASA"
Print #1, ""

Print #1, Tab(1); "Angulo alfa para Clmáx.s/flap:"; Tab(60); Form7.alfasamax;
Print #1, Tab(1); "Angulo alfa para Clmáx.c/flap:"; Tab(60); Form7.alfaFLAPm;
Print #1, Tab(1); "Angulo alfa para Clo.s/flap:"; Tab(60); Form7.alfasazero;
Print #1, Tab(1); "Angulo alfa para Clo.c/flap:"; Tab(60); Form7.alfaFLAPo;
Print #1, Tab(1); "Coeficiente de arrasto do perfil:"; Tab(60); Form7.cdperfasa;
Print #1, Tab(1); "Coeficiente Cl3D.s/flap:"; Tab(60); Form7.cl3dasa;
Print #1, Tab(1); "Coeficiente Cl3D.c/flap:"; Tab(60); Form7.cl3dFLAP;
Print #1, Tab(1); "dCl/dalfa - s/flap:"; Tab(60); Form7.slopeasa; " /rad"
Print #1, Tab(1); "dCl/dalfa - c/flap:"; Tab(60); Form7.slopeFLAP; " /rad"
Print #1, Tab(1); "Área molhada da asa:"; Tab(60); Form7.swetasa; " m2"
Print #1, ""
Print #1, Tab(25); "2 - ESTABILIZADOR HORIZONTAL"
Print #1, ""
Print #1, Tab(1); "Angulo alfa para Clmáximo:"; Tab(60); Form7.alfahtmax;
Print #1, Tab(1); "Angulo alfa para Clzero:"; Tab(60); Form7.alfahtzero;
Print #1, Tab(1); "Coeficiente de arrasto do perfil:"; Tab(60); Form7.cdperfht;
Print #1, Tab(1); "Coeficiente de sustentação 3D:"; Tab(60); Form7.cl3dht;
Print #1, Tab(1); "Inclinação da curva de sustentação:"; Tab(60); Form7.slopeht; " /rad"
Print #1, Tab(1); "Área molhada do Estabilizador horizontal:"; Tab(60); Form7.swethorz; " m2"
Print #1, ""
Print #1, Tab(25); "3 - ESTABILIZADOR VERTICAL"
Print #1, ""
Print #1, Tab(1); "Coeficiente de arrasto do perfil:"; Tab(60); Form7.cdperfvt;
Print #1, Tab(1); "Área molhada do Estabilizador Vertical:"; Tab(60); Form7.swetvert; " m2"
Print #1, ""
Print #1, Tab(25); "4 - FUSELAGEM"
Print #1, ""
Print #1, Tab(1); "Área molhada da Fuselagem:"; Tab(60); Form7.swetfuse; " m2"
Print #1, ""
Print #1, Tab(25); "5 - SUSTENTAÇÃO"
Print #1, ""
Print #1, Tab(1); "Carga alar de projeto:"; Tab(60); Form7.wsproj; " kgf/m2"
Print #1, Tab(1); "Cl em estol:"; Tab(60); Form7.clk1;
Print #1, Tab(1); "Cl em decolagem:"; Tab(60); Form7.cldecol;
Print #1, Tab(1); "Cl em subida:"; Tab(60); Form7.clk2;
Print #1, Tab(1); "Cl em velocidade máxima:"; Tab(60); Form7.clk4;
Print #1, ""
Print #1, Tab(25); "6 - ARRASTO"
Print #1, ""
Print #1, Tab(1); "Fator de placa plana:"; Tab(60); Form7.placaplan; " m2"
Print #1, Tab(1); "Cdtotal em estol:"; Tab(60); Form7.cdtotstall;
Print #1, Tab(1); "Cdtotal em decolagem:"; Tab(60); Form7.cddecol;
Print #1, Tab(1); "Cdtotal em subida:"; Tab(60); Form7.cdtotsubida;
Print #1, Tab(1); "Cdtotal em velocidade máxima:"; Tab(60); Form7.cdtotvmax;
Print #1, ""
Print #1, Tab(25); "7 - EFICIENCIA AERODINÂMICA"
Print #1, ""

Print #1, Tab(1); "Melhor angulo de planeio:"; Tab(60); Form7.anplan; " º"
Print #1, Tab(1); "L/D em estol:"; Tab(60); Form7.ldk1;
Print #1, Tab(1); "L/D em decolagem:"; Tab(60); Form7.lddecol;
Print #1, Tab(1); "L/D em subida:"; Tab(60); Form7.ldk2;
Print #1, Tab(1); "L/D em velocidade máxima:"; Tab(60); Form7.ldk4;
Print #1, ""
Print #1, Tab(1); "Fim da página";

'PERFORMANCE
Print #1, ""
Print #1, ""
Print #1, Tab(25); "III- PERFORMANCE"
Print #1, ""
Print #1, Tab(25); "1 - MOTORIZAÇÃO"
Print #1, ""

Print #1, Tab(1); "Especificação do motor:"; Tab(60); Form8.mtr;
Print #1, Tab(1); "Peso do motor:"; Tab(60); Form8.pesomot; " kgf"
Print #1, Tab(1); "HP nominal do motor - original:"; Tab(60); Form8.hpnominal; " HP"
Print #1, Tab(1); "RPM nominal do motor - original:"; Tab(60); Form8.rpmf1; " RPM"
Print #1, ""
Print #1, Tab(25); "2 - HÉLICE"
Print #1, ""
Print #1, Tab(1); "Especificada por:"; Tab(60); Form8.cahel;
Print #1, Tab(1); "Diametro da hélice:"; Tab(60); Form8.diahel; " pol."
Print #1, Tab(1); "Passo da hélice:"; Tab(60); Form8.pashel; " pol."
Print #1, Tab(1); "Avanço da hélice:"; Tab(60); Form8.avahel; " pol."
Print #1, Tab(1); "Eficiencia da hélice:"; Tab(60); Form8.NI; " %"
Print #1, ""
Print #1, Tab(25); "3 - VELOCIDADE"
Print #1, ""
Print #1, Tab(1); "Velocidade de stall sem flap:"; Tab(40); Form8.vst2; " Km/h"; "     com flap:"; Form8.vstol; " Km/h"
Print #1, Tab(1); "Velocidade de decolagem sem flap:"; Tab(40); Form8.vdecolsf; " Km/h"; "     com flap:"; Form8.vdecol; " Km/h"
Print #1, Tab(1); "Velocidade de subida (ideal):"; Tab(40); Form8.vsubida; " Km/h"
Print #1, Tab(1); "Velocidade máxima:"; Tab(40); Form8.vmax; " Km/h"
Print #1, Tab(1); "Velocidade de toque no pouso sem flap:"; Tab(40); Form8.toque2; " Km/h"; "      com flap:"; Form8.toque; " Km/h"
Print #1, Tab(1); "Velocidade de planeio:"; Tab(40); Form8.velplan; " Km/h"
Print #1, Tab(1); "Razão de subida:"; Tab(40); Form8.climb; " ft/min"
Print #1, ""
Print #1, Tab(25); "4 - POTENCIA"
Print #1, ""
Print #1, Tab(1); "Potencia disponivel real:"; Tab(60); Form8.hpreal; " HP"
Print #1, Tab(1); "Potencia requerida em Vstall:"; Tab(60); Form8.potreqstol; " HP"
Print #1, Tab(1); "Potencia requerida em Vdecolagem:"; Tab(60); Form8.potreqdec; " HP"
Print #1, Tab(1); "Potencia requerida em Vsubida:"; Tab(60); Form8.potreqsub; " HP"
Print #1, Tab(1); "Potencia requerida em Vmáxima:"; Tab(60); Form8.potreq; " HP"
Print #1, Tab(1); "Sobra de potencia - Vstall/Vmáxima:"; Tab(60); Form8.hpreal; " HP"
Print #1, ""
Print #1, Tab(25); "5 - FORÇA E TRAÇÃO"
Print #1, ""
Print #1, Tab(1); "Força de arrasto em Vstall:"; Tab(60); Form8.FDstol; " kgf"
Print #1, Tab(1); "Força de arrasto em Vdecolagem:"; Tab(60); Form8.FDdecol; " kgf"
Print #1, Tab(1); "Força de arrasto em Vsubida:"; Tab(60); Form8.Fdsubida; " kgf"
Print #1, Tab(1); "Força de arrasto em Vmáxima:"; Tab(60); Form8.FD; " kgf"
Print #1, Tab(1); "Tração estática:"; Tab(60); Form8.tracestatica; " kgf"
Print #1, Tab(1); "Tração dinâmica:"; Tab(60); Form8.tradin; " kgf"
Print #1, ""
Print #1, Tab(25); "6 - GERAIS"
Print #1, ""
Print #1, Tab(1); "Angulo de subida(ideal):"; Tab(40); Form8.angclimb; " º"
Print #1, Tab(1); "Distancia de decolagem s/flap:"; Tab(40); Form8.ldecsf; " m"; "    com flap:"; Form8.Ldecol; " m"
Print #1, Tab(1); "Distancia de pouso s/flap:"; Tab(40); Form8.lpousf; " m"; "    com flap:"; Form8.Lpouso; " m"
Print #1, Tab(1); "Relação L/D e W/T:"; Tab(40); Form8.ldtw;
Print #1, Tab(1); "Alcance:"; Tab(40); Requisitos.alcanceBD; " nm"
Print #1, Tab(1); "Razão de subida:"; Tab(40); Form8.climb; " ft/min"
Print #1, Tab(1); "CAFE challenge:"; Tab(40); Form8.cafechal;
Print #1, Tab(1); "CAFE Triviathon:"; Tab(40); Form8.cafetria;
Print #1, ""
Print #1, Tab(1); "Fim da página";

'ESTABILIDADE
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, Tab(25); "IV - ESTABILIDADE"
Print #1, ""
Print #1, ""
Print #1, Tab(1); "ABREVIATURAS"
Print #1, Tab(1); "Etb.Horz. = Estabilizador horizontal"
Print #1, Tab(1); "C.A-asa = Centro aerodinâmico da asa"
Print #1, Tab(1); "C.A-ht = Centro aerodinâmico do estabilizador horizontal"
Print #1, Tab(1); "C.G = Centro de gravidade"
Print #1, Tab(1); "L.R = Linha de refrencia = eixo da hélice"
Print #1, Tab(1); "L.A = Linha de corda da asa"
Print #1, Tab(1); "P.N = Ponto neutro"
Print #1, ""
Print #1, ""
Print #1, Tab(1); "Localização da corda média da asa:"; Tab(60); Form10.yasa; " m"
Print #1, Tab(1); "Localização da corda média do Etb.Horz.:"; Tab(60); Form10.yht; " m"
Print #1, Tab(1); "Volume da empenagem horizontal:"; Tab(60); Form10.volht;
Print #1, Tab(1); "Volume da empenagem vertical:"; Tab(60); Form10.volvt;
Print #1, Tab(1); "Fator de eficiencia do Etb.Horz.:"; Tab(60); Form10.eficiht;
Print #1, Tab(1); "Variação do angulo de downwash da asa:"; Tab(60); Form10.DEDA;
Print #1, Tab(1); "Força de sustentação da asa:"; Tab(60); Form10.sustentasa; " kgf"
Print #1, Tab(1); "Força de sustentação do Etb.Horz.:"; Tab(60); Form10.sustentaht; " kgf"
Print #1, Tab(1); "Angulo de incidencia da asa:"; Tab(60); Form10.incidasa; " º"
Print #1, Tab(1); "Angulo de incidencia do Etb.Horz.:"; Tab(60); Form10.incidaht; " º"
Print #1, Tab(1); "Localização do Etb.Horz. referente a L.A:"; Tab(60); Form10.htail; " m"
Print #1, Tab(1); "Localização do C.A-asa na corda média:"; Tab(60); Form10.CAasa; " m"
Print #1, Tab(1); "Localizaçãodo C.A-ht na corda média:"; Tab(60); Form10.CAht; " m"
Print #1, Tab(1); "Distancia C.A da asa ao C.A-ht (LT):"; Tab(60); Form10.LT; " m"
Print #1, Tab(1); "Localização do Ponto neutro P.N:"; Tab(60); Form10.PN; " %"
Print #1, Tab(1); "Localização do Centro de gravidade C.G:"; Tab(60); Form10.cg; " %"
Print #1, Tab(1); "Margem estática em porcentagem:"; Tab(60); Form10.MARGEM; " %"
Print #1, Tab(1); "Distancia entre P.N e C.G:"; Tab(60); Form10.PNCG; " m"
Print #1, Tab(1); "Distancia do L.R ao P.N:"; Tab(60); Form10.LRPN; " m"
Print #1, Tab(1); "Distancia do L.R ao C.A da asa:"; Tab(60); Form10.LRCAasa; " m"
Print #1, Tab(1); "Distancia do L.R ao C.A do Etb.Horz.:"; Tab(60); Form10.LRCAht; " m"
Print #1, Tab(1); "Distancia do L.R ao C.G:"; Tab(60); Form10.LRCG; " m"
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, Tab(25); "NOTA:"
Print #1, Tab(1); "O valor do C.G calculado nesta planilha deverá ser equivalente"
Print #1, Tab(1); "ao calculado estáticamente em Pesos e Estações"
Print #1, ""
Print #1, ""
Print #1, Tab(1); "Fim da página";

'ESTRUTURA DA ASA
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, Tab(15); "V - ESTRUTURA DA ASA"
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, Tab(25); "ABREVIATURAS"
Print #1, Tab(1); "CL = Coeficiente de sustentação do perfil"""
Print #1, Tab(1); "Q = Carga aerodinâmica "
Print #1, Tab(1); "M = Momento fletor"
Print #1, Tab(1); "V = Esforço cortante"
Print #1, Tab(1); "L = Largura da viga"
Print #1, Tab(1); "H = Altura da viga"
Print #1, Tab(1); "e = Espessura da viga"
Print #1, ""
Print #1, ""
Print #1, Tab(25); "TABELA"
Print #1, Tab(1); "ESTAÇÃO:"; Tab(15); "LOCAL"; Tab(27); "CL"; Tab(39); "Q (kgf)"; Tab(51); "M (kgf x m)"; Tab(63); "V (kgf)"
Print #1, ""
Print #1, Tab(1); "Estação 01:"; Tab(15); Form11.e0; " m"; Tab(27); Form11.cl0; Tab(39); Form11.q0; Tab(51); Form11.m0; Tab(63); Form11.v0
Print #1, Tab(1); "Estação 02:"; Tab(15); Form11.e1; " m"; Tab(27); Form11.cl1; Tab(39); Form11.q1; Tab(51); Form11.m1; Tab(63); Form11.v1
Print #1, Tab(1); "Estação 03:"; Tab(15); Form11.e2; " m"; Tab(27); Form11.cl2; Tab(39); Form11.q2; Tab(51); Form11.m2; Tab(63); Form11.v2
Print #1, Tab(1); "Estação 04:"; Tab(15); Form11.e3; " m"; Tab(27); Form11.cl3; Tab(39); Form11.q3; Tab(51); Form11.m3; Tab(63); Form11.v3
Print #1, Tab(1); "Estação 05:"; Tab(15); Form11.e4; " m"; Tab(27); Form11.cl4; Tab(39); Form11.q4; Tab(51); Form11.m4; Tab(63); Form11.v4
Print #1, Tab(1); "Estação 06:"; Tab(15); Form11.e5; " m"; Tab(27); Form11.cl5; Tab(39); Form11.q5; Tab(51); Form11.m5; Tab(63); Form11.v5
Print #1, Tab(1); "Estação 07:"; Tab(15); Form11.e6; " m"; Tab(27); Form11.cl6; Tab(39); Form11.q6; Tab(51); Form11.m6; Tab(63); Form11.v6
Print #1, Tab(1); "Estação 08:"; Tab(15); Form11.e7; " m"; Tab(27); Form11.cl7; Tab(39); Form11.q7; Tab(51); Form11.m7; Tab(63); Form11.v7
Print #1, Tab(1); "Estação 09:"; Tab(15); Form11.e8; " m"; Tab(27); Form11.cl8; Tab(39); Form11.q8; Tab(51); Form11.m8; Tab(63); Form11.v8
Print #1, Tab(1); "Estação 10:"; Tab(15); Form11.e9; " m"; Tab(27); Form11.cl9; Tab(39); Form11.q9; Tab(51); Form11.m9; Tab(63); Form11.v9
Print #1, Tab(1); "Estação 11:"; Tab(15); Form11.e10; " m"; Tab(27); Form11.cl10; Tab(39); Form11.q10; Tab(51); Form11.m10; Tab(63); Form11.v10
Print #1, ""
Print #1, ""
Print #1, Tab(25); "DIMENSIONAMENTO"
Print #1, ""
Print #1, Tab(1); "Secção transversal da longarina tipo:"; Tab(60); Form11.section;
Print #1, Tab(1); "Mesas da longarina em madeira de:"; Tab(60); Form11.materialM;
Print #1, Tab(1); "Alma da longarina em madeira de:"; Tab(60); Form11.materialA;
Print #1, Tab(1); "Tensão admissível a tração - mesa:"; Tab(60); Form11.sigtrac; " kgf/cm2"
Print #1, Tab(1); "Tensão admissível a compressão - mesa:"; Tab(60); Form11.sigcomp; " kgf/cm2"
Print #1, Tab(1); "Tensão admissível ao cizalhamento:"; Tab(60); Form11.sigcizal; " %"
Print #1, Tab(1); "Altura da viga:"; Tab(60); Form11.altmax; " cm"
Print #1, Tab(1); "Largura da viga:"; Tab(60); Form11.largamax; " cm"
Print #1, Tab(1); "Mesa superior (L x H):"; Tab(60); Form11.mesacp; " cm"
Print #1, Tab(1); "Mesa inferior (L x H):"; Tab(60); Form11.mesatc; " cm"
Print #1, Tab(1); "Alma (e x H):"; Tab(60); Form11.alma; " cm"
Print #1, Tab(1); "Nº de Gs positivos:"; Tab(60); Form13.gsatual;
Print #1, Tab(1); "Nº de Gs negativos:"; Tab(60); Form13.gsnegatual;
Print #1, ""

Print #1, ""
Print #1, ""
Print #1, Tab(1); "Fim da página";

'PESO E BALANCEAMENTO
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, Tab(1); "ABREVIATURAS"
Print #1, Tab(1); "C.G = Centro de gravidade"
Print #1, Tab(1); "L.R = Linha de referencia - eixo da hélice "
Print #1, Tab(1); "P.N = Ponto neutro"
Print #1, ""
Print #1, ""
Print #1, Tab(25); "TABELA"
Print #1, ""
Print #1, Tab(1); "EQUIPAMENTO:"; Tab(40); "PESO (gr)"; Tab(65); "ESTAÇÃO (cm)";
Print #1, Tab(1); "Conjunto motor:"; Tab(40); Form14.p1; Tab(65); Form14.d1;
Print #1, Tab(1); "Reservatório + combustível:"; Tab(40); Form14.p2; Tab(65); Form14.d2;
Print #1, Tab(1); "Asa:"; Tab(40); Form14.p11; Tab(65); Form14.d11;
Print #1, Tab(1); "Fuselagem:"; Tab(40); Form14.p12; Tab(65); Form14.d12;
Print #1, Tab(1); "Piloto e co-piloto:"; Tab(40); Form14.p5; Tab(65); Form14.d5;
Print #1, Tab(1); "Ocupantes:"; Tab(40); Form14.p6; Tab(65); Form14.d6;
Print #1, Tab(1); "Bagagem:"; Tab(40); Form14.p7; Tab(65); Form14.d7;
Print #1, Tab(1); "Estabilizador horizontal:"; Tab(40); Form14.p10; Tab(65); Form14.d10;
Print #1, Tab(1); "Estabilizador vertical:"; Tab(40); Form14.p4; Tab(65); Form14.d4;
Print #1, Tab(1); "Instrumentos:"; Tab(40); Form14.p3; Tab(65); Form14.d3;
Print #1, Tab(1); "Bequilha:"; Tab(40); Form14.p13; Tab(65); Form14.d13;
Print #1, Tab(1); "Trem principal:"; Tab(40); Form14.p14; Tab(65); Form14.d14;
Print #1, Tab(1); "Outros:"; Tab(40); Form14.p15; Tab(65); Form14.d15;
Print #1, ""
Print #1, ""
Print #1, Tab(25); "CALCULOS"
Print #1, ""
Print #1, Tab(1); "Peso total do avião nesta planilha:"; Tab(60); Form14.pt; " kgf"
Print #1, Tab(1); "Peso total do avião de projeto:"; Tab(60); Form14.w13; " kgf"
Print #1, Tab(1); "Carga na bequilha:"; Tab(60); Form14.pbq; " kgf"
Print #1, Tab(1); "Carga no trem principal:"; Tab(60); Form14.ptp; " kgf"
Print #1, Tab(1); "Localização do C.G em relação a L.R:"; Tab(60); Form14.cg; " cm"
Print #1, Tab(1); "Peso do avião sem combustível:"; Tab(60); Form14.wfuel; " kgf"
Print #1, ""
Print #1, ""
Print #1, Tab(25); "PONTO NEUTRO E C.Gs"
Print #1, ""
Print #1, Tab(1); "Ponto Neutro:"; Tab(60); Form14.PN; " %Cmedia"
Print #1, Tab(1); "C.G calculado em Estabilidade:"; Tab(60); Form14.CGdinamico; " %Cmedia"
Print #1, Tab(1); "C.G calculado nesta planilha:"; Tab(60); Form14.CGestatico; " %Cmedia"
Print #1, Tab(1); "C.G do avião sem combustível:"; Tab(60); Form14.CGvazio; " %Cmedia"
Print #1, Tab(1); "Passeio do C.G:"; Tab(60); Form14.passeio; "; Cm; """
Print #1, Tab(1); "Peso na roda - bequilha:"; Tab(60); Form14.bequilha; " kgf"
Print #1, Tab(1); "Peso na roda esquerda:"; Tab(60); Form14.rodaesq; " kgf"
Print #1, Tab(1); "Peso na roda direita:"; Tab(60); Form14.rodadir; " kgf"
Print #1, ""
Print #1, ""
Print #1, Tab(1); "Fim da página";

' Trimagem - form27
If Form27.Option3.Value = True Then regime = "Subida"
If Form27.Option4.Value = True Then regime = "Veloc. máxima"
'Pagina
Print #1, ""
Print #1, ""
Print #1, ""
Print #1, Tab(25); "VII - TRIMAGEM"
Print #1, ""
Print #1, Tab(1); "Coeficiente de momento da fuselagem:"; Tab(60); Form27.cmfuse
Print #1, Tab(1); "Área da asa / área do estab. horizontal:"; Tab(60); Form27.shsw;
Print #1, Tab(1); "Distancia do CG ao CA do estab. horizontal:"; Tab(60); Form27.xachcg; " /cordas"
Print #1, Tab(1); "Coeficiente drô/dfi:"; Tab(60); Form27.drodf;
Print #1, Tab(1); "Coeficiente depson/dalfa:"; Tab(60); Form27.dedalfa;
Print #1, Tab(1); "Espessura 2w no estab. horizontal:"; Tab(60); Form27.Ew; " m"
Print #1, Tab(1); "Regime estudado:"; Tab(60); regime;
Print #1, Tab(1); "Coeficiente Cmcg:"; Tab(60); Form27.cmcgok1;
Print #1, Tab(1); "Deflexão do estab. horizontal:"; Tab(60); Form27.defgrau;
Print #1, ""
Print #1, Tab(1); "Fim da página";
Close #1

End Sub

Private Sub xx_GotFocus()
xx.SelStart = 0
xx.SelLength = Len(xx)
End Sub
Private Sub xx_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then yy.SetFocus
End Sub
Private Sub YY_GotFocus()
yy.SelStart = 0
yy.SelLength = Len(yy)
End Sub
Private Sub YY_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then xx.SetFocus
End Sub

Private Sub yy_LostFocus()

If yy = Empty Then
MsgBox "Entre com o valor das coordenadas", vbInformation
Exit Sub
End If

Dim escalax, fatx As Double
fatx = esc18   'fator de escala
escalax = 800 * fatx

Form18.Line (XStart, YStart)-(1200 + (Form18.xx * escalax), Form18.yy * escalax)
XStart = 1200 + (Form18.xx * escalax)
YStart = Form18.yy * escalax

med1.Visible = True
med2.Visible = True

med1 = xx
med2 = yy

    'Form18.label2 = "X = " & Format$(Form18.xx, "&H00000000&")
    'Form18.label3 = "Y = " & Format$(Form18.yy, "&H00000000&")
End Sub

