VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfis aerodinâmicos"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10050
   ControlBox      =   0   'False
   FillColor       =   &H00E0E0E0&
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aerofolios"
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   7955
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   10035
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   7585
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Width           =   9615
         Begin VB.PictureBox Picture7 
            BackColor       =   &H00E0E0E0&
            Height          =   1135
            Left            =   120
            ScaleHeight     =   1080
            ScaleWidth      =   9315
            TabIndex        =   69
            Top             =   6360
            Width           =   9375
            Begin VB.CommandButton Command9 
               Height          =   1075
               Left            =   0
               Picture         =   "Form4.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   89
               Top             =   0
               Width           =   9315
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00E0E0E0&
               Caption         =   "&Leme de direção"
               Height          =   915
               Left            =   6720
               TabIndex        =   84
               Top             =   60
               Width           =   2535
               Begin VB.HScrollBar hsestver 
                  Height          =   255
                  Left            =   1080
                  Max             =   50
                  Min             =   15
                  TabIndex        =   87
                  Top             =   480
                  Value           =   20
                  Width           =   1215
               End
               Begin VB.TextBox porcEV 
                  Alignment       =   2  'Center
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   240
                  Locked          =   -1  'True
                  TabIndex        =   86
                  Top             =   480
                  Width           =   615
               End
               Begin VB.TextBox origcEV 
                  Height          =   195
                  Left            =   1680
                  TabIndex        =   85
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   150
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "% Corda do Estab. vertical"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   88
                  Top             =   240
                  Width           =   1860
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00E0E0E0&
               Caption         =   "&Profundor"
               Height          =   915
               Left            =   4080
               TabIndex        =   79
               Top             =   60
               Width           =   2535
               Begin VB.HScrollBar hsesthor 
                  Height          =   255
                  Left            =   1080
                  Max             =   50
                  Min             =   15
                  TabIndex        =   82
                  Top             =   480
                  Value           =   15
                  Width           =   1215
               End
               Begin VB.TextBox porcEH 
                  Alignment       =   2  'Center
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   240
                  Locked          =   -1  'True
                  TabIndex        =   81
                  Top             =   480
                  Width           =   615
               End
               Begin VB.TextBox origcEH 
                  Height          =   195
                  Left            =   1920
                  TabIndex        =   80
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   150
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "% Corda do Estab. horizontal"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   83
                  Top             =   240
                  Width           =   2025
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E0E0E0&
               Caption         =   "&Aileron "
               Height          =   915
               Left            =   120
               TabIndex        =   70
               Top             =   60
               Width           =   3855
               Begin VB.HScrollBar hsbail 
                  Height          =   255
                  Left            =   2760
                  Max             =   100
                  Min             =   20
                  TabIndex        =   76
                  Top             =   480
                  Value           =   40
                  Width           =   855
               End
               Begin VB.TextBox porcenvail 
                  Alignment       =   2  'Center
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2040
                  Locked          =   -1  'True
                  TabIndex        =   75
                  Top             =   480
                  Width           =   615
               End
               Begin VB.TextBox porcAIL 
                  Alignment       =   2  'Center
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   120
                  Locked          =   -1  'True
                  TabIndex        =   74
                  Top             =   480
                  Width           =   615
               End
               Begin VB.HScrollBar hsaileron 
                  Height          =   255
                  Left            =   840
                  Max             =   30
                  Min             =   10
                  TabIndex        =   73
                  Top             =   480
                  Value           =   10
                  Width           =   855
               End
               Begin VB.TextBox origcail 
                  Height          =   195
                  Left            =   1320
                  TabIndex        =   72
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   150
               End
               Begin VB.TextBox origbail 
                  Height          =   195
                  Left            =   3000
                  TabIndex        =   71
                  Text            =   "Text1"
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   150
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "% Envergadura"
                  Height          =   195
                  Left            =   2040
                  TabIndex        =   78
                  Top             =   240
                  Width           =   1080
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "% Corda da asa"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   77
                  Top             =   240
                  Width           =   1110
               End
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&BIBLIOTECA DE PERFIS AERODINÂMICOS"
            Height          =   2295
            Left            =   120
            TabIndex        =   58
            Top             =   1440
            Width           =   9375
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00E0E0E0&
               Height          =   1935
               Left            =   6960
               ScaleHeight     =   1875
               ScaleWidth      =   2235
               TabIndex        =   59
               Top             =   240
               Width           =   2295
               Begin VB.CommandButton Command6 
                  BackColor       =   &H00C0C0C0&
                  Height          =   1875
                  Left            =   0
                  MouseIcon       =   "Form4.frx":892C
                  MousePointer    =   99  'Custom
                  Picture         =   "Form4.frx":8C36
                  Style           =   1  'Graphical
                  TabIndex        =   60
                  ToolTipText     =   "Clique aqui para abrir os perfis aerodnâmicos"
                  Top             =   0
                  Width           =   2235
               End
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "OBS. Para alterar o perfil de alguma superfície clique ""Modificar""."
               Height          =   195
               Left            =   240
               TabIndex        =   68
               Top             =   1920
               Width           =   4620
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "6 - Em seguida, Clique Enviar."
               Height          =   195
               Left            =   240
               TabIndex        =   67
               Top             =   1680
               Width           =   2115
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "5 - Se for utilizar Flap, clique em ""Flap"" e escolha o tipo e o angulo máximo de projeto."
               Height          =   195
               Left            =   240
               TabIndex        =   66
               Top             =   1440
               Width           =   6060
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "4 - Em ""Eficiência - eHt"", posicione o Estabilizador horizontal em relação a linha da asa."
               Height          =   195
               Left            =   240
               TabIndex        =   65
               Top             =   1200
               Width           =   6180
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "3-  Repita o procedimento para o Estabilizador horizontal e Estabilizador vertical."
               Height          =   195
               Left            =   240
               TabIndex        =   64
               Top             =   960
               Width           =   5640
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "2-  Clique ""Asa"" e abrirá os perfis para asa. Clique 2 vezes sobre o perfil de sua escolha."
               Height          =   195
               Left            =   240
               TabIndex        =   63
               Top             =   720
               Width           =   6225
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1 - Clique o botão a direita para abrir a biblioteca. ( Perfis Aerodinâmicos ) --------------------------->"
               Height          =   195
               Left            =   240
               TabIndex        =   62
               Top             =   480
               Width           =   6495
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Como utilizar a biblioteca :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   240
               TabIndex        =   61
               Top             =   240
               Width           =   2325
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H00E0E0E0&
            Height          =   2130
            Left            =   120
            ScaleHeight     =   2070
            ScaleWidth      =   9150
            TabIndex        =   57
            Top             =   1560
            Width           =   9215
            Begin MSDBGrid.DBGrid DBGrid3 
               Bindings        =   "Form4.frx":F110
               Height          =   720
               Left            =   0
               OleObjectBlob   =   "Form4.frx":F120
               TabIndex        =   90
               Top             =   0
               Width           =   9135
            End
            Begin MSDBGrid.DBGrid DBGrid2 
               Bindings        =   "Form4.frx":1068A
               Height          =   720
               Left            =   0
               OleObjectBlob   =   "Form4.frx":1069A
               TabIndex        =   91
               Top             =   0
               Width           =   9135
            End
            Begin MSDBGrid.DBGrid DBGrid1 
               Bindings        =   "Form4.frx":11C04
               Height          =   720
               Left            =   0
               OleObjectBlob   =   "Form4.frx":11C14
               TabIndex        =   92
               Top             =   0
               Width           =   9150
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00E0E0E0&
            Height          =   2415
            Left            =   5760
            ScaleHeight     =   2355
            ScaleWidth      =   3675
            TabIndex        =   46
            Top             =   3840
            Width           =   3735
            Begin VB.CommandButton Command8 
               Height          =   2355
               Left            =   0
               MouseIcon       =   "Form4.frx":13166
               MousePointer    =   99  'Custom
               Picture         =   "Form4.frx":13470
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Clique aqui para escolher a posição do estab. horizontal"
               Top             =   0
               Width           =   3675
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00E0E0E0&
               Caption         =   "&Eficiencia do Estabilizador Horizontal"
               ForeColor       =   &H00000000&
               Height          =   2295
               Left            =   120
               TabIndex        =   47
               Top             =   0
               Width           =   3495
               Begin VB.TextBox m 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFFF&
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   2400
                  Locked          =   -1  'True
                  TabIndex        =   53
                  Text            =   "0,00"
                  Top             =   1920
                  Width           =   495
               End
               Begin VB.CommandButton Command7 
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
                  Left            =   3000
                  Style           =   1  'Graphical
                  TabIndex        =   52
                  Top             =   1920
                  Width           =   255
               End
               Begin VB.PictureBox Picture5 
                  BackColor       =   &H00400000&
                  Height          =   1455
                  Left            =   120
                  ScaleHeight     =   1395
                  ScaleWidth      =   3195
                  TabIndex        =   50
                  Top             =   360
                  Width           =   3255
                  Begin VB.VScrollBar Vseficht 
                     Height          =   1400
                     Left            =   3000
                     Max             =   8
                     Min             =   1
                     TabIndex        =   51
                     Top             =   0
                     Value           =   1
                     Width           =   185
                  End
                  Begin VB.Line lasa 
                     BorderColor     =   &H008080FF&
                     BorderWidth     =   2
                     X1              =   120
                     X2              =   2640
                     Y1              =   840
                     Y2              =   840
                  End
                  Begin VB.Line Line2 
                     BorderColor     =   &H008080FF&
                     X1              =   1320
                     X2              =   1320
                     Y1              =   840
                     Y2              =   360
                  End
                  Begin VB.Shape Shape1 
                     BorderColor     =   &H008080FF&
                     FillColor       =   &H008080FF&
                     FillStyle       =   0  'Solid
                     Height          =   255
                     Left            =   1200
                     Shape           =   2  'Oval
                     Top             =   720
                     Width           =   255
                  End
                  Begin VB.Line HT 
                     BorderColor     =   &H008080FF&
                     X1              =   960
                     X2              =   1680
                     Y1              =   840
                     Y2              =   840
                  End
               End
               Begin VB.TextBox effHT 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   49
                  Text            =   "0,00"
                  Top             =   1920
                  Width           =   495
               End
               Begin VB.TextBox valefHT 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   3120
                  TabIndex        =   48
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   150
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "m :"
                  Height          =   195
                  Left            =   2160
                  TabIndex        =   55
                  Top             =   1920
                  Width           =   210
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fator eficiencia:"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   54
                  Top             =   1920
                  Width           =   1125
               End
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00E0E0E0&
            Height          =   2415
            Left            =   120
            ScaleHeight     =   2355
            ScaleWidth      =   5475
            TabIndex        =   28
            Top             =   3840
            Width           =   5535
            Begin VB.CommandButton Command5 
               BackColor       =   &H00C0C0C0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2355
               Left            =   0
               MouseIcon       =   "Form4.frx":1A2AA
               MousePointer    =   99  'Custom
               Picture         =   "Form4.frx":1A5B4
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   "Clique aqui para optar por flaps"
               Top             =   0
               Width           =   5455
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E0E0E0&
               Caption         =   "&Flap"
               Height          =   2295
               Left            =   120
               TabIndex        =   29
               Top             =   0
               Width           =   5295
               Begin VB.OptionButton semflap 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Sem flap"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   44
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton splitflap 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Tipo Split "
                  Height          =   255
                  Left            =   120
                  TabIndex        =   43
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.OptionButton fowlerflap 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Tipo Fowler "
                  Height          =   255
                  Left            =   120
                  TabIndex        =   42
                  Top             =   1800
                  Width           =   1215
               End
               Begin VB.OptionButton plainflap 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Tipo Plain"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   41
                  Top             =   840
                  Width           =   1215
               End
               Begin VB.Frame Frame7 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "&Deflexão"
                  Height          =   2295
                  Left            =   2040
                  TabIndex        =   30
                  Top             =   0
                  Width           =   3255
                  Begin VB.TextBox clsflap 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E0E0E0&
                     Height          =   285
                     Left            =   960
                     Locked          =   -1  'True
                     TabIndex        =   36
                     Top             =   1680
                     Width           =   495
                  End
                  Begin VB.PictureBox Picture6 
                     Height          =   975
                     Left            =   120
                     Picture         =   "Form4.frx":20A8E
                     ScaleHeight     =   915
                     ScaleWidth      =   2955
                     TabIndex        =   34
                     Top             =   240
                     Width           =   3015
                     Begin VB.VScrollBar VSflap 
                        Height          =   915
                        Left            =   2700
                        Max             =   5
                        Min             =   1
                        TabIndex        =   35
                        Top             =   0
                        Value           =   1
                        Width           =   255
                     End
                     Begin VB.Line flap 
                        BorderColor     =   &H00FF0000&
                        BorderWidth     =   2
                        X1              =   1800
                        X2              =   2160
                        Y1              =   480
                        Y2              =   480
                     End
                  End
                  Begin VB.TextBox deflap 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E0E0E0&
                     ForeColor       =   &H00FF0000&
                     Height          =   285
                     Left            =   2280
                     Locked          =   -1  'True
                     TabIndex        =   33
                     Text            =   "0"
                     Top             =   1320
                     Width           =   495
                  End
                  Begin VB.TextBox clmaxfinal 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E0E0E0&
                     ForeColor       =   &H00FF0000&
                     Height          =   285
                     Left            =   2280
                     Locked          =   -1  'True
                     TabIndex        =   32
                     Top             =   1680
                     Width           =   495
                  End
                  Begin VB.TextBox valflap 
                     Height          =   285
                     Left            =   2880
                     TabIndex        =   31
                     Text            =   "1"
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   150
                  End
                  Begin VB.Label Label12 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cl s/flap:"
                     Height          =   195
                     Left            =   240
                     TabIndex        =   40
                     Top             =   1680
                     Width           =   630
                  End
                  Begin VB.Label Label15 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "* Corda do flap = 0.30% da Cmédia da asa"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   180
                     Left            =   240
                     TabIndex        =   39
                     Top             =   2040
                     Width           =   2700
                  End
                  Begin VB.Label Label14 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Angulo de deflexão :                         º"
                     Height          =   195
                     Left            =   240
                     TabIndex        =   38
                     Top             =   1320
                     Width           =   2640
                  End
                  Begin VB.Label Label16 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cl c/flap:"
                     Height          =   195
                     Left            =   1560
                     TabIndex        =   37
                     Top             =   1680
                     Width           =   645
                  End
               End
            End
         End
         Begin VB.PictureBox Picture1 
            Height          =   1035
            Left            =   5880
            ScaleHeight     =   975
            ScaleWidth      =   3555
            TabIndex        =   26
            Top             =   240
            Width           =   3615
            Begin VB.CommandButton Command23 
               Enabled         =   0   'False
               Height          =   975
               Left            =   0
               MouseIcon       =   "Form4.frx":30370
               MousePointer    =   99  'Custom
               Picture         =   "Form4.frx":3067A
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   0
               Width           =   3550
            End
         End
         Begin VB.TextBox perfil3 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3960
            TabIndex        =   25
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox perfil2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2040
            TabIndex        =   24
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox perfil1 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   960
            Width           =   1815
         End
         Begin VB.PictureBox Picture8 
            Height          =   555
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   5475
            TabIndex        =   19
            Top             =   240
            Width           =   5535
            Begin VB.CommandButton Command4 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Empenagem &Vertical"
               Enabled         =   0   'False
               Height          =   495
               Left            =   3720
               MouseIcon       =   "Form4.frx":343AC
               MousePointer    =   99  'Custom
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Clique aqui para escolher o perfil do estab. vertical"
               Top             =   0
               Width           =   1755
            End
            Begin VB.CommandButton Command3 
               BackColor       =   &H00C0C0C0&
               Caption         =   "&Empenagem &Horizontal"
               Enabled         =   0   'False
               Height          =   495
               Left            =   1800
               MouseIcon       =   "Form4.frx":346B6
               MousePointer    =   99  'Custom
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Clique aqui para escolher o perfil do estab. horizontal"
               Top             =   0
               Width           =   1935
            End
            Begin VB.CommandButton Command28 
               BackColor       =   &H00C0C0C0&
               Caption         =   "&Asa"
               Enabled         =   0   'False
               Height          =   495
               Left            =   0
               MouseIcon       =   "Form4.frx":349C0
               MousePointer    =   99  'Custom
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Clique aqui para escolher o perfil da asa"
               Top             =   0
               Width           =   1815
            End
         End
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "PerfilHT"
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "PerfilVT"
         Top             =   0
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Arquivando novos perfis"
      ForeColor       =   &H00800000&
      Height          =   2275
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   10035
      Begin VB.TextBox perfil15 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox cl15 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox cd15 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox util15 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox rey15 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox porcent15 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Editar o Perfil"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4080
         TabIndex        =   3
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox angulo15 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Ajuda"
         Height          =   255
         Left            =   6960
         TabIndex        =   1
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Perfil :"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cl máximo do Perfil :"
         Height          =   195
         Left            =   4080
         TabIndex        =   15
         Top             =   480
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cd máximo do Perfil :"
         Height          =   195
         Left            =   6960
         TabIndex        =   14
         Top             =   480
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Utilizado para :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Reynolds (tunel) :"
         Height          =   195
         Left            =   4080
         TabIndex        =   12
         Top             =   1080
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espessura máxima ( % ) :"
         Height          =   195
         Left            =   6960
         TabIndex        =   11
         Top             =   1080
         Width           =   1740
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angulo em que ocorre o Cl máximo  (º) :"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   2775
      End
   End
   Begin VB.Menu filemodifica4 
      Caption         =   "&Modificar"
   End
   Begin VB.Menu cadastranovo 
      Caption         =   "&Cadastrar "
      Begin VB.Menu perfasa 
         Caption         =   "Perfil - Asa"
      End
      Begin VB.Menu estabht 
         Caption         =   "Perfil - Estab. horizontal"
      End
      Begin VB.Menu estabvt 
         Caption         =   "Perfil - Estab. vertical"
      End
   End
   Begin VB.Menu filexclui 
      Caption         =   "&Excluir"
      Begin VB.Menu excasa 
         Caption         =   "Perfil - Asa"
      End
      Begin VB.Menu excht 
         Caption         =   "Perfil - Estab. horizontal"
      End
      Begin VB.Menu excvt 
         Caption         =   "Perfil - Estab. vertical"
      End
   End
   Begin VB.Menu filefer 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "&Calculadora && Conversor"
      End
   End
   Begin VB.Menu filesai4 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command23_Click()
Command4.Enabled = True

If perfil1 = Empty Then
MsgBox "Voce não escolheu o perfil para a asa"
Exit Sub
End If

If perfil2 = Empty Then
MsgBox "Voce não escolheu o perfil para a Estabilizador horizontal"
Exit Sub
End If

If perfil3 = Empty Then
MsgBox "Voce não escolheu o perfil para a Estabilizador vertical"
Exit Sub
End If

If Form4.effHT = 0 Then
Form4.Vseficht.Value = 5
MsgBox "Escolha a posição do estabilizador horizontal"
Form4.Visible = True
Form4.Command8.Visible = False
Exit Sub
End If


Form1.Command4.Enabled = True
Form1.Command4.SetFocus

'Mantendo o cursor da biblioteca na ASA escolhida
If perfil1 <> Empty Then
variavelw = "Select * from Aerofolios where Perfil='" & perfil1.Text + "'"
Form4.Data1.RecordSource = variavelw
Form4.Data1.Refresh
End If

'Fim

Form1.clmaxFLAP = Form4.clmaxfinal
Form1.clmaxFLAP = Format(Form1.clmaxFLAP, "0.00")

'Slope sem Flap
'Form7.alfasamax = Form4.Data1.Recordset("angulo")
'Form7.alfasamax = Format(Form7.alfasamax, "0.00")
'Form7.alfasazero = Form4.Data1.Recordset("angulo0")
'Form7.alfasazero = Format(Form7.alfasazero, "0.00")
'deltaSF = Abs(Form7.alfasamax) + Abs(Form7.alfasazero)
'slopesf = (Form1.clmaximo / deltaSF) * 57.3
'Form7.slopeasa = slopesf
'Form7.slopeasa = Format(Form7.slopeasa, "0.00")
'Form7.cl3dasa = ((Form7.slopeasa / (1 + (Abs(Form7.slopeasa / (3.1415 * Form1.arw))))) / 57.3) * deltaSF
'Form7.cl3dasa = Format(Form7.cl3dasa, "0.00")

'Slope com Flap
'difCL = Abs(Form1.clmaxFLAP) - Abs(Form1.clmaximo)
'FK = 0.055
'FK2 = 0.12
'k = FK * Form4.deflap * difCL
'k1 = k
'k2 = FK2 * Form4.deflap * difCL
'k11 = Abs(Form7.alfasamax) - Abs(k)
'k22 = (Form7.alfasazero) - (k2)

'kk = Abs(k11) + Abs(k22)
'Form7.alfaFLAPm = k11
'Form7.alfaFLAPm = Format(Form7.alfaFLAPm, "0.00")
'Form7.alfaFLAPo = k22
'Form7.alfaFLAPo = Format(Form7.alfaFLAPo, "0.00")
'slopeCF = (Form1.clmaxFLAP / kk) * 57.3
'Form7.slopeFLAP = slopeCF
'Form7.slopeFLAP = Format(Form7.slopeFLAP, "0.00")
'Form7.cl3dFLAP = ((Form7.slopeFLAP / (1 + (Abs(Form7.slopeFLAP / (3.1415 * Form1.arw))))) / 57.3) * kk
'Form7.cl3dFLAP = Format(Form7.cl3dFLAP, "0.00")


Form4.Visible = False
Form4.Command1.Enabled = False

Frame5.Visible = True
Command5.Visible = True

Command28.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command9.Visible = False

Command8.Visible = True


End Sub

Private Sub Command28_Click()
DBGrid1.Height = 2040
DBGrid1.Enabled = True
DBGrid1.Visible = True

DBGrid2.Visible = False
DBGrid3.Visible = False

End Sub

Private Sub Command3_Click()
DBGrid2.Height = 2040
DBGrid2.Enabled = True
DBGrid2.Visible = True

DBGrid1.Visible = False
DBGrid3.Visible = False

End Sub


Private Sub Command4_Click()
DBGrid3.Height = 2040
DBGrid3.Enabled = True
DBGrid3.Visible = True

DBGrid1.Visible = False
DBGrid2.Visible = False

End Sub

Private Sub Command5_Click()

If perfil1 = Empty Then
MsgBox "Voce deve escolher o perfil para a asa.", vbInformation
Command5.Visible = True
Exit Sub
End If

Command5.Visible = False

If perfil1 <> Empty Then
variavelw = "Select * from Aerofolios where Perfil='" & perfil1.Text + "'"
Form4.Data1.RecordSource = variavelw
Form4.Data1.Refresh
End If

Command23.Enabled = True
End Sub

Private Sub Command6_Click()
Call filemodifica4_Click
Frame5.Visible = False

'Mantendo o cursor da biblioteca na ASA escolhida
If Form1.clmaximo <> Empty Then
If perfil1 <> Empty Then
variavelw = "Select * from Aerofolios where Perfil='" & perfil1.Text + "'"
Form4.Data1.RecordSource = variavelw
Form4.Data1.Refresh
End If
End If
'Fim

If Form1.clmaximo = Empty Then
If perfil1 = Empty Then
Form4.Data1.Refresh
End If
End If



Command28.Enabled = True
Command3.Enabled = True
Command4.Enabled = True


End Sub

Private Sub Command7_Click()
MsgBox "Determina a eficiencia do estabilizador horizontal, cauda ""T"" eficiencia 95%", vbInformation
If Form1.posialta.Value = True Or Form1.posibaixa.Value = True Or Form1.posimedia.Value = True Then
Form4.valefHT = 4
Form4.Vseficht.Value = Form4.valefHT
End If
End Sub

Private Sub Command8_Click()
If perfil1 = Empty Then
MsgBox "Voce deve escolher o perfil para a asa.", vbInformation
Command5.Visible = True
Exit Sub
End If

Command8.Visible = False

Command23.Enabled = True
End Sub


Private Sub Command9_Click()

If perfil1 = Empty Then
MsgBox "Voce deve escolher o perfil para a asa.", vbInformation
Command9.Visible = True
Exit Sub
End If

Command9.Visible = False
End Sub





Private Sub DBGrid1_DblClick()

Form1.clmaximo = Form4.Data1.Recordset("clmax")
Form1.clmaximo = Format(Form1.clmaximo, "0.00")
Form7.alfasamax = Form4.Data1.Recordset("angulo")
Form7.alfasazero = Form4.Data1.Recordset("angulo0")
Form7.cdperfasa = Form4.Data1.Recordset("cdmax")
Form2.espessura = Form4.Data1.Recordset("Espessura")
Form4.perfil1 = Form4.Data1.Recordset("Perfil")
Form2.perfasa = Form4.perfil1

variavelw = "Select * from Aerofolios where Perfil='" & perfil1.Text + "'"
Form4.Data1.RecordSource = variavelw
Form4.Data1.Refresh

Form4.clmaxfinal = Form4.Data1.Recordset("clmax")
Form4.clmaxfinal = Format(Form4.clmaxfinal, "0.00")

'Cl sem flap para calcular Fowler Flap
Form4.clsflap = Form4.Data1.Recordset("clmax")

VSflap.Value = 1

perfil1.SetFocus

If perfil1 = Empty And perfil2 = Empty And perfil3 = Empty Then Command28.SetFocus
If perfil1 <> Empty And perfil2 = Empty And perfil3 = Empty Then Command3.SetFocus
If perfil1 <> Empty And perfil2 <> Empty And perfil3 = Empty Then Command4.SetFocus

DBGrid1.Height = 480
DBGrid1.Enabled = False
End Sub
Private Sub DBGrid2_DblClick()
Form1.clmaxht = Form4.Data2.Recordset("Clmax")
Form7.alfahtmax = Form4.Data2.Recordset("Angulo")
Form7.alfahtzero = Form4.Data2.Recordset("Angulo0")
Form7.cdperfht = Form4.Data2.Recordset("Cdmax")
Form2.espesht = Form4.Data2.Recordset("Espessura")
Form4.perfil2 = Form4.Data2.Recordset("Perfil")
Form2.perfht = Form4.perfil2

variavelht = "Select * from PerfilHT where Perfil='" & perfil2.Text + "'"
Form4.Data2.RecordSource = variavelht
Form4.Data2.Refresh

DBGrid2.Height = 480
DBGrid2.Enabled = False
End Sub

Private Sub DBGrid3_DblClick()

variavelw = "Select * from Perfilasa where Perfil='" & perfil3.Text + "'"
Form4.Data3.RecordSource = variavelw


variavelvht = "Select * from Perfilht where Perfil='" & perfil3.Text + "'"
Form4.Data3.RecordSource = variavelht


variavelvt = "Select * from PerfilVT where Perfil='" & perfil3.Text + "'"
Form4.Data3.RecordSource = variavelvt


Form4.perfil3 = Form4.Data3.Recordset("Perfil")
Form2.espesvt = Form4.Data3.Recordset("Espessura")
Form7.cdperfvt = Form4.Data3.Recordset("cdmax")
Form2.perfvt = Form4.perfil3


DBGrid3.Height = 480

'Capacitando Flap
Command5.Enabled = True

'Capacitando Enviar
Command23.Enabled = True

DBGrid3.Enabled = False
End Sub

Private Sub estabht_Click()
Form28.Visible = True
Form4.Visible = False
Form28.Frame1.Visible = False
Form28.Frame2.Visible = True
Form28.Frame3.Visible = False
End Sub

Private Sub estabvt_Click()
Form28.Visible = True
Form4.Visible = False
Form28.Frame1.Visible = False
Form28.Frame2.Visible = True
Form28.Frame3.Visible = False
End Sub

Private Sub excasa_Click()
MsgBox "Verifique se existe algum projeto utilizando este perfil e então modifique-o.", vbInformation
If Command28.Enabled = False Then MsgBox " Clique Modificar para abrir a listagem de perfis", vbInformation

    If Command28.Enabled = True Then
        resp2 = MsgBox("Voce tem certeza de que deseja excluir o perfil selecionado?", vbOKCancel)

        If resp2 = vbOK Then
            Form4.Data1.Recordset.Delete
            Form4.Data1.Refresh
        End If
 
            If resp2 = vbCancel Then
                Form4.Visible = False
            End If
    End If
End Sub

Private Sub excht_Click()
MsgBox "Verifique se existe algum projeto utilizando este perfil e então modifique-o.", vbInformation
If Command28.Enabled = False Then MsgBox " Clique Modificar para abrir a listagem de perfis", vbInformation

    If Command28.Enabled = True Then
    resp2 = MsgBox("Voce tem certeza de que dejeja excluir o perfil selecionado?", vbOKCancel)

        If resp2 = vbOK Then
            Form4.Data2.Recordset.Delete
            Form4.Data2.Refresh
        End If
 
            If resp2 = vbCancel Then
                Form4.Visible = False
            End If
     End If
End Sub

Private Sub excvt_Click()
MsgBox "Verifique se existe algum projeto utilizando este perfil e então modifique-o.", vbInformation
If Command28.Enabled = False Then MsgBox " Clique Modificar para abrir a listagem de perfis", vbInformation

    If Command28.Enabled = True Then

    resp2 = MsgBox("Voce tem certeza de que dejeja excluir o perfil selecionado?", vbOKCancel)

        If resp2 = vbOK Then
            Form4.Data3.Recordset.Delete
            Form4.Data3.Refresh
        End If
 
            If resp2 = vbCancel Then
                Form4.Visible = False
            End If
     End If
End Sub

Private Sub filecalcon_Click()
Form29.Show
End Sub

Private Sub filemodifica4_Click()

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

Call Command28_Click


Command23.Enabled = True


End Sub

Private Sub filesai4_Click()

Command23.Enabled = False

Form4.Visible = False
Form4.Command1.Enabled = False

Frame5.Visible = True
Command5.Visible = True

Command28.Enabled = False
Command3.Enabled = False
Command4.Enabled = False

Command8.Visible = True

If perfil1 <> Empty Or perfil1 <> Empty Or perfil1 <> Empty Then
resp4 = MsgBox("Os perfis já escolhidos devem permanecer?", vbYesNo)
Form4.Refresh

If resp4 = vbOK Then
        Form4.Visible = False
        End If
        
If resp4 = vbNo Then
            Form4.perfil1 = Empty
            Form4.perfil2 = Empty
            Form4.perfil3 = Empty
            Form1.clmaximo = Empty
            Form1.clmaxht = Empty
            Exit Sub
            End If
End If
End Sub

Private Sub Form_Load()
    Data1.DatabaseName = SR
    Data2.DatabaseName = SR
    Data3.DatabaseName = SR
    
Form4.VSflap.Value = 1


hsaileron.Value = 20
hsbail.Value = 50

hsesthor.Value = 50
hsestver.Value = 50

End Sub

Private Sub fowlerflap_Click()
VSflap = 1
variavelw = "Select * from Aerofolios where Perfil='" & perfil1.Text + "'"
Form4.Data1.RecordSource = variavelw
Form4.Data1.Refresh
End Sub

Private Sub hsaileron_Change()
origcail = hsaileron.Value
porcAIL = hsaileron.Value / 100
End Sub

Private Sub hsbail_Change()
origbail = hsbail.Value
porcenvail = hsbail.Value / 100
If Form4.fowlerflap.Value = True Or Form4.plainflap.Value = True Or Form4.splitflap.Value = True Then
Form4.hsbail.Value = 42
MsgBox "A envergadura do aileron será fixada em 42% da semi envergadura da asa.", vbInformation
End If
End Sub

Private Sub hsesthor_Change()
origcEH = hsesthor.Value
porcEH = hsesthor.Value / 100
End Sub

Private Sub hsestver_Change()
origcEV = hsestver.Value
porcEV = hsestver.Value / 100
End Sub

Private Sub perfasa_Click()
Form28.Visible = True
Form4.Visible = False
Form28.Frame1.Visible = False
Form28.Frame2.Visible = True
Form28.Frame3.Visible = False
End Sub

Private Sub plainflap_Click()
VSflap = 1
variavelw = "Select * from Aerofolios where Perfil='" & perfil1.Text + "'"
Form4.Data1.RecordSource = variavelw
Form4.Data1.Refresh
End Sub

Private Sub semflap_Click()
VSflap.Value = 1
End Sub

Private Sub splitflap_Click()
VSflap = 1
variavelw = "Select * from Aerofolios where Perfil='" & perfil1.Text + "'"
Form4.Data1.RecordSource = variavelw
Form4.Data1.Refresh
End Sub

Private Sub Vseficht_Change()

If Vseficht.Value = 8 Then
HT.y1 = Form18.linhaleme.y1
HT.y2 = Form18.linhaleme.y2
If Form1.posialta.Value = True Then effHT = 1.09
If Form1.posimedia.Value = True Then effHT = 1.085
If Form1.posibaixa.Value = True Then effHT = 1.07
 m = -0.06
End If


If Vseficht.Value = 7 Then
HT.y1 = 860
HT.y2 = 860
If Form1.posialta.Value = True Then effHT = 1.088
If Form1.posimedia.Value = True Then effHT = 1.075
If Form1.posibaixa.Value = True Then effHT = 1.082
 m = -0.04
End If


If Vseficht.Value = 6 Then
HT.y1 = 800
HT.y2 = 800
If Form1.posialta.Value = True Then effHT = 1.085
If Form1.posimedia.Value = True Then effHT = 1.08
If Form1.posibaixa.Value = True Then effHT = 1.085
 m = -0.025
End If


If Vseficht.Value = 5 Then
HT.y1 = 740
HT.y2 = 740
If Form1.posialta.Value = True Then effHT = 1.085
If Form1.posimedia.Value = True Then effHT = 1.08
If Form1.posibaixa.Value = True Then effHT = 1.085
 m = -0.01
End If


If Vseficht.Value = 4 Then
HT.y1 = 680
HT.y2 = 680
If Form1.posialta.Value = True Then effHT = 1.085
If Form1.posimedia.Value = True Then effHT = 1.085
If Form1.posibaixa.Value = True Then effHT = 1.088
 m = 0.01
End If


If Vseficht.Value = 3 Then
HT.y1 = 620
HT.y2 = 620
If Form1.posialta.Value = True Then effHT = 1.08
If Form1.posimedia.Value = True Then effHT = 1.088
If Form1.posibaixa.Value = True Then effHT = 1.09
 m = 0.06
End If


If Vseficht.Value = 2 Then
HT.y1 = 480
HT.y2 = 480
If Form1.posialta.Value = True Then effHT = 1.07
If Form1.posimedia.Value = True Then effHT = 1.08
If Form1.posibaixa.Value = True Then effHT = 1.09
 m = 0.11
End If


If Vseficht.Value = 1 Then
HT.y1 = 360
HT.y2 = 360
If Form1.posialta.Value = True Then effHT = 1.08
If Form1.posimedia.Value = True Then effHT = 1.09
If Form1.posibaixa.Value = True Then effHT = 1.1
 m = 0.22
End If


valefHT = Vseficht.Value

End Sub

Private Sub VSflap_Change()

If perfil1 = Empty Then
'MsgBox "Inicialmente escolha os perfis para asa e estabilizadores.", vbInformation
Command5.Visible = True
Exit Sub
End If

If fowlerflap.Value = False Then kflap = 0
If fowlerflap.Value = True Then kflap = 240

        'PARA SEM FLAP

If semflap.Value = True Then
    VSflap.Value = 1
    Form4.clmaxfinal = Form4.Data1.Recordset("clmax")
    Form4.clmaxfinal = Format(Form4.clmaxfinal, "0.00")
End If

        'PARA COM FLAP

'PARA 0 GRAUS
If VSflap.Value = 1 Then

flap.x1 = 1800
flap.y1 = 480
flap.X2 = 2160
flap.y2 = 480

deflap = 0
    If semflap.Value = True Then coeflap = 1
    If plainflap.Value = True Then coeflap = 1 + Abs(0.0075 * deflap)
    If splitflap.Value = True Then coeflap = 1 + Abs(0.01 * deflap)
    If fowlerflap.Value = True Then coeflap = 1 + Abs(0.015 * deflap)
       
            clmaxfinal = Form4.Data1.Recordset("clmax") * coeflap
            Form4.clmaxfinal = Format(Form4.clmaxfinal, "0.00")
End If

'PARA 10 GRAUS

If VSflap.Value = 2 Then

flap.x1 = 1800 + kflap
flap.y1 = 480
flap.X2 = 2160 + kflap
flap.y2 = 540

deflap = 10
    If semflap.Value = True Then coeflap = 1
    If plainflap.Value = True Then coeflap = 1 + Abs(0.0075 * deflap)
    If splitflap.Value = True Then coeflap = 1 + Abs(0.01 * deflap)
    If fowlerflap.Value = True Then coeflap = 1 + Abs(0.015 * deflap)

            clmaxfinal = Form4.Data1.Recordset("clmax") * coeflap
            Form4.clmaxfinal = Format(Form4.clmaxfinal, "0.00")
End If

'PARA 20 GRAUS
If VSflap.Value = 3 Then

flap.x1 = 1800 + kflap
flap.y1 = 480
flap.X2 = 2160 + kflap
flap.y2 = 600

deflap = 20
    If semflap.Value = True Then coeflap = 1
    If plainflap.Value = True Then coeflap = 1 + Abs(0.0075 * deflap)
    If splitflap.Value = True Then coeflap = 1 + Abs(0.01 * deflap)
    If fowlerflap.Value = True Then coeflap = 1 + Abs(0.015 * deflap)

            clmaxfinal = Form4.Data1.Recordset("clmax") * coeflap
            Form4.clmaxfinal = Format(Form4.clmaxfinal, "0.00")
End If

'PARA 30 GRAUS
If VSflap.Value = 4 Then

flap.x1 = 1800 + kflap
flap.y1 = 480
flap.X2 = 2160 + kflap
flap.y2 = 660

deflap = 30
    If semflap.Value = True Then coeflap = 1
    If plainflap.Value = True Then coeflap = 1 + Abs(0.0075 * deflap)
    If splitflap.Value = True Then coeflap = 1 + Abs(0.01 * deflap)
    If fowlerflap.Value = True Then coeflap = 1 + Abs(0.015 * deflap)

            clmaxfinal = Form4.Data1.Recordset("clmax") * coeflap
            Form4.clmaxfinal = Format(Form4.clmaxfinal, "0.00")
End If

'PARA 45 GRAUS
If VSflap.Value = 5 Then

flap.x1 = 1800 + kflap
flap.y1 = 480
flap.X2 = 2040 + kflap
flap.y2 = 720

deflap = 45
    If semflap.Value = True Then coeflap = 1
    If plainflap.Value = True Then coeflap = 1 + Abs(0.0075 * deflap)
    If splitflap.Value = True Then coeflap = 1 + Abs(0.01 * deflap)
    If fowlerflap.Value = True Then coeflap = 1 + Abs(0.015 * deflap)

            clmaxfinal = Form4.Data1.Recordset("clmax") * coeflap
            Form4.clmaxfinal = Format(Form4.clmaxfinal, "0.00")
End If

Form4.valflap = Form4.VSflap.Value

If deflap = 0 Then semflap.Value = True
    
End Sub
