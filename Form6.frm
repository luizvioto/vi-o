VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Motorização"
   ClientHeight    =   7935
   ClientLeft      =   150
   ClientTop       =   615
   ClientWidth     =   7095
   ControlBox      =   0   'False
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&ESTUDOS DE PERFORMANCE  - Hélice existente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3375
         Left            =   240
         TabIndex        =   2
         Top             =   4320
         Width           =   6615
         Begin VB.OptionButton temhel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "O usuário já possui a hélice para o projeto"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   360
            Width           =   3735
         End
         Begin VB.PictureBox Picture4 
            Height          =   375
            Left            =   4320
            ScaleHeight     =   315
            ScaleWidth      =   1875
            TabIndex        =   62
            Top             =   360
            Width           =   1935
            Begin VB.CommandButton Command6 
               Caption         =   "Hélice existente"
               Enabled         =   0   'False
               Height          =   315
               Left            =   0
               TabIndex        =   63
               Top             =   0
               Width           =   1875
            End
         End
         Begin VB.TextBox NIporc 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   1320
            Width           =   615
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   3120
            Max             =   10
            TabIndex        =   59
            Top             =   3000
            Width           =   855
         End
         Begin VB.TextBox altitude 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2040
            TabIndex        =   58
            Text            =   "0,00"
            Top             =   3000
            Width           =   615
         End
         Begin VB.TextBox densidade 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5040
            TabIndex        =   56
            Top             =   3000
            Width           =   855
         End
         Begin VB.TextBox NI 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox rpm19 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5400
            TabIndex        =   32
            Top             =   1680
            Width           =   615
         End
         Begin VB.HScrollBar hsrpm 
            Height          =   255
            Left            =   4320
            Max             =   30000
            Min             =   100
            SmallChange     =   100
            TabIndex        =   31
            Top             =   1680
            Value           =   2500
            Width           =   975
         End
         Begin VB.HScrollBar hsNi 
            Height          =   255
            Left            =   4320
            Max             =   100
            Min             =   30
            SmallChange     =   5
            TabIndex        =   30
            Top             =   1320
            Value           =   30
            Width           =   975
         End
         Begin VB.TextBox hp19 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox Vestimada 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "0,00"
            Top             =   840
            Width           =   615
         End
         Begin VB.CheckBox optavest 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pré estabelecendo a Velocidade Máxima de projeto ...................."
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   4935
         End
         Begin VB.TextBox aval6 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox torque 
            Height          =   285
            Left            =   5400
            TabIndex        =   54
            Top             =   2160
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox hs1salva 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   60
            Top             =   3000
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dens. do ar :                      kg/m3"
            Height          =   195
            Left            =   4080
            TabIndex        =   57
            Top             =   3000
            Width           =   2370
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Altitiude máxima de vôo :                mts."
            Height          =   195
            Left            =   240
            TabIndex        =   55
            Top             =   3000
            Width           =   2760
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Variar RPM p/ estudar performance :"
            ForeColor       =   &H80000015&
            Height          =   195
            Left            =   240
            TabIndex        =   37
            Top             =   1680
            Width           =   2610
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rpm"
            ForeColor       =   &H80000015&
            Height          =   195
            Left            =   6120
            TabIndex        =   36
            Top             =   1680
            Width           =   330
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Variar a Eficiencia da hélice  p/ estudar performance :"
            ForeColor       =   &H80000015&
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   1320
            Width           =   3810
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H80000015&
            Height          =   195
            Left            =   6120
            TabIndex        =   34
            Top             =   1320
            Width           =   120
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H80000015&
            Height          =   195
            Left            =   6120
            TabIndex        =   24
            Top             =   2520
            Width           =   120
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Hp."
            ForeColor       =   &H80000015&
            Height          =   195
            Left            =   6120
            TabIndex        =   20
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Potencia Nominal calculada em função da variação de Rpm  (HPc) :"
            ForeColor       =   &H80000015&
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   2160
            Width           =   4845
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Km/h."
            ForeColor       =   &H80000015&
            Height          =   195
            Left            =   6120
            TabIndex        =   7
            Top             =   840
            Width           =   435
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Relação entre Potencia Nominal original e Potencia Nominal calculada :"
            ForeColor       =   &H80000015&
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   2520
            Width           =   5085
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&REQUERIMENTOS OBRIGATÓRIOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3975
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   6615
         Begin VB.PictureBox Picture1 
            Height          =   615
            Left            =   4320
            ScaleHeight     =   555
            ScaleWidth      =   1995
            TabIndex        =   40
            Top             =   3240
            Width           =   2055
            Begin VB.CommandButton Command2 
               Caption         =   "&Enviar dados"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   0
               MouseIcon       =   "Form6.frx":030A
               MousePointer    =   99  'Custom
               TabIndex        =   41
               Top             =   0
               Width           =   1995
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Left            =   4320
            TabIndex        =   38
            Top             =   2280
            Width           =   2055
            Begin VB.PictureBox Picture3 
               Height          =   495
               Left            =   120
               ScaleHeight     =   435
               ScaleWidth      =   1755
               TabIndex        =   42
               Top             =   240
               Width           =   1815
               Begin VB.CommandButton Command3 
                  Caption         =   "&Hélices"
                  Height          =   435
                  Left            =   0
                  MouseIcon       =   "Form6.frx":0614
                  MousePointer    =   99  'Custom
                  TabIndex        =   43
                  ToolTipText     =   "Escolha a hélice por missão"
                  Top             =   0
                  Width           =   1755
               End
            End
            Begin VB.ListBox listhel 
               Height          =   255
               ItemData        =   "Form6.frx":091E
               Left            =   120
               List            =   "Form6.frx":0934
               TabIndex        =   39
               Top             =   240
               Visible         =   0   'False
               Width           =   1815
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Material da hélice"
            ForeColor       =   &H00C00000&
            Height          =   735
            Left            =   240
            TabIndex        =   25
            Top             =   3120
            Width           =   3975
            Begin VB.OptionButton Ophelfib 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Hélice em Fibra"
               Height          =   255
               Left            =   2160
               TabIndex        =   27
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton Ophelmad 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Hélice em Madeira"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Material de construção do avião"
            ForeColor       =   &H00C00000&
            Height          =   735
            Left            =   240
            TabIndex        =   16
            Top             =   2280
            Width           =   3975
            Begin VB.OptionButton opfibra 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "Fibra (compostos) "
               Height          =   255
               Left            =   2160
               TabIndex        =   18
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton opmadeira 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Madeira (ou) e Tela"
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0E0&
            Height          =   1935
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   6135
            Begin VB.Frame combustivel 
               BackColor       =   &H00E0E0E0&
               Height          =   1515
               Left            =   1320
               TabIndex        =   47
               Top             =   240
               Visible         =   0   'False
               Width           =   4695
               Begin VB.ListBox tanque 
                  BackColor       =   &H00C0FFFF&
                  ForeColor       =   &H00FF0000&
                  Height          =   450
                  ItemData        =   "Form6.frx":099E
                  Left            =   120
                  List            =   "Form6.frx":09BD
                  TabIndex        =   51
                  Top             =   240
                  Width           =   4455
               End
               Begin VB.TextBox wfuel 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   1800
                  Locked          =   -1  'True
                  TabIndex        =   50
                  Text            =   "0.00"
                  Top             =   960
                  Width           =   615
               End
               Begin VB.PictureBox Picture2 
                  Height          =   375
                  Left            =   3000
                  ScaleHeight     =   315
                  ScaleWidth      =   1515
                  TabIndex        =   48
                  Top             =   960
                  Width           =   1575
                  Begin VB.CommandButton Command5 
                     Caption         =   "&Ok !"
                     Height          =   315
                     Left            =   0
                     MouseIcon       =   "Form6.frx":0BB5
                     MousePointer    =   99  'Custom
                     TabIndex        =   49
                     Top             =   0
                     Width           =   1515
                  End
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Peso do combustível :"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   53
                  Top             =   960
                  Width           =   1590
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "grs"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   52
                  Top             =   960
                  Width           =   210
               End
            End
            Begin VB.TextBox pesomotor 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   1440
               Width           =   735
            End
            Begin VB.CommandButton Command4 
               Height          =   615
               Left            =   240
               MouseIcon       =   "Form6.frx":0EBF
               MousePointer    =   99  'Custom
               Picture         =   "Form6.frx":11C9
               Style           =   1  'Graphical
               TabIndex        =   29
               ToolTipText     =   "Escolha o reservatório de combustível"
               Top             =   1080
               Width           =   975
            End
            Begin VB.CommandButton Command1 
               Height          =   615
               Left            =   240
               MouseIcon       =   "Form6.frx":160B
               MousePointer    =   99  'Custom
               Picture         =   "Form6.frx":1915
               Style           =   1  'Graphical
               TabIndex        =   28
               ToolTipText     =   "Escolha o motor"
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox rpmoriginal 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   1080
               Width           =   735
            End
            Begin VB.TextBox hporiginal 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox cc 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "grs."
               Height          =   195
               Left            =   5520
               TabIndex        =   46
               Top             =   1440
               Width           =   270
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso estimado do motor selecionado:"
               Height          =   195
               Left            =   1560
               TabIndex        =   44
               Top             =   1440
               Width           =   2640
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rpm"
               Height          =   195
               Left            =   5520
               TabIndex        =   15
               Top             =   1080
               Width           =   330
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Hp."
               Height          =   255
               Left            =   5520
               TabIndex        =   14
               Top             =   720
               Width           =   255
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cu.In"
               Height          =   195
               Left            =   5520
               TabIndex        =   13
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rpm original do motor :"
               Height          =   195
               Left            =   1560
               TabIndex        =   12
               Top             =   1080
               Width           =   1620
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Potencia Nominal - original (PNo) :"
               Height          =   195
               Left            =   1560
               TabIndex        =   11
               Top             =   720
               Width           =   2415
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Motor selecionado :"
               Height          =   195
               Left            =   1560
               TabIndex        =   10
               Top             =   360
               Width           =   1395
            End
         End
      End
   End
   Begin VB.Menu filemodifica6 
      Caption         =   "&Modificar"
   End
   Begin VB.Menu filefer6 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "&Calculadora && Conversor"
      End
   End
   Begin VB.Menu filesair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
