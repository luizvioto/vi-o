VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Aerodinâmica"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
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
      Begin VB.TextBox perfvh 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox perfesth 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox perfasa 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox CMaerodht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         TabIndex        =   83
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox CMaerodw 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         TabIndex        =   82
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox rey3d 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox rey2d 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox slopeFLAP 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox cl3dFLAP 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox alfaFLAPo 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox alfaFLAPm 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox lddecol 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox cddecol 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox cldecol 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox wsproj 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox alfasamax 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox alfasazero 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox slopeasa 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox cdperfasa 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox cl3dasa 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox alfahtmax 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox alfahtzero 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox slopeht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox cdperfht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox cl3dht 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox ldk1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox ldk2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox ldk4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox swetasa 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox swetfuse 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox swethorz 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox swetvert 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox anplan 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox placaplan 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox clk1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox clk2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox clk4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox cdtotstall 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   5880
         Width           =   615
      End
      Begin VB.TextBox cdtotsubida 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   6840
         Width           =   615
      End
      Begin VB.TextBox cdtotvmax 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox cdperfvt 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil Estab. Vertical :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8040
         TabIndex        =   90
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil Estab. Horizontal :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   88
         Top             =   720
         Width           =   1680
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perfil da asa"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   86
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda média aerodinâmica :                           m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   84
         Top             =   3120
         Width           =   3300
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corda média aerodinâmica :                         m"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   81
         Top             =   3120
         Width           =   3210
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3D~ :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9720
         TabIndex        =   79
         Top             =   4080
         Width           =   405
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2D :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8280
         TabIndex        =   77
         Top             =   4080
         Width           =   300
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Nº Reynolds"
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
         Left            =   8280
         TabIndex        =   76
         Top             =   3480
         Width           =   3165
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   75
         Top             =   1080
         Width           =   150
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   74
         Top             =   1560
         Width           =   150
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   360
         TabIndex        =   71
         Top             =   1560
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   360
         TabIndex        =   69
         Top             =   1080
         Width           =   150
      End
      Begin VB.Label label78 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eficiencia aerod. L/D Vdecol :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8280
         TabIndex        =   66
         Top             =   6360
         Width           =   2145
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente de arrasto - Vdecol :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   64
         Top             =   6360
         Width           =   2265
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente CL para Vdecolag. :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   62
         Top             =   6360
         Width           =   2250
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   7920
         X2              =   7920
         Y1              =   600
         Y2              =   7560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         X1              =   4080
         X2              =   4080
         Y1              =   600
         Y2              =   7560
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Eficiencia aerodinâmica"
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
         Left            =   8280
         TabIndex        =   60
         Top             =   4800
         Width           =   3225
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Arrasto - CD"
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
         Left            =   4320
         TabIndex        =   59
         Top             =   4800
         Width           =   3165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Clmáx. s/flap:                c/flap:                 º"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   58
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cl 3D       s/flap:                c/flap:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   57
         Top             =   2160
         Width           =   2340
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Clzero  s/flap:                c/flap:                 º"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   56
         Top             =   1680
         Width           =   3030
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slope (a)  s/flap:                c/flap:                /rad"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   55
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente de arrasto do perfil :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   54
         Top             =   3600
         Width           =   2235
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Asa"
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
         Left            =   240
         TabIndex        =   53
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slope - curva de sustentação :                      /rad"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   52
         Top             =   2640
         Width           =   3465
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cl zero do perfil :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4560
         TabIndex        =   51
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente de sustentação 3D :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   50
         Top             =   2160
         Width           =   2280
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cl máximo do perfil :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4560
         TabIndex        =   49
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente de arrasto do perfil :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   48
         Top             =   3600
         Width           =   2235
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Estabilizador Horizontal"
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
         Left            =   4320
         TabIndex        =   47
         Top             =   240
         Width           =   3165
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente CL para Vestol :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   46
         Top             =   5880
         Width           =   1965
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente CL para Vplaneio :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   45
         Top             =   6840
         Width           =   2145
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente CL para Vmáxima :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   44
         Top             =   7320
         Width           =   2160
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área molhada da Asa :                                 m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   43
         Top             =   4080
         Width           =   3300
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área molhada da  Fuselagem :                  m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8280
         TabIndex        =   42
         Top             =   2880
         Width           =   3165
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área molhada da Emp. Horiz. :                       m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   41
         Top             =   4080
         Width           =   3390
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Área molhada da Emp. Vert. :                    m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8280
         TabIndex        =   40
         Top             =   1680
         Width           =   3180
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Melhor angulo de planeio :                             º"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8280
         TabIndex        =   39
         Top             =   5400
         Width           =   3240
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fator de Placa Plana  em Vmax :                    m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   38
         Top             =   5400
         Width           =   3420
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente de arrasto - Vestol :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   37
         Top             =   5880
         Width           =   2205
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente de arrasto - Vplaneio :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   36
         Top             =   6840
         Width           =   2385
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente de arrasto - Vmáxima :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   35
         Top             =   7320
         Width           =   2400
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eficiencia aerod. L/D Vestol :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8280
         TabIndex        =   34
         Top             =   5880
         Width           =   2085
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eficiencia aerod. L/D Vplaneio :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8280
         TabIndex        =   33
         Top             =   6840
         Width           =   2265
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eficiencia aerod.  L/D Vmáxima :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8280
         TabIndex        =   32
         Top             =   7320
         Width           =   2325
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carga alar de projeto :                                  kgf/m2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   5400
         Width           =   3600
      End
      Begin VB.Label Label16 
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
         Left            =   8280
         TabIndex        =   30
         Top             =   2280
         Width           =   3165
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Estabilizador Vertical"
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
         Left            =   8280
         TabIndex        =   29
         Top             =   240
         Width           =   3195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coeficiente de arrasto do perfil :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   8280
         TabIndex        =   28
         Top             =   1200
         Width           =   2235
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Sustentação - CL"
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
         TabIndex        =   27
         Top             =   4800
         Width           =   3045
      End
   End
   Begin VB.Menu fileirf7 
      Caption         =   "&Ir para ..."
      Begin VB.Menu fileperform 
         Caption         =   "Performance"
      End
      Begin VB.Menu filedimension 
         Caption         =   "Dimensionamento"
      End
      Begin VB.Menu fileprincipia 
         Caption         =   "Página inicial"
      End
      Begin VB.Menu filedesenho7 
         Caption         =   "Retornar ao desenho"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu filevisual 
      Caption         =   "&Visualizar"
   End
   Begin VB.Menu filefer7 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "&Calculadora && Conversor"
      End
   End
   Begin VB.Menu fileimprimi7 
      Caption         =   "Im&primir"
   End
   Begin VB.Menu sai7 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub filecalcon_Click()
Form29.Show
End Sub

Private Sub filedesenho7_Click()
Form18.Visible = True
Form7.Visible = False
Form18.Refresh
End Sub

Private Sub filedimension_Click()
Form7.Visible = False
Form2.Visible = True
End Sub

Private Sub fileimprimi7_Click()
If LIBER <> "KFAR0852" Then
    MsgBox "A rotina de impressão sómente estará disponível após a liberação permanente", 16, "Cancelamento de impressão": Exit Sub
End If

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
Printer.Print Tab(20); "Angulo alfa para Clo.s/flap:"; Tab(70); alfasazero;
Printer.Print Tab(20); "Angulo alfa para Clo.c/flap:"; Tab(70); Form7.alfaFLAPo;
Printer.Print Tab(20); "Coeficiente de arrasto do perfil:"; Tab(70); cdperfasa;
Printer.Print Tab(20); "Coeficiente Cl3D.s/flap:"; Tab(70); cl3dasa;
Printer.Print Tab(20); "Coeficiente Cl3D.c/flap:"; Tab(70); Form7.cl3dFLAP;
Printer.Print Tab(20); "dCl/dalfa - s/flap:"; Tab(70); slopeasa; " /rad"
Printer.Print Tab(20); "dCl/dalfa - c/flap:"; Tab(70); Form7.slopeFLAP; " /rad"
Printer.Print Tab(20); "Área molhada da asa:"; Tab(70); swetasa; " m2"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "2 - ESTABILIZADOR HORIZONTAL"
Printer.FontSize = 10
Printer.Print Tab(20); "Angulo alfa para Clmáximo:"; Tab(70); alfahtmax;
Printer.Print Tab(20); "Angulo alfa para Clzero:"; Tab(70); alfahtzero;
Printer.Print Tab(20); "Coeficiente de arrasto do perfil:"; Tab(70); cdperfht;
Printer.Print Tab(20); "Coeficiente de sustentação 3D:"; Tab(70); cl3dht;
Printer.Print Tab(20); "Inclinação da curva de sustentação:"; Tab(70); slopeht; " /rad"
Printer.Print Tab(20); "Área molhada do Estabilizador horizontal:"; Tab(70); swethorz; " m2"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "3 - ESTABILIZADOR VERTICAL"
Printer.FontSize = 10
Printer.Print Tab(20); "Coeficiente de arrasto do perfil:"; Tab(70); cdperfvt;
Printer.Print Tab(20); "Área molhada do Estabilizador Vertical:"; Tab(70); swetvert; " m2"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "4 - FUSELAGEM"
Printer.FontSize = 10
Printer.Print Tab(20); "Área molhada da Fuselagem:"; Tab(70); swetfuse; " m2"
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print Tab(25); "5 - SUSTENTAÇÃO"
Printer.FontSize = 10
Printer.Print Tab(20); "Carga alar de projeto:"; Tab(70); wsproj; " kgf/m2"
Printer.Print Tab(20); "Cl em estol:"; Tab(70); clk1;
Printer.Print Tab(20); "Cl em decolagem:"; Tab(70); cldecol;
Printer.Print Tab(20); "Cl em subida:"; Tab(70); clk2;
Printer.Print Tab(20); "Cl em velocidade máxima:"; Tab(70); clk4;
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print ""
Printer.Print Tab(25); "6 - ARRASTO"
Printer.FontSize = 10
Printer.Print Tab(20); "Fator de placa plana:"; Tab(70); placaplan; " m2"
Printer.Print Tab(20); "Cdtotal em estol:"; Tab(70); cdtotstall;
Printer.Print Tab(20); "Cdtotal em decolagem:"; Tab(70); cddecol;
Printer.Print Tab(20); "Cdtotal em subida:"; Tab(70); cdtotsubida;
Printer.Print Tab(20); "Cdtotal em velocidade máxima:"; Tab(70); cdtotvmax;
Printer.Print ""
Printer.FontBold = True
Printer.FontSize = 12
Printer.Print ""
Printer.Print Tab(25); "7 - EFICIENCIA AERODINÂMICA"
Printer.FontSize = 10
Printer.Print Tab(20); "Melhor angulo de planeio:"; Tab(70); anplan; " º"
Printer.Print Tab(20); "L/D em estol:"; Tab(70); ldk1;
Printer.Print Tab(20); "L/D em decolagem:"; Tab(70); lddecol;
Printer.Print Tab(20); "L/D em subida:"; Tab(70); ldk2;
Printer.Print Tab(20); "L/D em velocidade máxima:"; Tab(70); ldk4;
Printer.Print ""
Printer.Print Tab(20); "Fim da página";
Printer.EndDoc

End Sub

Private Sub fileperform_Click()


Form8.Visible = True
Form7.Visible = False

'..................ABERTURA FORM9...........................
                        'Performance
    'Calculo com Motor
Form8.mtr = Requisitos.motorBD

Form8.pesomot = 1 * Form1.motor

Form8.rpmf1 = Form1.RPM

'COLOCAR Banco Dados >>>>Requisitos.wfR = 200

Form8.fuel = (Requisitos.wfR / 2.2) * 0.725
Form8.fuel = Format(Form8.fuel, "0.00")

Form8.vstol = Form1.vstol
Form8.vstol = Format(Form8.vstol, "0.00")

Form8.NI = Requisitos.NIcruz * 100

torque = Form1.motor * 5252 / Form1.RPM
torque = Format(torque, "0.00")

'cancelada em 10/02/04 Form8.hpreq = (torque * Form1.RPM) / 5252
'Form8.hpreq = Format(Form8.hpreq, "0.00")

Form8.hpnominal = Form1.motor
Form8.hpnominal = Format(Form8.hpnominal, "0.00")

'Potencia do motor com eficiencia
hpreal = Form8.hpnominal * Requisitos.NIcruz
Form8.hpreal = hpreal
Form8.hpreal = Format(Form8.hpreal, "0.00")

vclimb = ((19.77 * ((Form1.wo * 2.2 / (Form7.clk2 * Form2.sw * 10.76)) ^ 0.5))) * 1.61

kclimb = (hpreal * 375) / (Form8.vmax / 1.61) * (vclimb / 1.61)
kcclimb = ((wo * 2.2) / Form7.ldk2) * (vclimb / 1.61)

climb = (kclimb - kcclimb) * 88 / (Form1.wo * 2.2)
Form8.climb = climb
Form8.climb = Format(Form8.climb, "0")

angclimb = 0.65 * (Form8.climb / vclimb / 1.61)
Form8.angclimb = angclimb
Form8.angclimb = Format(Form8.angclimb, "0")

Form8.rpmf1 = Form1.RPM


'Sem ter hélice
 Vverdade = Form8.vmax
 
If Form1.pas = 2 Then kpas = 20
If Form1.pas = 3 Then kpas = 18
Form8.npas = Form1.pas

Form8.diahel = Int(kpas * ((Form8.hpnominal) ^ 0.25))
Form8.diahel = Format(Form8.diahel, "0.00")
    
avahel = (Vverdade / 1.61) / Form1.RPM * Form8.diahel
Form8.avahel = avahel
Form8.avahel = Format(Form8.avahel, "0.00")

pashel1 = (Vverdade / 1.61) / (Form1.RPM / 1056)

    
    If pashel <= Tpashel Then
        Form8.pashel = Int(pashel1)
        Form8.pashel = Format(Form8.pashel, "0.00")
    End If
    
    
Form8.cahel = Form8.diahel + " x " + Form8.pashel

clotm = Form7.clk2
velplano = 19.77 * (((Form1.wo * 2.2) / (clotm * Form2.sw * 10.76)) ^ 0.5)
Form8.velplan = velplano * 1.61
Form8.velplan = Format(Form8.velplan, "0.00")

Form8.hpreal = Form8.hpnominal * Requisitos.NIcruz
Form8.hpreal = Format(Form8.hpreal, "0.00")

Form8.potreq = ((Form8.FD * 2.2) * Form8.vmax / 1.61) / 375
Form8.potreq = Format(Form8.potreq, "0.00")

Form8.potreqstol = ((Form8.FDstol * 2.2) * Form8.vstol / 1.61) / 375
Form8.potreqstol = Format(Form8.potreqstol, "0.00")

Form8.potreqsub = ((Form8.Fdsubida * 2.2) * Form8.vsubida / 1.61) / 375
Form8.potreqsub = Format(Form8.potreqsub, "0.00")

Form8.potreqdec = ((Form8.FDdecol * 2.2) * Form8.vdecol / 1.61) / 375
Form8.potreqdec = Format(Form8.potreqdec, "0.00")

Form8.sobrahp = Abs(Form8.potreq) - Abs(Form8.potreqstol)
Form8.sobrahp = Format(Form8.sobrahp, "0.00")

toque1 = Form1.vstol * 0.95
Form8.toque = toque1
Form8.toque = Format(Form8.toque, "0.00")

dar = Form1.roar * 0.00194
 
vstk2 = ((Form1.wo * 2.2 * 2) / (dar * Form7.cl3dasa * Form2.sw * 10.76)) ^ 0.5
vst2 = vstk2 * 1.09728
vst2 = Format(vst2, "0.00")

toque2 = vst2 * 0.95
Form8.toque2 = toque2
Form8.toque2 = Format(Form8.toque2, "0.00")

'................. 'Distancia de Decolagem ....................
Dg = Form1.wo * 2.2 * (Form7.cddecol / Form7.cldecol)
Rs = 0.02 * Form1.wo * 2.2

Tsta = (Form8.hpnominal * Requisitos.NIcruz * 550) / (Form8.vmax * 0.911)
Form8.tracestatica = Tsta / 2.2
Form8.tracestatica = Format(Form8.tracestatica, "0.00")

Fs = Abs(Tsta) - Abs(Rs)
If Fs <= 0 Then
Form8.Ldecol = 0
Form8.Ldecol = Format(Form8.Ldecol, "0")
Form8.Lpouso = 0
Form8.Lpouso = Format(Form8.Lpouso, "0")
Form8.lpousf = 0
Form8.lpousf = Format(Form8.lpousf, "0")
Form8.ldecsf = 0
Form8.ldecsf = Format(Form8.ldecsf, "0")
GoTo 2000
End If

If Fs <> 0 Then GoTo 1500

1500:
Fg = Abs(Tsta) - Abs(Dg)

' DECOLAGEM E POUSO COM FLAP
Tto = ((50000 * Form1.motor) / (Form1.RPM * Form8.diahel * 0.08333))
Ta = (375 * Form8.potreqdec) / (Form8.vdecol * 0.6213)

Fo = Abs(Tto) - Abs(0.02 * Form1.wo * 2.2)
Ft = Abs(Ta) - Abs((Form1.wo * 2.2) * (1 / Form7.lddecol))

Kdec = (Abs(Tto) - Abs(Ta)) / Tto

Lsobbst = 0.011 * (((Form1.wo * 2.2) * (Form8.vdecol * 0.621) ^ 2) / Ft)
Ldecol = ((((Form8.vdecol * 0.911) ^ 2) * Form1.wo * 2.2) / (32.2 * Fo)) * ((1 / Kdec) * (-1 - (1 / Kdec * (Log(1 - Kdec))))) + Lsobbst
Form8.Ldecol = (Ldecol * 0.305)
Form8.Ldecol = Format(Form8.Ldecol, "0")

vcontacto = toque1 * 0.911
Lpouso = (Form1.wo * 2.2 / 64.4) * ((vcontacto ^ 2) / (0.8 * Fs))
Form8.Lpouso = (Lpouso * 0.305)
Form8.Lpouso = Format(Form8.Lpouso, "0")

' DECOLAGEM E POUSO SEM FLAP


' STALL SEM FLAP .........

dar = Form1.roar * 0.00194

vstk2 = ((Form1.wo * 2.2 * 2) / (dar * Form7.cl3dasa * Form2.sw * 10.76)) ^ 0.5
Form8.vst2 = vstk2 * 1.09728
Form8.vst2 = Format(Form8.vst2, "0.00")

'Velocidade de Decolagem s/flap
Form8.vdecolsf = 1.1 * Form8.vst2
Form8.vdecolsf = Format(Form8.vdecolsf, "0.00")

Lsosf = 0.011 * (((Form1.wo * 2.2) * (Form8.vdecolsf * 0.621) ^ 2) / Ft)
ldecsf = ((((Form8.vdecolsf * 0.911) ^ 2) * Form1.wo * 2.2) / (32.2 * Fo)) * ((1 / Kdec) * (-1 - (1 / Kdec * (Log(1 - Kdec))))) + Lsosf
Form8.ldecsf = 0.305 * ldecsf
Form8.ldecsf = Format(Form8.ldecsf, "0")

vcontacto2 = toque2 * 0.911
lpousf = (Form1.wo * 2.2 / 64.4) * ((vcontacto2 ^ 2) / (0.8 * Fs))
Form8.lpousf = (lpousf * 0.305)
Form8.lpousf = Format(Form8.lpousf, "0")




2000:
tradinamica = Tsta * (1# - (0.106 * (Form8.vmax * 0.911) * (((Form1.RPM * ((Form8.diahel * 0.08333) ^ 3)) / (Form8.rpmf1 * (10 ^ 7))) ^ 0.5)))
Form8.tradin = tradinamica / 2.2
Form8.tradin = Format(Form8.tradin, "0.00")

Form8.TD = Form1.wo / Form8.FD
Form8.TD = Format(Form8.TD, "0.00")

Form8.ldtw = (Form7.ldk4 / Form8.TD)
Form8.ldtw = Format(Form8.ldtw, "0.0")

'.................. OUTROS ..........................

If Form8.vmax = Form8.vstol Then
MsgBox "A potencia do motor é insuficiente. Baixo Nº de RPMs.", vbCritical
Form8.vmax.ForeColor = &HFF&
Form8.velplan.ForeColor = &HFF&
Form8.vsubida.ForeColor = &HFF&
Form8.Ldecol.ForeColor = &HFF&
Form8.Lpouso.ForeColor = &HFF&
Form8.lpousf.ForeColor = &HFF&
Form8.ldecsf.ForeColor = &HFF&
Form8.climb.ForeColor = &HFF&
Form8.climb = 0
Form8.climb = Format(Form8.climb, "0.00")
End If

If Form8.vmax <> Form8.vstol Then
Form8.vmax.ForeColor = &HFF0000
Form8.vdecol.ForeColor = &HFF0000
Form8.vstol.ForeColor = &HFF0000
Form8.velplan.ForeColor = &HFF0000
Form8.vsubida.ForeColor = &HFF0000
Form8.Ldecol.ForeColor = &HFF0000
Form8.Lpouso.ForeColor = &HFF0000
Form8.lpousf.ForeColor = &HFF0000
Form8.ldecsf.ForeColor = &HFF0000
Form8.climb.ForeColor = &HFF0000
End If


'RESUMO PARA SEM FLAP
If Form4.semflap.Value = True Then
    Form8.vst2 = Form8.vstol
    Form8.vdecolsf = Form8.vdecol
    Form8.toque2 = Form8.toque
    
    Form8.vstol = Form8.vst2
    Form8.vdecol = Form8.vdecolsf
    Form8.toque = Form8.toque2
    
    Form8.Ldecol = Form8.ldecsf
    Form8.Lpouso = Form8.lpousf
    
    Form8.Label46.Caption = "s/flap:"
    Form8.Label47.Caption = "s/flap:"
    Form8.Label48.Caption = "s/flap:"
    Form8.Label46.ForeColor = &HFFFF&
    Form8.Label47.ForeColor = &HFFFF&
    Form8.Label48.ForeColor = &HFFFF&
End If

If Form4.semflap.Value = False Then
    Form8.Label46.Caption = "c/flap:"
    Form8.Label47.Caption = "c/flap:"
    Form8.Label48.Caption = "c/flap:"
    Form8.Label46.ForeColor = &HFFFF00
    Form8.Label47.ForeColor = &HFFFF00
    Form8.Label48.ForeColor = &HFFFF00
End If


'VELOCIDADE de CRUZEIRO
Form8.vcruzeirook = ((0.75 * (Form8.hpnominal * Requisitos.NIcruz) * 550) / Tsta) * 1.0972
Form8.vcruzeirook = Format(Form8.vcruzeirook, "0.00")

'CAFE Challenge
Form8.cafechal = ((Form8.vmax / 1.61) ^ 1.3) * ((Requisitos.alcR) / (Requisitos.wfR / 6)) * (Requisitos.CPR ^ 0.6) / 10000
Form8.cafechal = Format(Form8.cafechal, "0.00")
'CAFE Trviathon
Form8.cafetria = (28110625 * ((Form8.vmax / 1.61) * Form8.climb) ^ 2) / ((4100625 + ((Form8.vstol / 1.61) ^ 4)) * 10 ^ 9)
Form8.cafetria = Format(Form8.cafetria, "0.00")

Form7.Refresh
Form8.Refresh

End Sub

Private Sub fileprincipia_Click()
Form1.Visible = True
Form7.Visible = False

Form1.Refresh
End Sub

Private Sub filevisual_Click()
Form9.Show
'Enviando vminima = vstol para form9
Form9.velominima = Form1.vstol
Form9.velominima = Format(Form9.velominima, "0.00")

Form9.vinicio = Form9.velominima
Form9.vinicio = Format(Form9.vinicio, "0.00")
'Form9.velominima = Format(Form9.velominima, "0.00")
Form9.Refresh
End Sub


Private Sub sai7_Click()
Form1.Visible = True
Form7.Visible = False
Form1.Refresh
End Sub
