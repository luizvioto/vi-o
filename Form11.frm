VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Estrutura da asa"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
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
      Begin VB.TextBox materialM 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox sigcizal 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox alma 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   7200
         Width           =   975
      End
      Begin VB.TextBox mesatc 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   6600
         Width           =   975
      End
      Begin VB.TextBox largamax 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox sigcomp 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox sigtrac 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox mesacp 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox altmax 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox section 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox materialA 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox m0 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox m1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox m2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox m3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox m4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox m5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox m6 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox m7 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox m8 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   6000
         Width           =   615
      End
      Begin VB.TextBox m9 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   6600
         Width           =   615
      End
      Begin VB.TextBox m10 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   7200
         Width           =   615
      End
      Begin VB.TextBox v0 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox v10 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   7200
         Width           =   615
      End
      Begin VB.TextBox v9 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   6600
         Width           =   615
      End
      Begin VB.TextBox v8 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   6000
         Width           =   615
      End
      Begin VB.TextBox v7 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox v6 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox v5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox v4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox v3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox v2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox v1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox q0 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox q1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox q2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox q3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox q4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox q5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox q6 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox q7 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   5400
         Width           =   735
      End
      Begin VB.TextBox q8 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   6000
         Width           =   735
      End
      Begin VB.TextBox q9 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   6600
         Width           =   735
      End
      Begin VB.TextBox q10 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   7200
         Width           =   735
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alma em :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9000
         TabIndex        =   109
         Top             =   2400
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   " Estação  "
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
         TabIndex        =   107
         Top             =   360
         Width           =   2805
      End
      Begin VB.Label npecas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9450
         TabIndex        =   106
         Top             =   7200
         Width           =   45
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cz. adm :                     kgf/cm2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9120
         TabIndex        =   105
         Top             =   4200
         Width           =   2190
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s"
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
         Height          =   345
         Left            =   9000
         TabIndex        =   103
         Top             =   3480
         Width           =   165
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alma            :                         cm"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9000
         TabIndex        =   100
         Top             =   7200
         Width           =   2265
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mesa Inf :                              cm"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9000
         TabIndex        =   99
         Top             =   6600
         Width           =   2265
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s"
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
         Height          =   345
         Left            =   9000
         TabIndex        =   95
         Top             =   2880
         Width           =   165
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "t adm :                       kgf/cm2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9195
         TabIndex        =   94
         Top             =   3000
         Width           =   2115
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Largura :                                cm"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9000
         TabIndex        =   93
         Top             =   5400
         Width           =   2280
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "c adm :                        kgf/cm2"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9120
         TabIndex        =   92
         Top             =   3600
         Width           =   2205
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s"
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
         Height          =   345
         Left            =   9000
         TabIndex        =   91
         Top             =   4080
         Width           =   165
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mesa Sup :                            cm"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9000
         TabIndex        =   89
         Top             =   6000
         Width           =   2280
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Altura viga :                           cm"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9000
         TabIndex        =   87
         Top             =   4800
         Width           =   2265
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Secção :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9000
         TabIndex        =   85
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mesas em :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   9000
         TabIndex        =   83
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label cl10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         TabIndex        =   82
         Top             =   7200
         Width           =   60
      End
      Begin VB.Label cl9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         TabIndex        =   81
         Top             =   6600
         Width           =   60
      End
      Begin VB.Label cl8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         TabIndex        =   80
         Top             =   6000
         Width           =   60
      End
      Begin VB.Label cl7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         TabIndex        =   79
         Top             =   5400
         Width           =   60
      End
      Begin VB.Label cl6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         TabIndex        =   78
         Top             =   4800
         Width           =   60
      End
      Begin VB.Label cl5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         TabIndex        =   77
         Top             =   4200
         Width           =   60
      End
      Begin VB.Label cl4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         TabIndex        =   76
         Top             =   3600
         Width           =   60
      End
      Begin VB.Label cl3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         TabIndex        =   75
         Top             =   3000
         Width           =   60
      End
      Begin VB.Label cl2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         TabIndex        =   74
         Top             =   2400
         Width           =   60
      End
      Begin VB.Label cl1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         TabIndex        =   73
         Top             =   1800
         Width           =   60
      End
      Begin VB.Label cl0 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         TabIndex        =   72
         Top             =   1200
         Width           =   60
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Carga"
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
         Left            =   3600
         TabIndex        =   71
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "CL"
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
         Left            =   2880
         TabIndex        =   70
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Longarina"
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
         Left            =   8880
         TabIndex        =   69
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Cortante"
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
         Left            =   6960
         TabIndex        =   68
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Momento"
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
         Left            =   5040
         TabIndex        =   67
         Top             =   360
         Width           =   2235
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   8760
         X2              =   8760
         Y1              =   840
         Y2              =   7680
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   7320
         X2              =   7320
         Y1              =   840
         Y2              =   7680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         X1              =   5400
         X2              =   5400
         Y1              =   840
         Y2              =   7680
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kgf x m                                kg. "
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   66
         Top             =   7200
         Width           =   2220
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kgf x m                                kg. "
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   65
         Top             =   6600
         Width           =   2220
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kgf x m                               kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   64
         Top             =   6000
         Width           =   2130
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kgf x m                               kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   63
         Top             =   5400
         Width           =   2130
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kgf x m                               kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   62
         Top             =   4800
         Width           =   2130
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kgf x m                               kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   61
         Top             =   4200
         Width           =   2130
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kg fx m                               kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   60
         Top             =   3600
         Width           =   2130
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kgf x m                               kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   59
         Top             =   3000
         Width           =   2130
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kgf x m                               kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   58
         Top             =   2400
         Width           =   2130
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kgf x m                               kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   57
         Top             =   1800
         Width           =   2130
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kgf x m                               kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   56
         Top             =   1200
         Width           =   2130
      End
      Begin VB.Label e0 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2160
         TabIndex        =   33
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label e10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2160
         TabIndex        =   32
         Top             =   7200
         Width           =   75
      End
      Begin VB.Label e9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2160
         TabIndex        =   31
         Top             =   6600
         Width           =   75
      End
      Begin VB.Label e8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2160
         TabIndex        =   30
         Top             =   6000
         Width           =   75
      End
      Begin VB.Label e7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2160
         TabIndex        =   29
         Top             =   5400
         Width           =   75
      End
      Begin VB.Label e6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2160
         TabIndex        =   28
         Top             =   4800
         Width           =   75
      End
      Begin VB.Label e5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2160
         TabIndex        =   27
         Top             =   4200
         Width           =   75
      End
      Begin VB.Label e4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2160
         TabIndex        =   26
         Top             =   3600
         Width           =   75
      End
      Begin VB.Label e3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2160
         TabIndex        =   25
         Top             =   3000
         Width           =   75
      End
      Begin VB.Label e2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2160
         TabIndex        =   24
         Top             =   2400
         Width           =   75
      End
      Begin VB.Label e1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2160
         TabIndex        =   23
         Top             =   1800
         Width           =   75
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estação 11 - (1.0 x b/2) :            m.                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   7200
         Width           =   4770
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estação 10 - (0.9 x b/2) :            m.                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   6600
         Width           =   4770
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estação 9 - (0.8 x b/2) :              m.                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   6000
         Width           =   4770
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estação 8 - (0.7 x b/2) :              m.                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   5400
         Width           =   4770
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estação 7 - (0.6 x b/2) :              m.                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   4800
         Width           =   4770
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estação 6 - (0.5 x b/2) :              m.                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   4200
         Width           =   4770
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estação 5 - (0.4 x b/2) :              m.                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   3600
         Width           =   4770
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estação 4 - (0.3 x b/2) :              m.                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   3000
         Width           =   4770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estação 3 - (0.2 x b/2) :              m.                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   4770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estação 2 - (0.1 x b/2) :              m.                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   4770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estação 1 - (0.0 x b/2) :              m.                                              kgf"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   4770
      End
   End
   Begin VB.Menu fileir11 
      Caption         =   "&Ir para ..."
      Begin VB.Menu filestabestat 
         Caption         =   "Peso && Balanceamento"
      End
      Begin VB.Menu fileestab11 
         Caption         =   "Estabilidade"
      End
      Begin VB.Menu fileperform11 
         Caption         =   "Performance"
      End
      Begin VB.Menu fileaerodin11 
         Caption         =   "Aerodinâmica"
      End
      Begin VB.Menu filedimensiona11 
         Caption         =   "Dimensionamento"
      End
      Begin VB.Menu fileprincipia11 
         Caption         =   "Página inicial"
      End
      Begin VB.Menu filedesenho11 
         Caption         =   "Retornar ao desenho"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu fileparametros 
      Caption         =   "P&arametros"
   End
   Begin VB.Menu filevisual12 
      Caption         =   "&Visualizar"
   End
   Begin VB.Menu filefer11 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "&Calculadora && Conversor"
      End
   End
   Begin VB.Menu fileimprimi11 
      Caption         =   "&Imprimir"
   End
   Begin VB.Menu sai11 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fileaerodin11_Click()
Form7.Visible = True
Form11.Visible = False
End Sub

Private Sub filecalcon_Click()
Form29.Show
End Sub

Private Sub filedesenho11_Click()
Form18.Visible = True
Form11.Visible = False
Form18.Refresh
End Sub

Private Sub filedimensiona11_Click()
Form12.Visible = True
Form11.Visible = False
End Sub

Private Sub fileestab11_Click()
Form11.Visible = False
Form10.Visible = True
End Sub

Private Sub fileimprimi11_Click()
If LIBER <> "KFAR0852" Then
    MsgBox "A rotina de impressão sómente estará disponível após a liberação permanente", 16, "Cancelamento de impressão": Exit Sub
End If

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
Printer.Print Tab(20); "Estação 01:"; Tab(40); e0; " m"; Tab(55); cl0; Tab(70); q0; Tab(85); m0; Tab(100); v0
Printer.Print Tab(20); "Estação 02:"; Tab(40); e1; " m"; Tab(55); cl1; Tab(70); q1; Tab(85); m1; Tab(100); v1
Printer.Print Tab(20); "Estação 03:"; Tab(40); e2; " m"; Tab(55); cl2; Tab(70); q2; Tab(85); m2; Tab(100); v2
Printer.Print Tab(20); "Estação 04:"; Tab(40); e3; " m"; Tab(55); cl3; Tab(70); q3; Tab(85); m3; Tab(100); v3
Printer.Print Tab(20); "Estação 05:"; Tab(40); e4; " m"; Tab(55); cl4; Tab(70); q4; Tab(85); m4; Tab(100); v4
Printer.Print Tab(20); "Estação 06:"; Tab(40); e5; " m"; Tab(55); cl5; Tab(70); q5; Tab(85); m5; Tab(100); v5
Printer.Print Tab(20); "Estação 07:"; Tab(40); e6; " m"; Tab(55); cl6; Tab(70); q6; Tab(85); m6; Tab(100); v6
Printer.Print Tab(20); "Estação 08:"; Tab(40); e7; " m"; Tab(55); cl7; Tab(70); q7; Tab(85); m7; Tab(100); v7
Printer.Print Tab(20); "Estação 09:"; Tab(40); e8; " m"; Tab(55); cl8; Tab(70); q8; Tab(85); m8; Tab(100); v8
Printer.Print Tab(20); "Estação 10:"; Tab(40); e9; " m"; Tab(55); cl9; Tab(70); q9; Tab(85); m9; Tab(100); v9
Printer.Print Tab(20); "Estação 11:"; Tab(40); e10; " m"; Tab(55); cl10; Tab(70); q10; Tab(85); m10; Tab(100); v10
Printer.Print ""
Printer.Print ""
Printer.Print Tab(25); "DIMENSIONAMENTO"
Printer.Print ""
Printer.Print Tab(20); "Secção transversal da longarina tipo:"; Tab(70); section;
Printer.Print Tab(20); "Mesas da longarina em madeira de:"; Tab(70); materialM;
Printer.Print Tab(20); "Alma da longarina em madeira de:"; Tab(70); materialA;
Printer.Print Tab(20); "Tensão admissível a tração - mesa:"; Tab(70); sigtrac; " kgf/cm2"
Printer.Print Tab(20); "Tensão admissível a compressão - mesa:"; Tab(70); sigcomp; " kgf/cm2"
Printer.Print Tab(20); "Tensão admissível ao cizalhamento:"; Tab(70); sigcizal; " %"
Printer.Print Tab(20); "Altura da viga:"; Tab(70); altmax; " cm"
Printer.Print Tab(20); "Largura da viga:"; Tab(70); largamax; " cm"
Printer.Print Tab(20); "Mesa superior (L x H):"; Tab(70); mesacp; " cm"
Printer.Print Tab(20); "Mesa inferior (L x H):"; Tab(70); mesatc; " cm"
Printer.Print Tab(20); "Alma (e x H):"; Tab(70); alma; " cm"
Printer.Print Tab(20); "Nº de Gs positivos:"; Tab(70); Form13.gsatual;
Printer.Print Tab(20); "Nº de Gs negativos:"; Tab(70); Form13.gsnegatual;
Printer.Print ""

Printer.Print ""
Printer.Print ""
Printer.Print Tab(20); "Fim da página";
Printer.EndDoc

End Sub

Private Sub fileparametros_Click()
Form13.Show
Form13.Refresh
End Sub

Private Sub fileperform11_Click()
Form8.Visible = True
Form11.Visible = False
End Sub

Private Sub fileprincipia11_Click()
Form1.Visible = True
Form11.Visible = False

Form1.Refresh
End Sub

Private Sub filestabestat_Click()
If Form11.altmax = Empty Then

    MsgBox " A longarina deverá ser calculada", vbInformation
    Form13.Show
        Exit Sub
End If

If Form11.altmax Or Form11.largamax <> Empty Then
    Form14.Visible = True
    Form11.Visible = False
    If Form21.Data1.Recordset.RecordCount < 1 Then GoTo 10
        If Form21.Data1.Recordset("projeto") <> Form20.p1 Then
10:
            MsgBox " Entre com Pesos e Distancia para os equipamentos, distancia em centimetros", vbInformation
        End If
End If

End Sub

Private Sub filevisual12_Click()


If Form12.lambda <> Empty Then
Form12.Visible = True
Form13.Visible = False
End If

If Form1.tremconv.Value = True Then Form12.beq.Visible = False
If Form1.tremconv.Value = True Then Form12.eixobeq.Visible = False

If Form1.tremtric.Value = True Then Form12.beq.Visible = True
If Form1.tremtric.Value = True Then Form12.eixobeq.Visible = True

If Form1.posialta Then asa = 3360
If Form1.posimedia Then asa = 3790
If Form1.posibaixa Then asa = 3990

Form12.sp1.y1 = asa
Form12.sp1.y2 = asa

Form12.Sw = Form2.Sw
Form12.bw = Form2.bw
Form12.cr = Form2.crw
Form12.cp = Form2.cpw
Form12.lambda = Form1.Afilasa
Form12.Visible = True
End Sub


Private Sub sai11_Click()
Form1.Visible = True
Form11.Visible = False

Form1.Refresh
End Sub
