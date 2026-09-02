VERSION 5.00
Begin VB.Form Form19 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo do CG - servos"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form19.frx":0000
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2505
      ScaleWidth      =   7785
      TabIndex        =   98
      Top             =   1080
      Visible         =   0   'False
      Width           =   7815
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   7440
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   101
         Top             =   30
         Width           =   315
         Begin VB.CommandButton Command5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "x"
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
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "distancia CGservo a LR"
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   1920
         TabIndex        =   105
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.R"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   600
         TabIndex        =   104
         Top             =   2160
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   720
         X2              =   5040
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderStyle     =   2  'Dash
         X1              =   720
         X2              =   720
         Y1              =   480
         Y2              =   2040
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estimativa rudimentar do CG de um servo."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   600
         TabIndex        =   103
         Top             =   60
         Width           =   2955
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CG - vista frontal"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6240
         TabIndex        =   100
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Shape Shape8 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   6840
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CG - vista lateral"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4560
         TabIndex        =   99
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Shape Shape7 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4920
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   135
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   6600
         Top             =   630
         Width           =   615
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   6840
         Top             =   720
         Width           =   135
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4920
         Top             =   720
         Width           =   135
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   4680
         Top             =   630
         Width           =   615
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   6600
         Top             =   840
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   4560
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5535
      Left            =   6240
      TabIndex        =   7
      Top             =   360
      Width           =   1935
      Begin VB.TextBox cggiga 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0"
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox gv1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox gv2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   25
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox gv3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox gv4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox gv5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Text            =   "0"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox gv6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Text            =   "0"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox gv7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Text            =   "0"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox gv8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Text            =   "0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox gv9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Text            =   "0"
         Top             =   4560
         Width           =   615
      End
      Begin VB.OptionButton ts4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "GIGANTE"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 9"
         Height          =   195
         Left            =   120
         TabIndex        =   96
         Top             =   4560
         Width           =   780
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 8"
         Height          =   195
         Left            =   120
         TabIndex        =   95
         Top             =   4080
         Width           =   780
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 7"
         Height          =   195
         Left            =   120
         TabIndex        =   94
         Top             =   3600
         Width           =   780
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 6"
         Height          =   195
         Left            =   120
         TabIndex        =   93
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 5"
         Height          =   195
         Left            =   120
         TabIndex        =   92
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 4"
         Height          =   195
         Left            =   120
         TabIndex        =   91
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 3"
         Height          =   195
         Left            =   120
         TabIndex        =   90
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 2"
         Height          =   195
         Left            =   120
         TabIndex        =   89
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 1"
         Height          =   195
         Left            =   120
         TabIndex        =   88
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "centímetros"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   960
         TabIndex        =   60
         Top             =   4920
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CG grupo :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   5160
         Width           =   765
      End
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H00FF0000&
      Height          =   5535
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   1935
      Begin VB.TextBox cgretrat 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0"
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox rv1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox rv2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox rv3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox rv4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox rv5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Text            =   "0"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox rv6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Text            =   "0"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox rv7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Text            =   "0"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox rv8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Text            =   "0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox rv9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "0"
         Top             =   4560
         Width           =   615
      End
      Begin VB.OptionButton ts3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RETRÁTIL"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 9"
         Height          =   195
         Left            =   120
         TabIndex        =   87
         Top             =   4560
         Width           =   780
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 8"
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   4080
         Width           =   780
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 7"
         Height          =   195
         Left            =   120
         TabIndex        =   85
         Top             =   3600
         Width           =   780
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 6"
         Height          =   195
         Left            =   120
         TabIndex        =   84
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 5"
         Height          =   195
         Left            =   120
         TabIndex        =   83
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 4"
         Height          =   195
         Left            =   120
         TabIndex        =   82
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 3"
         Height          =   195
         Left            =   120
         TabIndex        =   81
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 2"
         Height          =   195
         Left            =   120
         TabIndex        =   80
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 1"
         Height          =   195
         Left            =   120
         TabIndex        =   79
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "centímetros"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   960
         TabIndex        =   59
         Top             =   4920
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CG grupo :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   5160
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5535
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Width           =   1935
      Begin VB.TextBox sv9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   54
         Text            =   "0"
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox sv8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   53
         Text            =   "0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox sv7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   52
         Text            =   "0"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox sv6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   51
         Text            =   "0"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox sv5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   50
         Text            =   "0"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox sv4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   49
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox sv3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   48
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox sv2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   47
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox sv1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   46
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox cgconj 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "0"
         Top             =   5160
         Width           =   615
      End
      Begin VB.OptionButton ts2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STANDARD"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 9"
         Height          =   195
         Left            =   120
         TabIndex        =   78
         Top             =   4560
         Width           =   780
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 8"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   4080
         Width           =   780
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 7"
         Height          =   195
         Left            =   120
         TabIndex        =   76
         Top             =   3600
         Width           =   780
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 6"
         Height          =   195
         Left            =   120
         TabIndex        =   75
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 5"
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 4"
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 3"
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 2"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 1"
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "centímetros"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   960
         TabIndex        =   58
         Top             =   4920
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CG grupo :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   5160
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   7995
      TabIndex        =   2
      Top             =   6000
      Width           =   8055
      Begin VB.CommandButton Command4 
         Caption         =   "CG ????"
         Height          =   375
         Left            =   6720
         MouseIcon       =   "Form19.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   97
         Top             =   0
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ca&ncelar e Sair"
         Height          =   375
         Left            =   4440
         MouseIcon       =   "Form19.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   0
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Enviar"
         Height          =   375
         Left            =   2160
         MouseIcon       =   "Form19.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   0
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Calcular"
         Height          =   375
         Left            =   0
         MouseIcon       =   "Form19.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FF0000&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1935
      Begin VB.TextBox mv9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   43
         Text            =   "0"
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox mv8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   42
         Text            =   "0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox mv7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         Text            =   "0"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox mv6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   40
         Text            =   "0"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox mv5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   39
         Text            =   "0"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox mv4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   38
         Text            =   "0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox mv3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   37
         Text            =   "0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox mv2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   36
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox mv1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   35
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox cgmicro 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0"
         Top             =   5160
         Width           =   615
      End
      Begin VB.OptionButton ts1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MICRO"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 9"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   4560
         Width           =   780
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 8"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   4080
         Width           =   780
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 7"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   3600
         Width           =   780
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 6"
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 5"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 4"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 3"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 2"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servo Nº 1"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "centímetros"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   960
         TabIndex        =   57
         Top             =   4920
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CG grupo :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   5160
         Width           =   765
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entre com a distancia entre a linha de eixo da hélice (L.R) ao CG de cada servo, em centímetros."
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   56
      Top             =   120
      Width           =   6870
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

If Form14.s1.Value = Checked And Form14.n1 = 0 Or Form14.s2.Value = Checked And Form14.n2 = 0 Or Form14.s3.Value = Checked And Form14.n3 = 0 Or Form14.s4.Value = Checked And Form14.n6 = 0 Then
MsgBox "Voce deve inserir o Nº de servos.", vbExclamation
 If Form14.s1.Value = Checked Then Form14.n1.SetFocus
 If Form14.s2.Value = Checked Then Form14.n2.SetFocus
 If Form14.s3.Value = Checked Then Form14.n3.SetFocus
 If Form14.s4.Value = Checked Then Form14.n6.SetFocus
Exit Sub
End If

Dim peso As Double
Dim nsv As Double
If ts1.Value = False And ts2.Value = False And ts3.Value = False And ts4.Value = False Then
MsgBox "Escolha o tipo de servos.", vbExclamation
Exit Sub
End If


If ts2.Value = True Then Form19.cgconj = (Abs(44 * sv1) + Abs(44 * sv2) + Abs(44 * sv3) + Abs(44 * sv4) + Abs(44 * sv5) + Abs(44 * sv6) + Abs(44 * sv7) + Abs(44 * sv8) + Abs(44 * sv9)) / (Form14.n2 * 44)

If ts1.Value = True Then Form19.cgmicro = (Abs(24 * mv1) + Abs(24 * mv2) + Abs(24 * mv3) + Abs(24 * mv4) + Abs(24 * mv5) + Abs(24 * mv6) + Abs(24 * mv7) + Abs(24 * mv8) + Abs(24 * mv9)) / (Form14.n1 * 24)

If ts3.Value = True Then Form19.cgretrat = (Abs(34 * rv1) + Abs(34 * rv2) + Abs(34 * rv3) + Abs(34 * rv4) + Abs(34 * rv5) + Abs(34 * rv6) + Abs(34 * rv7) + Abs(34 * rv8) + Abs(34 * rv9)) / (Form14.n3 * 34)

If ts4.Value = True Then Form19.cggiga = (Abs(44 * gv1) + Abs(44 * gv2) + Abs(44 * gv3) + Abs(44 * gv4) + Abs(44 * gv5) + Abs(44 * gv6) + Abs(44 * gv7) + Abs(44 * gv8) + Abs(44 * gv9)) / (Form14.n6 * 44)

Form18.Command6.Enabled = True

Command2.SetFocus
End Sub
Private Sub Command2_Click()
If ts1.Value = True Then Form14.d5 = cgmicro
If ts2.Value = True Then Form14.d6 = cgconj
If ts3.Value = True Then Form14.d7 = cgretrat
If ts4.Value = True Then Form14.d10 = cggiga
Form19.Visible = False
'sv1 = 0
'sv2 = 0
'sv3 = 0
'sv4 = 0
'sv5 = 0
'sv6 = 0
'sv7 = 0
'sv8 = 0
'sv9 = 0
'cgconj = 0
End Sub

Private Sub Command3_Click()
Form19.Visible = False
End Sub

Private Sub Command4_Click()
Picture1.Visible = True
End Sub

Private Sub Command5_Click()
Picture1.Visible = False
End Sub

Private Sub Form_Load()
ts1.Value = False
ts2.Value = False
ts3.Value = False
ts4.Value = False
End Sub

Private Sub sv1_GotFocus()
sv1.SelStart = 0
sv1.SelLength = Len(sv1)
End Sub
Private Sub sv1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If sv2.Enabled = True Then If KeyAscii = 13 Then sv2.SetFocus
End Sub
Private Sub sv2_GotFocus()
sv2.SelStart = 0
sv2.SelLength = Len(sv2)
End Sub
Private Sub sv2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If sv3.Enabled = True Then If KeyAscii = 13 Then sv3.SetFocus
End Sub
Private Sub sv3_GotFocus()
sv3.SelStart = 0
sv3.SelLength = Len(sv3)
End Sub
Private Sub sv3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If sv4.Enabled = True Then If KeyAscii = 13 Then sv4.SetFocus
End Sub
Private Sub sv4_GotFocus()
sv4.SelStart = 0
sv4.SelLength = Len(sv4)
End Sub
Private Sub sv4_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If sv5.Enabled = True Then If KeyAscii = 13 Then sv5.SetFocus
End Sub
Private Sub sv5_GotFocus()
sv5.SelStart = 0
sv5.SelLength = Len(sv5)
End Sub
Private Sub sv5_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If sv6.Enabled = True Then If KeyAscii = 13 Then sv6.SetFocus
End Sub
Private Sub sv6_GotFocus()
sv6.SelStart = 0
sv6.SelLength = Len(sv6)
End Sub
Private Sub sv6_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If sv7.Enabled = True Then If KeyAscii = 13 Then sv7.SetFocus
End Sub
Private Sub sv7_GotFocus()
sv7.SelStart = 0
sv7.SelLength = Len(sv7)
End Sub
Private Sub sv7_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If sv8.Enabled = True Then If KeyAscii = 13 Then sv8.SetFocus
End Sub
Private Sub sv8_GotFocus()
sv8.SelStart = 0
sv8.SelLength = Len(sv8)
End Sub
Private Sub sv8_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If sv9.Enabled = True Then If KeyAscii = 13 Then sv9.SetFocus
End Sub
Private Sub sv9_GotFocus()
sv9.SelStart = 0
sv9.SelLength = Len(sv9)
End Sub
Private Sub sv9_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command2.SetFocus
End Sub
'Para micro servos
Private Sub mv1_GotFocus()
mv1.SelStart = 0
mv1.SelLength = Len(mv1)
End Sub
Private Sub mv1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If mv2.Enabled = True Then If KeyAscii = 13 Then mv2.SetFocus
End Sub
Private Sub mv2_GotFocus()
mv2.SelStart = 0
mv2.SelLength = Len(mv2)
End Sub
Private Sub mv2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If mv3.Enabled = True Then If KeyAscii = 13 Then mv3.SetFocus
End Sub
Private Sub mv3_GotFocus()
mv3.SelStart = 0
mv3.SelLength = Len(mv3)
End Sub
Private Sub mv3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If mv4.Enabled = True Then If KeyAscii = 13 Then mv4.SetFocus
End Sub
Private Sub mv4_GotFocus()
mv4.SelStart = 0
mv4.SelLength = Len(mv4)
End Sub
Private Sub mv4_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If mv5.Enabled = True Then If KeyAscii = 13 Then mv5.SetFocus
End Sub
Private Sub mv5_GotFocus()
mv5.SelStart = 0
mv5.SelLength = Len(mv5)
End Sub
Private Sub mv5_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If mv6.Enabled = True Then If KeyAscii = 13 Then mv6.SetFocus
End Sub
Private Sub mv6_GotFocus()
mv6.SelStart = 0
mv6.SelLength = Len(mv6)
End Sub
Private Sub mv6_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If mv7.Enabled = True Then If KeyAscii = 13 Then mv7.SetFocus
End Sub
Private Sub mv7_GotFocus()
mv7.SelStart = 0
mv7.SelLength = Len(mv7)
End Sub
Private Sub mv7_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If mv8.Enabled = True Then If KeyAscii = 13 Then mv8.SetFocus
End Sub
Private Sub mv8_GotFocus()
mv8.SelStart = 0
mv8.SelLength = Len(mv8)
End Sub
Private Sub mv8_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If mv9.Enabled = True Then If KeyAscii = 13 Then mv9.SetFocus
End Sub
Private Sub mv9_GotFocus()
mv9.SelStart = 0
mv9.SelLength = Len(mv9)
End Sub
Private Sub mv9_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command2.SetFocus
End Sub
'Para servos retratil
Private Sub rv1_GotFocus()
rv1.SelStart = 0
rv1.SelLength = Len(rv1)
End Sub
Private Sub rv1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If rv2.Enabled = True Then If KeyAscii = 13 Then rv2.SetFocus
End Sub
Private Sub rv2_GotFocus()
rv2.SelStart = 0
rv2.SelLength = Len(rv2)
End Sub
Private Sub rv2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If rv3.Enabled = True Then If KeyAscii = 13 Then rv3.SetFocus
End Sub
Private Sub rv3_GotFocus()
rv3.SelStart = 0
rv3.SelLength = Len(rv3)
End Sub
Private Sub rv3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If rv4.Enabled = True Then If KeyAscii = 13 Then rv4.SetFocus
End Sub
Private Sub rv4_GotFocus()
rv4.SelStart = 0
rv4.SelLength = Len(rv4)
End Sub
Private Sub rv4_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If rv5.Enabled = True Then If KeyAscii = 13 Then rv5.SetFocus
End Sub
Private Sub rv5_GotFocus()
rv5.SelStart = 0
rv5.SelLength = Len(rv5)
End Sub
Private Sub rv5_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If rv6.Enabled = True Then If KeyAscii = 13 Then rv6.SetFocus
End Sub
Private Sub rv6_GotFocus()
rv6.SelStart = 0
rv6.SelLength = Len(rv6)
End Sub
Private Sub rv6_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If rv7.Enabled = True Then If KeyAscii = 13 Then rv7.SetFocus
End Sub
Private Sub rv7_GotFocus()
rv7.SelStart = 0
rv7.SelLength = Len(rv7)
End Sub
Private Sub rv7_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If rv8.Enabled = True Then If KeyAscii = 13 Then rv8.SetFocus
End Sub
Private Sub rv8_GotFocus()
rv8.SelStart = 0
rv8.SelLength = Len(rv8)
End Sub
Private Sub rv8_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If rv9.Enabled = True Then If KeyAscii = 13 Then rv9.SetFocus
End Sub
Private Sub rv9_GotFocus()
rv9.SelStart = 0
rv9.SelLength = Len(rv9)
End Sub
Private Sub rv9_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command2.SetFocus
End Sub
'Para servos gigantes
Private Sub gv1_GotFocus()
gv1.SelStart = 0
gv1.SelLength = Len(gv1)
End Sub
Private Sub gv1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If gv2.Enabled = True Then If KeyAscii = 13 Then gv2.SetFocus
End Sub
Private Sub gv2_GotFocus()
gv2.SelStart = 0
gv2.SelLength = Len(gv2)
End Sub
Private Sub gv2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If gv3.Enabled = True Then If KeyAscii = 13 Then gv3.SetFocus
End Sub
Private Sub gv3_GotFocus()
gv3.SelStart = 0
gv3.SelLength = Len(gv3)
End Sub
Private Sub gv3_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If gv4.Enabled = True Then If KeyAscii = 13 Then gv4.SetFocus
End Sub
Private Sub gv4_GotFocus()
gv4.SelStart = 0
gv4.SelLength = Len(gv4)
End Sub
Private Sub gv4_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If gv5.Enabled = True Then If KeyAscii = 13 Then gv5.SetFocus
End Sub
Private Sub gv5_GotFocus()
gv5.SelStart = 0
gv5.SelLength = Len(gv5)
End Sub
Private Sub gv5_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If gv6.Enabled = True Then If KeyAscii = 13 Then gv6.SetFocus
End Sub
Private Sub gv6_GotFocus()
gv6.SelStart = 0
gv6.SelLength = Len(gv6)
End Sub
Private Sub gv6_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If gv7.Enabled = True Then If KeyAscii = 13 Then gv7.SetFocus
End Sub
Private Sub gv7_GotFocus()
gv7.SelStart = 0
gv7.SelLength = Len(gv7)
End Sub
Private Sub gv7_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If gv8.Enabled = True Then If KeyAscii = 13 Then gv8.SetFocus
End Sub
Private Sub gv8_GotFocus()
gv8.SelStart = 0
gv8.SelLength = Len(gv8)
End Sub
Private Sub gv8_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If gv9.Enabled = True Then If KeyAscii = 13 Then gv9.SetFocus
End Sub
Private Sub gv9_GotFocus()
gv9.SelStart = 0
gv9.SelLength = Len(gv9)
End Sub
Private Sub gv9_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command2.SetFocus
End Sub
