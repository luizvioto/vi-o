VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualização - Estruturas"
   ClientHeight    =   7335
   ClientLeft      =   285
   ClientTop       =   570
   ClientWidth     =   8445
   ForeColor       =   &H00808080&
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H000007A8&
      Height          =   1960
      Left            =   240
      ScaleHeight     =   1905
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   5160
      Width           =   7935
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   0
         Width           =   7695
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CL em Vstall"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   185
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CL em Vdecolagem"
            Height          =   255
            Left            =   1920
            TabIndex        =   36
            Top             =   185
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CL em Vsubida"
            Height          =   255
            Left            =   4200
            TabIndex        =   35
            Top             =   185
            Width           =   1455
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CL em Vmáxima"
            Height          =   255
            Left            =   6120
            TabIndex        =   34
            Top             =   185
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   7695
         Begin VB.TextBox lambda 
            Alignment       =   2  'Center
            BackColor       =   &H00400000&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   6840
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   495
            Width           =   615
         End
         Begin VB.TextBox clestim 
            Alignment       =   2  'Center
            BackColor       =   &H00400000&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   495
            Width           =   615
         End
         Begin VB.TextBox cp 
            Alignment       =   2  'Center
            BackColor       =   &H00400000&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   495
            Width           =   615
         End
         Begin VB.TextBox cr 
            Alignment       =   2  'Center
            BackColor       =   &H00400000&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   495
            Width           =   615
         End
         Begin VB.TextBox bw 
            Alignment       =   2  'Center
            BackColor       =   &H00400000&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   495
            Width           =   615
         End
         Begin VB.TextBox sw 
            Alignment       =   2  'Center
            BackColor       =   &H00400000&
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   495
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Área da asa"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Envergadura"
            Height          =   195
            Left            =   1320
            TabIndex        =   31
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Corda da raiz"
            Height          =   195
            Left            =   2640
            TabIndex        =   30
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Corda da ponta"
            Height          =   195
            Left            =   3840
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "CL selecionado"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5280
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Afilamento"
            Height          =   195
            Left            =   6720
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   7635
         TabIndex        =   13
         Top             =   1440
         Width           =   7695
         Begin VB.CommandButton Command2 
            Caption         =   "&Sair"
            Height          =   315
            Left            =   6240
            TabIndex        =   18
            Top             =   0
            Width           =   1400
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&CL"
            Height          =   315
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "C&argas"
            Height          =   315
            Left            =   1560
            TabIndex        =   16
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton Command5 
            Caption         =   "C&ortante"
            Height          =   315
            Left            =   3120
            TabIndex        =   15
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton Command6 
            Caption         =   "&Fletor"
            Height          =   315
            Left            =   4680
            TabIndex        =   14
            Top             =   0
            Width           =   1575
         End
      End
   End
   Begin VB.PictureBox gf 
      BackColor       =   &H00400000&
      Height          =   4900
      Left            =   240
      ScaleHeight     =   4845
      ScaleWidth      =   7845
      TabIndex        =   0
      Top             =   120
      Width           =   7905
      Begin VB.Shape Shape4 
         BorderColor     =   &H00E0E0E0&
         Height          =   1215
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   6840
         TabIndex        =   19
         Top             =   0
         Width           =   90
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   255
         Left            =   4680
         Top             =   4200
         Width           =   135
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   255
         Left            =   2640
         Top             =   4200
         Width           =   135
      End
      Begin VB.Line Line40 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   6960
         X2              =   6960
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line38 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   480
         X2              =   480
         Y1              =   4655
         Y2              =   120
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00FFFFFF&
         Height          =   4560
         Left            =   120
         Top             =   120
         Width           =   7095
      End
      Begin VB.Label cl11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7440
         TabIndex        =   12
         Top             =   5040
         Width           =   75
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00400000&
         Height          =   255
         Left            =   3360
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   750
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   3600
         X2              =   3840
         Y1              =   3885
         Y2              =   3885
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   3600
         Width           =   255
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   3600
         X2              =   3840
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   255
         Left            =   3480
         Shape           =   2  'Oval
         Top             =   3600
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FFFF&
         Height          =   735
         Left            =   3360
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label cl10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7320
         TabIndex        =   11
         Top             =   4560
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label cl9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7320
         TabIndex        =   10
         Top             =   4080
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label cl8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7320
         TabIndex        =   9
         Top             =   3600
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label cl7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7320
         TabIndex        =   8
         Top             =   3120
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label cl6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7320
         TabIndex        =   7
         Top             =   2640
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label cl5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7320
         TabIndex        =   6
         Top             =   2160
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label cl4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7320
         TabIndex        =   5
         Top             =   1680
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label cl3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7320
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label cl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7320
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label cl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   7320
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Line sp1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   4
         X1              =   6940
         X2              =   510
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   3  'Dot
         BorderWidth     =   3
         X1              =   3720
         X2              =   3720
         Y1              =   3360
         Y2              =   2640
      End
      Begin VB.Shape beq 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   255
         Left            =   3660
         Top             =   4200
         Width           =   135
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   2400
         X2              =   5040
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FFFF&
         X1              =   3360
         X2              =   2760
         Y1              =   4080
         Y2              =   4320
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0000FFFF&
         X1              =   4080
         X2              =   4680
         Y1              =   4080
         Y2              =   4320
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   804
         X2              =   804
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line6 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   1128
         X2              =   1128
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line7 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   1452
         X2              =   1452
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line8 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   1776
         X2              =   1776
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line9 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   2100
         X2              =   2100
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line10 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   2424
         X2              =   2424
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line11 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   2748
         X2              =   2748
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line12 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   3072
         X2              =   3072
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line13 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   3396
         X2              =   3396
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line16 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   7200
         X2              =   120
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line17 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   7200
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line18 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   7200
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line19 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   7200
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line20 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   7200
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line15 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   4044
         X2              =   4044
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line21 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   4368
         X2              =   4368
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line22 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   4692
         X2              =   4692
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line23 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   5340
         X2              =   5340
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line24 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   5664
         X2              =   5664
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line25 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   5016
         X2              =   5016
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line26 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   6312
         X2              =   6312
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line27 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   5988
         X2              =   5988
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line28 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   6636
         X2              =   6636
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line Line29 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   7200
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line31 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   7200
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line32 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   7200
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line33 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   7200
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line34 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   7200
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line35 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   7200
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line36 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   7200
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line30 
         BorderColor     =   &H0000C0C0&
         BorderStyle     =   3  'Dot
         X1              =   3720
         X2              =   3720
         Y1              =   120
         Y2              =   4655
      End
      Begin VB.Line eixobeq 
         BorderColor     =   &H0000C0C0&
         X1              =   3720
         X2              =   3720
         Y1              =   4680
         Y2              =   4440
      End
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Form12.Option1.Value = False And Form12.Option2.Value = False And Form12.Option3.Value = False And Form12.Option4.Value = False Then
MsgBox " Escolha o Cl para visualização", vbInformation
Exit Sub
End If
gf.Refresh

gf.DrawWidth = 1

cl = clestim

'CONVERSÃO P/ POLEGADA
swcalc = Sw * 1550
bwcalc = (bw / 2) * 39.37
bwcalc = Format(bwcalc, "0,00")
crcalc = cr * 39.37
cpcalc = cp * 39.37

For estaca = 0 To bwcalc Step 0.01
'1
Y = estaca
Y = Format(Y, "0.000")
'2
yd = Y / (bwcalc)
yd = Format(yd, "0.000")

'3
c = crcalc * (1 + ((lambda - 1) * yd))
c = Format(c, "0.000")

'4
cla = ((((1 + (4 * swcalc)) / (3.1415 * (bwcalc) * c) * (1 - (yd ^ 2)) ^ 0.5)) / 2) * clestim
cla = Format(cla, "0.000")

sp1.x1 = 480
sp1.X2 = 6960

'GRAFICO
If Form1.posialta Then asa = 3360
If Form1.posimedia Then asa = 3790
If Form1.posibaixa Then asa = 3990

gf.PSet (3720 + (Y * (3240 / bwcalc)), asa - (cla * 1000)), &HFFFF80
gf.PSet (3720 - (Y * (3240 / bwcalc)), asa - (cla * 1000)), &HFFFF80

'VALORES DE CL

Label1.Caption = "CL"

If yd = 0 Then
cl1 = cla
cl1 = Format(cl1, "0.000")
gf.Line (3720, asa)-(3720, asa - (cla * 1000)), &HFFFF80
cl1.Visible = True
End If

If yd = 0.1 Then
cl2 = cla
cl2 = Format(cl2, "0.000")
gf.Line (3396, asa)-(3396, asa - (cla * 1000)), &HFFFF80
gf.Line (4044, asa)-(4044, asa - (cla * 1000)), &HFFFF80
cl2.Visible = True
End If

If yd = 0.2 Then
cl3 = cla
cl3 = Format(cl3, "0.000")
gf.Line (3072, asa)-(3072, asa - (cla * 1000)), &HFFFF80
gf.Line (4368, asa)-(4368, asa - (cla * 1000)), &HFFFF80
cl3.Visible = True
End If

If yd = 0.3 Then
cl4 = cla
cl4 = Format(cl4, "0.000")
gf.Line (2748, asa)-(2748, asa - (cla * 1000)), &HFFFF80
gf.Line (4692, asa)-(4692, asa - (cla * 1000)), &HFFFF80
cl4.Visible = True
End If

If yd = 0.4 Then
cl5 = cla
cl5 = Format(cl5, "0.000")
gf.Line (2424, asa)-(2424, asa - (cla * 1000)), &HFFFF80
gf.Line (5016, asa)-(5016, asa - (cla * 1000)), &HFFFF80
cl5.Visible = True
End If

If yd = 0.5 Then
cl6 = cla
cl6 = Format(cl6, "0.000")
gf.Line (2100, asa)-(2100, asa - (cla * 1000)), &HFFFF80
gf.Line (5340, asa)-(5340, asa - (cla * 1000)), &HFFFF80
cl6.Visible = True
End If

If yd = 0.6 Then
cl7 = cla
cl7 = Format(cl7, "0.000")
gf.Line (1776, asa)-(1776, asa - (cla * 1000)), &HFFFF80
gf.Line (5664, asa)-(5664, asa - (cla * 1000)), &HFFFF80
cl7.Visible = True
End If

If yd = 0.7 Then
cl8 = cla
cl8 = Format(cl8, "0.000")
gf.Line (1452, asa)-(1452, asa - (cla * 1000)), &HFFFF80
gf.Line (5988, asa)-(5988, asa - (cla * 1000)), &HFFFF80
cl8.Visible = True
End If

If yd = 0.8 Then
cl9 = cla
cl9 = Format(cl9, "0.000")
gf.Line (1128, asa)-(1128, asa - (cla * 1000)), &HFFFF80
gf.Line (6312, asa)-(6312, asa - (cla * 1000)), &HFFFF80
cl9.Visible = True
End If

If yd = 0.9 Then
cl10 = cla
cl10 = Format(cl10, "0.000")
gf.Line (804, asa)-(804, asa - (cla * 1000)), &HFFFF80
gf.Line (6636, asa)-(6636, asa - (cla * 1000)), &HFFFF80
cl10.Visible = True
End If

If yd = 1 Then
cl11 = cla
cl11 = Format(cl11, "0.000")
gf.Line (480, asa)-(480, asa - (cla * 1000)), &HFFFF80
gf.Line (6960, asa)-(6960, asa - (cla * 1000)), &HFFFF80
cl10.Visible = True
End If

Next

End Sub

Private Sub Command2_Click()
Form12.Visible = False
Label1 = "-"
cl1 = "-"
cl2 = "-"
cl3 = "-"
cl4 = "-"
cl5 = "-"
cl6 = "-"
cl7 = "-"
cl8 = "-"
cl9 = "-"
cl10 = "-"
cl11 = "-"
End Sub

Private Sub Command4_Click()
If clestim = 0 Then
MsgBox " Escolha o Cl para visualização", vbInformation
Form12.Visible = False
Form13.Visible = True
Exit Sub
End If

gf.Refresh
bwcalc = (bw / 2) * 39.37

If Form1.posialta Then asa = 3360
If Form1.posimedia Then asa = 3790
If Form1.posibaixa Then asa = 3990

a1 = Form11.q0 * 10
a2 = Form11.q1 * 10
a3 = Form11.q2 * 10
a4 = Form11.q3 * 10
a5 = Form11.q4 * 10
a6 = Form11.q5 * 10
a7 = Form11.q6 * 10
a8 = Form11.q7 * 10
a9 = Form11.q8 * 10
a10 = Form11.q9 * 10
a95 = (Form11.q9 * 0.6) * 10
a11 = Form11.q10 * 10

Label1.Caption = "Q(kgf)"
cl1 = Form11.q0
cl2 = Form11.q1
cl3 = Form11.q2
cl4 = Form11.q3
cl5 = Form11.q4
cl6 = Form11.q5
cl7 = Form11.q6
cl8 = Form11.q7
cl9 = Form11.q8
cl10 = Form11.q9
cl11 = Form11.q10


gf.Line (3360, asa - a1)-(4044, asa - a2), &HFF00&
gf.Line (4044, asa - a2)-(4368, asa - a3), &HFF00&
gf.Line (4368, asa - a3)-(4692, asa - a4), &HFF00&
gf.Line (4692, asa - a4)-(5016, asa - a5), &HFF00&
gf.Line (5016, asa - a5)-(5340, asa - a6), &HFF00&
gf.Line (5340, asa - a6)-(5664, asa - a7), &HFF00&
gf.Line (5664, asa - a7)-(5988, asa - a8), &HFF00&
gf.Line (5988, asa - a8)-(6312, asa - a9), &HFF00&
gf.Line (6312, asa - a9)-(6636, asa - a10), &HFF00&
gf.Line (6636, asa - a10)-(6830, asa - a95), &HFF00&
gf.Line (6830, asa - a95)-(6960, asa - a11), &HFF00&

gf.Line (3360, asa - a1)-(3396, asa - a2), &HFF00&
gf.Line (3396, asa - a2)-(3072, asa - a3), &HFF00&
gf.Line (3072, asa - a3)-(2748, asa - a4), &HFF00&
gf.Line (2748, asa - a4)-(2424, asa - a5), &HFF00&
gf.Line (2424, asa - a5)-(2100, asa - a6), &HFF00&
gf.Line (2100, asa - a6)-(1776, asa - a7), &HFF00&
gf.Line (1776, asa - a7)-(1452, asa - a8), &HFF00&
gf.Line (1452, asa - a8)-(1128, asa - a9), &HFF00&
gf.Line (1128, asa - a9)-(804, asa - a10), &HFF00&
gf.Line (804, asa - a10)-(610, asa - a95), &HFF00&
gf.Line (610, asa - a95)-(480, asa - a11), &HFF00&
End Sub

Private Sub Command5_Click()
If clestim = 0 Then
MsgBox " Escolha o Cl para visualização", vbInformation
Form12.Visible = False
Form13.Visible = True
Exit Sub
End If

gf.Refresh
bwcalc = (bw / 2) * 39.37

'For estaca = 0 To bwcalc Step 0.01
'y = estaca

If Form1.posialta Then asa = 3360
If Form1.posimedia Then asa = 3790
If Form1.posibaixa Then asa = 3990

a1 = Form11.v0
a2 = Form11.v1
a3 = Form11.v2
a4 = Form11.v3
a5 = Form11.v4
a6 = Form11.v5
a7 = Form11.v6
a8 = Form11.v7
a9 = Form11.v8
a10 = Form11.v9
a11 = Form11.v10

Label1.Caption = "C(kgf)"
cl1 = Form11.v0
cl2 = Form11.v1
cl3 = Form11.v2
cl4 = Form11.v3
cl5 = Form11.v4
cl6 = Form11.v5
cl7 = Form11.v6
cl8 = Form11.v7
cl9 = Form11.v8
cl10 = Form11.v9
cl11 = Form11.v10

gf.Line (3720, asa - a1)-(4044, asa - a2), &HFF00FF
gf.Line (4044, asa - a2)-(4368, asa - a3), &HFF00FF
gf.Line (4368, asa - a3)-(4692, asa - a4), &HFF00FF
gf.Line (4692, asa - a4)-(5016, asa - a5), &HFF00FF
gf.Line (5016, asa - a5)-(5340, asa - a6), &HFF00FF
gf.Line (5340, asa - a6)-(5664, asa - a7), &HFF00FF
gf.Line (5664, asa - a7)-(5988, asa - a8), &HFF00FF
gf.Line (5988, asa - a8)-(6312, asa - a9), &HFF00FF
gf.Line (6312, asa - a9)-(6636, asa - a10), &HFF00FF
gf.Line (6636, asa - a10)-(6960, asa - a11), &HFF00FF

gf.Line (3720, asa - a1)-(3396, asa - a2), &HFF00FF
gf.Line (3396, asa - a2)-(3072, asa - a3), &HFF00FF
gf.Line (3072, asa - a3)-(2748, asa - a4), &HFF00FF
gf.Line (2748, asa - a4)-(2424, asa - a5), &HFF00FF
gf.Line (2424, asa - a5)-(2100, asa - a6), &HFF00FF
gf.Line (2100, asa - a6)-(1776, asa - a7), &HFF00FF
gf.Line (1776, asa - a7)-(1452, asa - a8), &HFF00FF
gf.Line (1452, asa - a8)-(1128, asa - a9), &HFF00FF
gf.Line (1128, asa - a9)-(804, asa - a10), &HFF00FF
gf.Line (804, asa - a10)-(480, asa - a11), &HFF00FF


'Next

End Sub

Private Sub Command6_Click()
If clestim = 0 Then
MsgBox " Escolha o Cl para visualização", vbInformation
Form12.Visible = False
Form13.Visible = True
Exit Sub
End If

gf.Refresh
If Form1.posialta Then asa = 3360
If Form1.posimedia Then asa = 3790
If Form1.posibaixa Then asa = 3990

a1 = Form11.m0
a2 = Form11.m1
a3 = Form11.m2
a4 = Form11.m3
a5 = Form11.m4
a6 = Form11.m5
a7 = Form11.m6
a8 = Form11.m7
a9 = Form11.m8
a10 = Form11.m9
a11 = Form11.m10

Label1.Caption = "M(kgfxcm)"
cl1 = Form11.m0
cl2 = Form11.m1
cl3 = Form11.m2
cl4 = Form11.m3
cl5 = Form11.m4
cl6 = Form11.m5
cl7 = Form11.m6
cl8 = Form11.m7
cl9 = Form11.m8
cl10 = Form11.m9
cl11 = Form11.m10

gf.Line (3720, asa - a1)-(4044, asa - a2), &HFF&
gf.Line (4044, asa - a2)-(4368, asa - a3), &HFF&
gf.Line (4368, asa - a3)-(4692, asa - a4), &HFF&
gf.Line (4692, asa - a4)-(5016, asa - a5), &HFF&
gf.Line (5016, asa - a5)-(5340, asa - a6), &HFF&
gf.Line (5340, asa - a6)-(5664, asa - a7), &HFF&
gf.Line (5664, asa - a7)-(5988, asa - a8), &HFF&
gf.Line (5988, asa - a8)-(6312, asa - a9), &HFF&
gf.Line (6312, asa - a9)-(6636, asa - a10), &HFF&
gf.Line (6636, asa - a10)-(6960, asa - a11), &HFF&

gf.Line (3720, asa - a1)-(3396, asa - a2), &HFF&
gf.Line (3396, asa - a2)-(3072, asa - a3), &HFF&
gf.Line (3072, asa - a3)-(2748, asa - a4), &HFF&
gf.Line (2748, asa - a4)-(2424, asa - a5), &HFF&
gf.Line (2424, asa - a5)-(2100, asa - a6), &HFF&
gf.Line (2100, asa - a6)-(1776, asa - a7), &HFF&
gf.Line (1776, asa - a7)-(1452, asa - a8), &HFF&
gf.Line (1452, asa - a8)-(1128, asa - a9), &HFF&
gf.Line (1128, asa - a9)-(804, asa - a10), &HFF&
gf.Line (804, asa - a10)-(480, asa - a11), &HFF&


End Sub

Private Sub Option1_Click()
Form12.clestim = Form7.clk1
End Sub

Private Sub Option2_Click()
Form12.clestim = Form7.cldecol
End Sub

Private Sub Option3_Click()
Form12.clestim = Form7.clk2
End Sub

Private Sub Option4_Click()
Form12.clestim = Form7.clk4
End Sub
