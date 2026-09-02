VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualização - Aerodinâmica"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   ForeColor       =   &H00808080&
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "V"
      Height          =   255
      Left            =   7800
      TabIndex        =   39
      Top             =   0
      Width           =   495
   End
   Begin VB.ListBox L1 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFF00&
      Height          =   1230
      ItemData        =   "Form9.frx":030A
      Left            =   240
      List            =   "Form9.frx":030C
      TabIndex        =   38
      Top             =   240
      Width           =   8055
   End
   Begin VB.PictureBox grafico 
      BackColor       =   &H00400000&
      Height          =   5055
      Left            =   240
      ScaleHeight     =   4995
      ScaleWidth      =   7995
      TabIndex        =   1
      Top             =   1440
      Width           =   8055
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidade de stall sem flap :              Km/h"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2400
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.Label vst 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "vst"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4560
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00FFFFFF&
         X1              =   360
         X2              =   7320
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   1800
         X2              =   1800
         Y1              =   3720
         Y2              =   840
      End
      Begin VB.Line Line43 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   1800
         X2              =   1800
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FFFFFF&
         X1              =   600
         X2              =   600
         Y1              =   840
         Y2              =   4080
      End
      Begin VB.Line Line42 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   840
         X2              =   840
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line41 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   1080
         X2              =   1080
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   1320
         X2              =   1320
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vplaneio"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   4440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label vsb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vsb"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   4680
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Line sp 
         BorderColor     =   &H0000FFFF&
         Visible         =   0   'False
         X1              =   600
         X2              =   600
         Y1              =   4320
         Y2              =   3720
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   7080
         X2              =   7080
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line38 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   6840
         X2              =   6840
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line37 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   6600
         X2              =   6600
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line36 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   6360
         X2              =   6360
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   6120
         X2              =   6120
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   5880
         X2              =   5880
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   5640
         X2              =   5640
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line32 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   5400
         X2              =   5400
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line31 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   5160
         X2              =   5160
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   4920
         X2              =   4920
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line29 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   4680
         X2              =   4680
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   4440
         X2              =   4440
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   4200
         X2              =   4200
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   3960
         X2              =   3960
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   3720
         X2              =   3720
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   3480
         X2              =   3480
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   3240
         X2              =   3240
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   3000
         X2              =   3000
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   2760
         X2              =   2760
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   2520
         X2              =   2520
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   2280
         X2              =   2280
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   2040
         X2              =   2040
         Y1              =   840
         Y2              =   3720
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   1560
         X2              =   1560
         Y1              =   3720
         Y2              =   840
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7320
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   7320
         X2              =   7320
         Y1              =   3720
         Y2              =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vmáx."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   7080
         TabIndex        =   8
         Top             =   3840
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vstall"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   3840
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label velomaxima 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   7080
         TabIndex        =   6
         Top             =   4080
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label velominima 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   1560
         TabIndex        =   5
         Top             =   4080
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   3840
         Width           =   120
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cd"
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
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V"
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
         Height          =   195
         Left            =   7440
         TabIndex        =   2
         Top             =   3600
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   6600
      Width           =   8055
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Step"
         Height          =   615
         Left            =   6240
         TabIndex        =   35
         Top             =   960
         Width           =   1695
         Begin VB.TextBox STPv 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1080
            TabIndex        =   37
            Top             =   240
            Width           =   495
         End
         Begin VB.HScrollBar stps 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   36
            Top             =   240
            Value           =   10
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Velocidade inicial"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   3120
         TabIndex        =   29
         Top             =   960
         Width           =   3015
         Begin VB.TextBox vinicio 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1320
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
         Begin VB.HScrollBar hsvinicio 
            Height          =   285
            Left            =   120
            Max             =   4
            Min             =   1
            TabIndex        =   30
            Top             =   240
            Value           =   4
            Width           =   1095
         End
         Begin VB.Label textvinic 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2040
            TabIndex        =   32
            Top             =   240
            Width           =   45
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Espessura das linhas"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   2895
         Begin VB.OptionButton elin3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "E = 3"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1920
            TabIndex        =   28
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton elin2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "E = 2"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1080
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton elin1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "E = 1"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   735
         Left            =   6240
         ScaleHeight     =   675
         ScaleWidth      =   1635
         TabIndex        =   14
         Top             =   120
         Width           =   1695
         Begin VB.CommandButton Command1 
            Caption         =   "&Sair"
            Height          =   690
            Left            =   0
            MouseIcon       =   "Form9.frx":030E
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   0
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   1635
         TabIndex        =   12
         Top             =   120
         Width           =   1695
         Begin VB.CommandButton Command3 
            Caption         =   "&Ver arrastos"
            Height          =   690
            Left            =   0
            MouseIcon       =   "Form9.frx":0618
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   0
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00400000&
         Height          =   735
         Left            =   1920
         ScaleHeight     =   675
         ScaleWidth      =   4155
         TabIndex        =   11
         Top             =   120
         Width           =   4215
         Begin VB.Label at 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00FF00FF&
            Height          =   195
            Left            =   3390
            TabIndex        =   24
            Top             =   360
            Width           =   75
         End
         Begin VB.Label a2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   2010
            TabIndex        =   23
            Top             =   360
            Width           =   75
         End
         Begin VB.Label a1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   450
            TabIndex        =   22
            Top             =   360
            Width           =   75
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Arrasto Induzido "
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   0
            Width           =   1185
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "kgf."
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   3840
            TabIndex        =   20
            Top             =   360
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "kgf."
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2520
            TabIndex        =   19
            Top             =   360
            Width           =   270
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "kgf."
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   960
            TabIndex        =   18
            Top             =   360
            Width           =   270
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Arrasto Total"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   3120
            TabIndex        =   17
            Top             =   0
            Width           =   900
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Arrasto Parasita"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   1680
            TabIndex        =   16
            Top             =   0
            Width           =   1110
         End
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Label7.Visible = False
Label8.Visible = False
velominima.Visible = False
velomaxima.Visible = False
Label5.Visible = False
vsb.Visible = False
Line1.Visible = False
Line2.Visible = False
Label11.Visible = False
vst.Visible = False
sp.Visible = False
a1.Caption = "- "
a2.Caption = "-"
at.Caption = "-"
hsvinicio.Value = 4

Form9.Visible = False
End Sub


Private Sub Command2_Click()
L1.Height = 6300
End Sub

Private Sub Command3_Click()
Form9.grafico.Refresh
'linha de stall
Line1.Visible = True

'POTENCIA + RANKINE + CFuse
            hp = Form1.motor
            trank = 519
            kfuse = 2.5
                      
'Opção quanto ao acabamento
If Requisitos.madeira.Value = True Then fatpaint = 0.00046
If Requisitos.composto.Value = True Then fatpaint = 0.0002
If Requisitos.tubo.Value = True Then fatpaint = 0.00048
If Requisitos.metal.Value = True Then fatpaint = 0.00035
If Requisitos.Tubalum.Value = True Then fatpaint = 0.00035
                'fuselagem
reyf7 = 38.21 * (Form2.lfs * 3.28 / fatpaint) ^ 1.053
reyf7 = Format(reyf7, "0.00")
lfsdia = (Form2.lfs * 3.28) / (((Form2.wfs * 3.28) * (Form2.hfs * 3.28)) ^ 0.5)
gf7 = Log(reyf7) / Log(10)

                'asa
reyw7 = 38.21 * (((Form7.CMaerodw * 3.28) / fatpaint) ^ 1.053)
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
areaflap = (((((Form2.cmflap * Cos(Form4.deflap * 3.1415 / 180)) * Form2.bflap) * 10.76) * 2) * 0.1)
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

'VELOCIDADES
'velok é dado em km/h

'Precisão da velocidade -  step e adição em Vstall

'If Form4.km5.Value = True Then
STP = STPv
If elin1.Value = True Then gdw = 1
If elin2.Value = True Then gdw = 2
If elin3.Value = True Then gdw = 3
grafico.DrawWidth = gdw
'End If


'Velocidade inicial do LOOP
'If vist.Value = True Then vinicio = velominima
'If vist50.Value = True Then vinicio = 0.75 * velominima


'LOOP utilizando velocidades já calculadas anteriormente

For velok = vinicio To Form9.velomaxima Step STP
'km/h para mph
velokmph = velok / 1.61

'km/h para ft/s
velokpe = velok * 0.911

'Em Mach
velokmach = velok / 1224
velokmach = Format(velokmach, "0.000")

'Area Molhada
'29/02/2004

If Form20.p1 = "PCA" Or Form20.p1 = "PCA I" Or Form20.p1 = "PCA II" Then swetf7 = (6.283 * ((0.3819 * Form2.lfs ^ 2 + 0.12795 * Form2.lfs ^ 2 ^ 2 - 0.0091 * Form2.lfs ^ 2 ^ 3)))
If Form20.p1 <> "PCA" Then
    If Val(Form2.lfs) < 8.5 Then swetf7 = (6.283 * ((0.1477 * Form2.lfs + 0.08275 * Form2.lfs ^ 2 - 0.007766 * Form2.lfs ^ 3)))
    If Val(Form2.lfs) > 8.5 Then swetf7 = 6.283 * ((0.2288 * Form2.lfs + 0.0397 * Form2.lfs ^ 2 - 0.003133 * Form2.lfs ^ 3))
    If Val(Form2.lfs) = 8.5 Then swetf7 = 6.283 * ((0.2288 * Form2.lfs + 0.0397 * Form2.lfs ^ 2 - 0.003133 * Form2.lfs ^ 3))
End If
'antigo swetf7 = kfuse * (((Form2.wfs * 3.28) * (Form2.hfs * 3.28)) ^ 0.5) * (Form2.lfs * 3.28)
Form7.swetfuse = swetf7
Form7.swetfuse = Format(Form7.swetfuse, "0.000")

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
'arrasto de interferencia cdif=1.10
cdif = 1.05
'Sem arrasto do Perfil
'cdototal7 = cdif * ((Abs(cdof7) + Abs(cdow7) + Abs(cdoht7) + Abs(cdovt7) + Abs(cdotp)) + Abs(cdocool) + Abs(cdomotor) + Abs(cdosuple))
'Com arrasto do Perfil
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
velokint = Int(velok)

'LISTAR ARRASTO
L1.AddItem "Velocidade (km/h) =  " & velokint & vbTab & "Cdi =  " & cdi7 & vbTab & "Cdf =  " & cdototal7 & vbTab & "Cdo = " & cd7

'.................. velocidade de subida ........................
    If Abs(cdototal7 * Form2.Sw * 10.76 * qd7) - Abs(cdi7 * Form2.Sw * 10.76 * qd7) < 0.000001 Then
            sp.Visible = True
            sp.x1 = i
            sp.X2 = i
            Label5.Visible = True
            Label5.Left = i - 235
            vsb.Visible = True
            vsb = velok
            vsb = Format(vsb, "0.00")
            vsb.Left = i - 235
    End If

difveloc = (Abs(velomaxima) - Abs(velominima))
If difveloc = 0 Then
difveloc = 1
Label5.Visible = False
sp.Visible = False
vsb.Visible = False
Label11.Visible = False
vst.Visible = False
End If

If difveloc <> 0 Then difveloc = difveloc

fatescx = 5460 / difveloc


fatescy = 2160 / (((cdototal7 * Form2.Sw * 10.76 * ((dar * (velomaxima ^ 2)) / 2))))

i = (1800 + Abs(velominima)) + (Abs(velok) - Abs(velominima)) * fatescx

h = 3720 - ((cdi7 * Form2.Sw * 10.76 * qd7)) * fatescy
a1 = (cdi7 * Form2.Sw * 10.76 * qd7) / 2.2
a1 = Format(a1, "0.00")

m = 3720 - ((cd7 * Form2.Sw * 10.76 * qd7)) * fatescy
a2 = (cd7 * Form2.Sw * 10.76 * qd7) / 2.2
a2 = Format(a2, "0.00")

j = 3720 - ((cdototal7 * Form2.Sw * 10.76 * qd7)) * fatescy
at = Abs((cdi7 * Form2.Sw * 10.76 * qd7) / 2.2) + Abs((cd7 * Form2.Sw * 10.76 * qd7) / 2.2)
at = Format(at, "0.00")

'Quadro de arrastos
a1.Visible = True

a2.Visible = True

at.Visible = True
    
Label7.Visible = True
velominima.Visible = True

'graficos

grafico.PSet (i, h), &HFF&

grafico.PSet (i, j), &HFF00&

grafico.PSet (i, m), 16711935
  
Next

Label8.Visible = True
Line2.Visible = True
velomaxima.Visible = True

anguloA = Abs(Form7.alfasamax) + Abs(Form7.alfasazero)
slopesf = (Form4.clsflap / anguloA) * 57.3
slopesf = Format(slopesf, "0.00")
cl3dsf = ((slopesf / (1 + (Abs(slopesf / (3.1415 * Form1.arw))))) / 57.3) * anguloA
cl3dsf = Format(cl3dsf, "0.00")

dar = Form1.roar * 0.00194

vstk2 = ((Form1.wo * 2.2 * 2) / (dar * cl3dsf * Form2.Sw * 10.76)) ^ 0.5
vst2 = vstk2 * 1.09728
vst2 = Format(vst2, "0.00")

If Form4.semflap.Value = True Then
Label11.Visible = False
vst.Visible = False
End If

If Form4.semflap.Value = False Then
Label11.Visible = True
vst.Visible = True
Form9.vst = vst2
End If

If difveloc = 1 Then
Label11.Visible = False
vst.Visible = False
End If

End Sub
Private Sub Form_Load()
hsvinicio.Value = 4
stps.Value = 1
End Sub

Private Sub HScroll1_Change()
STPv = stps.Value / 50
End Sub

Private Sub hsvinicio_Change()
If hsvinicio.Value = 4 Then
vinicio = velominima
textvinic.Caption = "VStall"
End If

If hsvinicio.Value = 3 Then
vinicio = 0.75 * velominima
textvinic.Caption = "75% VStall"
End If

If hsvinicio.Value = 2 Then
vinicio = 0.55 * velominima
textvinic.Caption = "55% VStall"
End If

If hsvinicio.Value = 1 Then
vinicio = 0.35 * velominima
textvinic.Caption = "35% VStall"
End If

End Sub

Private Sub L1_DblClick()
L1.Height = 1230
End Sub

Private Sub stps_Change()
STPv = stps.Value / 20
End Sub
