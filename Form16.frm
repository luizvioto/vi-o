VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualização - Potencias"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   ForeColor       =   &H00808080&
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LL3 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFF00&
      Height          =   1230
      ItemData        =   "Form16.frx":030A
      Left            =   240
      List            =   "Form16.frx":030C
      TabIndex        =   30
      Top             =   240
      Width           =   7905
   End
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1755
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   6480
      Width           =   7935
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   5280
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   23
         Top             =   720
         Width           =   2415
         Begin VB.Label torque16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1440
            TabIndex        =   29
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Torque do motor :"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   2760
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   22
         Top             =   720
         Width           =   2415
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "h"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   14.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   120
            TabIndex        =   31
            Top             =   0
            Width           =   165
         End
         Begin VB.Label NI16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "."
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1800
            TabIndex        =   28
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Hélice para Vmax :"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   120
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   21
         Top             =   720
         Width           =   2415
         Begin VB.Label rpm16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "."
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1320
            TabIndex        =   27
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RPM utilizada :"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   1065
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   2760
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   11
         Top             =   120
         Width           =   2414
         Begin VB.Label pdisp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00FF00FF&
            Height          =   195
            Left            =   1800
            TabIndex        =   13
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Potencia Disponível :"
            ForeColor       =   &H00FF00FF&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   5280
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   8
         Top             =   120
         Width           =   2415
         Begin VB.Label preqr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00FFFF00&
            Height          =   195
            Left            =   1800
            TabIndex        =   10
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Potencia Requerida :"
            ForeColor       =   &H00FFFF00&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   5
         Top             =   120
         Width           =   2415
         Begin VB.Label pnom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   1680
            TabIndex        =   7
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Potencia Nominal :"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   7395
         TabIndex        =   2
         Top             =   1320
         Width           =   7455
         Begin VB.CommandButton Command3 
            Caption         =   "&Sair"
            Height          =   315
            Left            =   3720
            MouseIcon       =   "Form16.frx":030E
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   0
            Width           =   3675
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Visualizar Potencias"
            Height          =   315
            Left            =   0
            MouseIcon       =   "Form16.frx":0618
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   0
            Width           =   3735
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   4935
      Left            =   240
      ScaleHeight     =   4875
      ScaleWidth      =   7845
      TabIndex        =   0
      Top             =   1440
      Width           =   7905
      Begin VB.TextBox hp2 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "-"
         Top             =   3240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox hp1 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Text            =   "-"
         Top             =   3600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox hpreq 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Text            =   "-"
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   600
         X2              =   600
         Y1              =   240
         Y2              =   4560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line vs 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   1680
         X2              =   1680
         Y1              =   360
         Y2              =   4680
      End
      Begin VB.Label vstall 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1440
         TabIndex        =   17
         Top             =   4680
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Line vline 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   7200
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vel."
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
         Left            =   7320
         TabIndex        =   16
         Top             =   4200
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hp"
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
         TabIndex        =   15
         Top             =   0
         Width           =   255
      End
      Begin VB.Line Line45 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line20 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label velkm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   4680
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Line Line44 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   2760
         X2              =   2760
         Y1              =   4200
         Y2              =   360
      End
      Begin VB.Line v1 
         BorderColor     =   &H00FFFF00&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   120
         X2              =   120
         Y1              =   4560
         Y2              =   360
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   600
         X2              =   7080
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   1560
         X2              =   1560
         Y1              =   4200
         Y2              =   360
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   2040
         X2              =   2040
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   2280
         X2              =   2280
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   2520
         X2              =   2520
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   3000
         X2              =   3000
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   3240
         X2              =   3240
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line24 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   3480
         X2              =   3480
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line25 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   3720
         X2              =   3720
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   3960
         X2              =   3960
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   4200
         X2              =   4200
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   4440
         X2              =   4440
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line29 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   4680
         X2              =   4680
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line30 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   4920
         X2              =   4920
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line31 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   5160
         X2              =   5160
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line32 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   5400
         X2              =   5400
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line33 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   5640
         X2              =   5640
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   5880
         X2              =   5880
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line35 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   6120
         X2              =   6120
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line36 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   6360
         X2              =   6360
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line37 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   6600
         X2              =   6600
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line38 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   6840
         X2              =   6840
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line39 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   7080
         X2              =   7080
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line40 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   1320
         X2              =   1320
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line41 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   1080
         X2              =   1080
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line42 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   840
         X2              =   840
         Y1              =   360
         Y2              =   4200
      End
      Begin VB.Line Line43 
         BorderColor     =   &H00008080&
         BorderStyle     =   3  'Dot
         X1              =   1800
         X2              =   1800
         Y1              =   360
         Y2              =   4200
      End
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture1.Refresh

For velocidade = Form1.vstol To Form8.vmax * 1.1 Step 0.5

potreq = ((Form8.FD * 2.2) * Form8.vmax / 1.61) / 375

'Formula do Nicola ....... v = ((6480 / Form8.vmax) * velocidade)

'Escala em X
If Form1.motor < 2 Then v = velocidade * 30
If Form1.motor > 2.01 Then v = velocidade * 20
'Escala em Y
fatescy = 3840 / Form1.motor

'LINHA HPnominal
hpnominal = 4200 - (Form8.hpnominal * (fatescy))
Form16.Picture1.PSet (v, hpnominal), &HFF00&
hp1 = Form8.hpnominal
hp1.Top = hpnominal - 80

'LINHA HPdisponivel
hpdisponivel = 4200 - (Form8.hpreal * (fatescy))
Form16.Picture1.PSet (v, hpdisponivel), &HFF00FF
hp2 = Form8.hpreal
hp2.Top = hpdisponivel - 80


            'LINHA HPrequerido "NAO NECESSARIA POIS HPreq=TRACAO"
            'hp = 4080 - ((potreq) * 800)
            'Form16.Picture1.PSet (v, hp), &HFFFF00

pnom = Form8.hpnominal
pnom = Format(pnom, "0.00")

pdisp = Form8.hpreal
pdisp = Format(pdisp, "0.00")

preqr = potreq
preqr = Format(preqr, "0.00")

Next

 '*************** 'PARA GRAFICO DE HPrequerida *****************

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
If Requisitos.Tubalum = True Then fatpaint = 0.00035

                'fuselagem
reyf7 = 38.21 * (Form2.lfs * 3.28 / fatpaint) ^ 1.053
reyf7 = Format(reyf7, "0.0000")
lfsdia = (Form2.lfs * 3.28) / (((Form2.wfs * 3.28) * (Form2.hfs * 3.28)) ^ 0.5)
gf7 = Log(reyf7) / Log(10)

                'asa
reyw7 = 38.21 * (((Form7.CMaerodw * 3.28) / fatpaint) ^ 1.053)
reyw7 = Format(reyw7, "0.0000")
gw7 = Log(reyw7) / Log(10)

                'empenagem horizontal
reyht7 = 38.21 * (Form2.cmht * (3.28) / fatpaint) ^ 1.053
reyht7 = Format(reyht7, "0.0000")
ght7 = Log(reyht7) / Log(10)

                'empenagem vertical
reyvt7 = 38.21 * (Form2.cmvt * (3.28) / fatpaint) ^ 1.053
reyvt7 = Format(reyvt7, "0.0000")
gvt7 = Log(reyvt7) / Log(10)

'Arrasto trem de pouso etc

If Form1.Checkretratil.Value = 1 Then cdotp = 0
If Form1.Checkretratil.Value = 0 Then

afpneu = Abs((Form1.diambeq * Form1.espbeq) * 0.001076) + Abs(((Form1.diamtp * Form1.esptp) * 0.001076) * 2)
dqpneu = afpneu * 0.25

cdotp = 1.2 * (dqpneu) / (Form2.Sw * 10.76)
cdotp = Format(cdotp, "0.0000")

End If

'Arrasto de áreas adicionadas suplementares

If Form1.adicionarea = 0 Then adicionarea = "0.0000"

If Form1.adicionarea <> 0 Then
dqoutras = (Form1.adicionarea * 10.76) * 0.1
cdosuple = 1.2 * (dqoutras) / (Form2.Sw * 10.76)
cdosuple = Format(cdosuple, "0.0000")
End If
                                                  
'.......................ARRASTO...................................

'VELOCIDADES
'velok é dado em km/h
STP = 0.5

vteste = Int(Form8.vmax)
vteste = Format(vteste, "0.00")

'End If
'LOOP utilizando velocidades já calculadas anteriormente
For velok = Form1.vstol To vteste * 1.1 Step STP

'km/h para mph
velokmph = velok / 1.61

'km/h para ft/s
velokpe = velok * 0.911

'Em Mach
velokmach = velok / 1224
velokmach = Format(velokmach, "0.0000")

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
Form7.swetasa = Format(Form7.swetasa, "0.0000")


swetht7 = ((Form2.sht * 10.76) * (1.977 + (0.52 * (Form2.espesht / 100)))) * 0.8
Form7.swethorz = swetht7 / 10.76
Form7.swethorz = Format(Form7.swethorz, "0.0000")


swetvt7 = ((Form2.svt * 10.76) * (1.977 + (0.52 * (Form2.espesvt / 100)))) * 0.8
Form7.swetvert = swetvt7 / 10.76
Form7.swetvert = Format(Form7.swetvert, "0.0000")


'Arrasto fuselagem
cff7 = 0.455 / ((gf7 ^ 2.58) * (1 + (0.144 * (velokmach)) ^ 2) ^ 0.65)
cff7 = Format(cff7, "0.0000")

fff7 = 1 + (60 / lfsdia ^ 3) + (lfsdia / 400)
fff7 = Format(fff7, "0.0000")

cdof7 = (cff7 * fff7 * swetf7) / (Form2.Sw * 10.76)
cdof7 = Format(cdof7, "0.0000")

'Arrasto asa

cfw7 = 0.455 / ((gw7 ^ 2.58) * (1 + (0.144 * (velokmach)) ^ 2) ^ 0.65)
cfw7 = Format(cfw7, "0.0000")

ffw7 = ((1 + ((0.6 / 0.25) * (Form2.espessura / 100)) + (100 * ((Form2.espessura / 100) ^ 4))) * (1.34 * ((velokmach) ^ 0.18)))
ffw7 = Format(ffw7, "0.0000")
'Cdo stall asa
cdow7 = (cfw7 * ffw7 * swetw7) / (Form2.Sw * 10.76)
cdow7 = Format(cdow7, "0.0000")

'Arrasto da empenagem horizontal

cfht7 = 0.455 / ((ght7 ^ 2.58) * (1 + (0.144 * (velokmach)) ^ 2) ^ 0.65)
cfht7 = Format(cfht7, "0.0000")

ffht7 = ((1 + ((0.6 / 0.25) * (Form2.espesht / 100)) + (100 * ((Form2.espesht / 100) ^ 4))) * (1.34 * ((velokmach) ^ 0.18)))
ffht7 = Format(ffht7, "0.0000")
'Cdo stall emp horz
cdoht7 = ((cfht7 * ffht7 * swetht7) / (Form2.Sw * 10.76))
cdoht7 = Format(cdoht7, "0.0000")

'Arrasto da empenagem vertical

cfvt7 = 0.455 / ((gvt7 ^ 2.58) * (1 + (0.144 * (velokmach)) ^ 2) ^ 0.65)
cfvt7 = Format(cfvt7, "0.0000")

ffvt7 = ((1 + ((0.6 / 0.25) * (Form2.espesvt / 100)) + (100 * ((Form2.espesvt / 100) ^ 4))) * (1.34 * ((velokmach) ^ 0.18)))
ffvt7 = Format(ffvt7, "0.0000")
'Cdo stall emp vert
cdovt7 = ((cfvt7 * ffvt7 * swetvt7) / (Form2.Sw * 10.76))
cdovt7 = Format(cdoht7, "0.0000")

'Arrasto Cooling e motor
If velok < 100 Then cdocool = 0
If velk < 100 Then cdomotor = 0
If velok >= 100 Then
'dqcool = ((4.9 * 10 ^ -7) * Form8.hpreq * (trank ^ 2)) / (velokpe)
'dqmotor = ((2 * 10 ^ -4) * Form8.hpreq)

cdocool = ((2.5) * dqcool) / (Form2.Sw * 10.76)
cdocool = Format(cdocool, "0.0000")

cdomotor = dqmotor / (Form2.Sw * 10.76)
cdomotor = Format(cdomotor, "0.0000")
End If

'ARRASTO PARASITA TOTAL
'arrasto de interferencia cdif=1.10
cdif = 1.05
'Com arrasto do Perfil
cdototal7 = cdif * ((Abs(cdof7) + Abs(cdow7) + Abs(Form7.cdperfasa) + Abs(cdoht7) + Abs(cdovt7) + Abs(cdotp)) + Abs(cdocool) + Abs(cdomotor) + Abs(cdosuple))
cdototal7 = Format(cdototal7, "0.0000")

'Fim de arrasto PARASITA ideal

'ARRASTO INDUZIDO
dar = Form1.roar * 0.00194

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
cd7 = cdif * ((Abs(cdof7) + Abs(cdow7) + Abs(Form7.cdperfasa) + Abs(cdoht7) + Abs(cdovt7) + Abs(cdotp)) + Abs(cdocool) + Abs(cdomotor) + Abs(cdosuple)) + Abs((Abs(k) * (cl7 ^ 2)))
cd7 = Format(cd7, "0.0000")

m = (cd7 * Form2.Sw * 10.76 * qd7) / 2.2
m = Format(m, "0.0000")

'BHP = F x V
bhp = 4200 - (((m * 2.2) * (velok / 1.61) / 375) * (fatescy))
bhp = Format(bhp, "0.0000")


hpreq = ((m * 2.2) * (Form1.vstol / 1.61) / 375)
hpreq = Format(hpreq, "0.00")
hpreq.Top = 4080 - (hpreq * fatescy)


'Velocidade em potencia disponível
'Form16.vpd = Form8.vmax / 1.61

'graficos
v = ((6480 / Form8.vmax) * velok)
v = Format(v, "0.0000")

If Form1.motor < 2 Then vreq = velok * 30
If Form1.motor > 2.01 Then vreq = velok * 20

VVelok = velok
VVelok = Format(VVelok, "0.00")

HPk = (((m * 2.2) * (velok * 0.911)) / 375)
HPk = Format(HPk, "0.00")

NI16 = (HPk / Form16.hp2) - 1
NI16 = Format(NI16, "0.00")
rpm16 = Form1.RPM

LL3.AddItem "   Velocidade (km/h) =  " & VVelok & vbTab & vbTab & "Potência Instantanea =  " & HPk '& vbTab & "Incidência da asa (graus)=  " & INCASA


Form16.Picture1.PSet (vreq, bhp), &HFFFF00


vs.x1 = Form1.vstol * 20
vs.X2 = Form1.vstol * 20

vstall.Left = Form1.vstol * 20
vstall = Form1.vstol

    
vs.Visible = True
vstall.Visible = True
v1.x1 = vreq
v1.X2 = vreq
velkm = Form8.vmax
velkm = Format(velkm, "0.00")
velkm.Left = vreq
hpreq.Visible = True
hp1.Visible = True
hp2.Visible = True
Next

velkm.Visible = True
v1.Visible = True




End Sub


Private Sub Command3_Click()
Form16.Visible = False
velkm.Visible = False
v1.Visible = False
vs.Visible = False
vstall.Visible = False
hpreq.Visible = False
hp1.Visible = False
hp2.Visible = False

End Sub




