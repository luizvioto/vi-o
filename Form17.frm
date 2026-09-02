VERSION 5.00
Begin VB.Form Form17 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualização - Estabilidade"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   ForeColor       =   &H00808080&
   Icon            =   "Form17.frx":0000
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "V"
      Height          =   255
      Left            =   7680
      TabIndex        =   24
      Top             =   0
      Width           =   495
   End
   Begin VB.ListBox LL2 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFF00&
      Height          =   1230
      ItemData        =   "Form17.frx":030A
      Left            =   240
      List            =   "Form17.frx":030C
      TabIndex        =   23
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
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   5280
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   11
         Top             =   720
         Width           =   2415
         Begin VB.Label mensagem 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   1155
            TabIndex        =   14
            Top             =   120
            Width           =   90
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   5280
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   10
         Top             =   120
         Width           =   2415
         Begin VB.Label W17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   22
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso total no CG :"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   1290
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   2760
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   9
         Top             =   720
         Width           =   2415
         Begin VB.Label Margem17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00FF00FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   21
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margem estática :"
            ForeColor       =   &H00FF00FF&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   2760
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   8
         Top             =   120
         Width           =   2415
         Begin VB.Label Lht 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   20
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sustentação Eh :"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   1230
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   7
         Top             =   720
         Width           =   2415
         Begin VB.Label PNlocal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1320
            TabIndex        =   19
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ponto Neutro :"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   1035
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00400000&
         Height          =   495
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   2355
         TabIndex        =   4
         Top             =   120
         Width           =   2415
         Begin VB.Label Lasa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   1560
            TabIndex        =   18
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sustentação  Asa :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   1350
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
         Begin VB.CommandButton Command2 
            Caption         =   "&Sair"
            Height          =   315
            Left            =   6000
            MouseIcon       =   "Form17.frx":030E
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   0
            Width           =   1400
         End
         Begin VB.CommandButton Command5 
            Caption         =   "&Margem estática"
            Height          =   315
            Left            =   2880
            MouseIcon       =   "Form17.frx":0618
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   0
            Width           =   3135
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Visualizar "
            Height          =   315
            Left            =   0
            MouseIcon       =   "Form17.frx":0922
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   0
            Width           =   2895
         End
      End
   End
   Begin VB.PictureBox pt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      Height          =   4935
      Left            =   240
      ScaleHeight     =   4875
      ScaleWidth      =   7845
      TabIndex        =   0
      Top             =   1440
      Width           =   7905
      Begin VB.Shape ca17 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Shape           =   3  'Circle
         Top             =   2400
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape cg17 
         BorderColor     =   &H00400000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Shape           =   3  'Circle
         Top             =   2400
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape pn17 
         BorderColor     =   &H00400000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Shape           =   3  'Circle
         Top             =   2400
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape catail 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Shape           =   3  'Circle
         Top             =   2400
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   7800
         Y1              =   2520
         Y2              =   2520
      End
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'LISTAR

' Incidencia da asa - usando "slope do avião"
For CLL = Form7.clk1 To Form7.clk4 Step -0.01

epson = 1 + ((2 * Form1.Afilasa) / (Form1.arw * (1 + Form1.Afilasa)))
slopeaviao = (Form7.slopeasa * 0.995) / (epson + (Form7.slopeasa / (3.141516 * Form1.arw)))

incidasa1 = (CLL / slopeaviao) + (Form7.alfasazero * 3.141516 / 180)
INCASA = incidasa1 * (180 / 3.141516)
INCASA = Format(INCASA, "0.000")


dar = 1.225 * 0.00194
dar = Format(dar, "0.0000")

VVelok = (((2 * Form1.wo * 2.2) / (dar * CLL * Form2.Sw * 10.47)) ^ 0.5) * 1.09
VVelok = Format(VVelok, "0.000")

CLL = Format(CLL, "0.000")
LL2.AddItem "   Velocidade (km/h) =  " & VVelok & vbTab & "CL =  " & CLL & vbTab & "Angulo de ataque - asa (graus)=  " & INCASA
Next

'................. AEROFOLIO DA ASA

dist = 480
k = 2520
'Cordenadas em X
L1 = 0
l2 = 37.5
l3 = 75
l4 = 150
l5 = 225
l6 = 300
l7 = 450
l8 = 600
l9 = 750
l10 = 900
l11 = 1200
l12 = 1500
l13 = 1800
l14 = 2100
l15 = 2400
l16 = 2700
l17 = 2850
l18 = 3000

'Cordenadas em Y

ya1 = 72.1
ya2 = 105.1
ya3 = 152.2
ya4 = 187.3
ya5 = 215.2
ya6 = 258.1
ya7 = 287.5
ya8 = 307.6
ya9 = 315.6
ya10 = 310.9
ya11 = 290.9
ya12 = 259
ya13 = 208.9
ya14 = 146.5
ya15 = 72.4
ya16 = 30.1
ya17 = 4.8

yb1 = -53.7
yb2 = -74.4
yb3 = -98.1
yb4 = -111.3
yb5 = -119.4
yb6 = -125.4
yb7 = -124.5
yb8 = -119.4
yb9 = -112.5
yb10 = -97.5
yb11 = -81.6
yb12 = -64.2
yb13 = -46.5
yb14 = -30.9
yb15 = -17.1
yb16 = -10.8
yb17 = -4.8

pt.DrawWidth = 1

'Linhas superiores
pt.Line (L1 + dist, k)-(l2 + dist, k - ya1), &HFFFF&
pt.Line (l2 + dist, k - ya1)-(l3 + dist, k - ya2), &HFFFF&
pt.Line (l3 + dist, k - ya2)-(l4 + dist, k - ya3), &HFFFF&
pt.Line (l4 + dist, k - ya3)-(l5 + dist, k - ya4), &HFFFF&
pt.Line (l5 + dist, k - ya4)-(l6 + dist, k - ya5), &HFFFF&
pt.Line (l6 + dist, k - ya5)-(l7 + dist, k - ya6), &HFFFF&
pt.Line (l7 + dist, k - ya6)-(l8 + dist, k - ya7), &HFFFF&
pt.Line (l8 + dist, k - ya7)-(l9 + dist, k - ya8), &HFFFF&
pt.Line (l9 + dist, k - ya8)-(l10 + dist, k - ya9), &HFFFF&
pt.Line (l10 + dist, k - ya9)-(l11 + dist, k - ya10), &HFFFF&
pt.Line (l11 + dist, k - ya10)-(l12 + dist, k - ya11), &HFFFF&
pt.Line (l12 + dist, k - ya11)-(l13 + dist, k - ya12), &HFFFF&
pt.Line (l13 + dist, k - ya12)-(l14 + dist, k - ya13), &HFFFF&
pt.Line (l14 + dist, k - ya13)-(l15 + dist, k - ya14), &HFFFF&
pt.Line (l15 + dist, k - ya14)-(l16 + dist, k - ya15), &HFFFF&
pt.Line (l16 + dist, k - ya15)-(l17 + dist, k - ya16), &HFFFF&
pt.Line (l17 + dist, k - ya16)-(l18 + dist, k - ya17), &HFFFF&
pt.Line (l18 + dist, k - ya17)-(l18 + dist, k), &HFFFF&

'Linhas inferiores
pt.Line (L1 + dist, k)-(l2 + dist, k - yb1), &HFFFF&
pt.Line (l2 + dist, k - yb1)-(l3 + dist, k - yb2), &HFFFF&
pt.Line (l3 + dist, k - yb2)-(l4 + dist, k - yb3), &HFFFF&
pt.Line (l4 + dist, k - yb3)-(l5 + dist, k - yb4), &HFFFF&
pt.Line (l5 + dist, k - yb4)-(l6 + dist, k - yb5), &HFFFF&
pt.Line (l6 + dist, k - yb5)-(l7 + dist, k - yb6), &HFFFF&
pt.Line (l7 + dist, k - yb6)-(l8 + dist, k - yb7), &HFFFF&
pt.Line (l8 + dist, k - yb7)-(l9 + dist, k - yb8), &HFFFF&
pt.Line (l9 + dist, k - yb8)-(l10 + dist, k - yb9), &HFFFF&
pt.Line (l10 + dist, k - yb9)-(l11 + dist, k - yb10), &HFFFF&
pt.Line (l11 + dist, k - yb10)-(l12 + dist, k - yb11), &HFFFF&
pt.Line (l12 + dist, k - yb11)-(l13 + dist, k - yb12), &HFFFF&
pt.Line (l13 + dist, k - yb12)-(l14 + dist, k - yb13), &HFFFF&
pt.Line (l14 + dist, k - yb13)-(l15 + dist, k - yb14), &HFFFF&
pt.Line (l15 + dist, k - yb14)-(l16 + dist, k - yb15), &HFFFF&
pt.Line (l16 + dist, k - yb15)-(l17 + dist, k - yb16), &HFFFF&
pt.Line (l17 + dist, k - yb16)-(l18 + dist, k - yb17), &HFFFF&
pt.Line (l18 + dist, k - yb17)-(l18 + dist, k), &HFFFF&

'..................AEROFOLIO DA EMPENAGEM

dist = 6120
k = 2520
'Cordenadas em X
L1 = 0
l2 = 18.75
l3 = 37.5
l4 = 75
l5 = 112.5
l6 = 150
l7 = 225
l8 = 300
l9 = 375
l10 = 450
l11 = 600
l12 = 750
l13 = 900
l14 = 1050
l15 = 1100
l16 = 1350
l17 = 1425
l18 = 1500

'Cordenadas em Y

ya1 = 14.2
ya2 = 19.6
ya3 = 26.7
ya4 = 31.5
ya5 = 35.11
ya6 = 40.09
ya7 = 43.03
ya8 = 44.45
ya9 = 45
ya10 = 43.5
ya11 = 39.7
ya12 = 34.23
ya13 = 27.48
ya14 = 19.68
ya15 = 10.86
ya16 = 6
ya17 = 0.945

yb1 = -14.2
yb2 = -19.6
yb3 = -26.7
yb4 = -31.5
yb5 = -35.11
yb6 = -40.09
yb7 = -43.03
yb8 = -44.45
yb9 = -45
yb10 = -43.5
yb11 = -39.7
yb12 = -34.23
yb13 = -27.48
yb14 = -19.68
yb15 = -10.86
yb16 = -6
yb17 = -0.945

pt.DrawWidth = 1

'Linhas superiores
pt.Line (L1 + dist, k)-(l2 + dist, k - ya1), &HFFFF&
pt.Line (l2 + dist, k - ya1)-(l3 + dist, k - ya2), &HFFFF&
pt.Line (l3 + dist, k - ya2)-(l4 + dist, k - ya3), &HFFFF&
pt.Line (l4 + dist, k - ya3)-(l5 + dist, k - ya4), &HFFFF&
pt.Line (l5 + dist, k - ya4)-(l6 + dist, k - ya5), &HFFFF&
pt.Line (l6 + dist, k - ya5)-(l7 + dist, k - ya6), &HFFFF&
pt.Line (l7 + dist, k - ya6)-(l8 + dist, k - ya7), &HFFFF&
pt.Line (l8 + dist, k - ya7)-(l9 + dist, k - ya8), &HFFFF&
pt.Line (l9 + dist, k - ya8)-(l10 + dist, k - ya9), &HFFFF&
pt.Line (l10 + dist, k - ya9)-(l11 + dist, k - ya10), &HFFFF&
pt.Line (l11 + dist, k - ya10)-(l12 + dist, k - ya11), &HFFFF&
pt.Line (l12 + dist, k - ya11)-(l13 + dist, k - ya12), &HFFFF&
pt.Line (l13 + dist, k - ya12)-(l14 + dist, k - ya13), &HFFFF&
pt.Line (l14 + dist, k - ya13)-(l15 + dist, k - ya14), &HFFFF&
pt.Line (l15 + dist, k - ya14)-(l16 + dist, k - ya15), &HFFFF&
pt.Line (l16 + dist, k - ya15)-(l17 + dist, k - ya16), &HFFFF&
pt.Line (l17 + dist, k - ya16)-(l18 + dist, k - ya17), &HFFFF&
pt.Line (l18 + dist, k - ya17)-(l18 + dist, k), &HFFFF&

'Linhas inferiores
pt.Line (L1 + dist, k)-(l2 + dist, k - yb1), &HFFFF&
pt.Line (l2 + dist, k - yb1)-(l3 + dist, k - yb2), &HFFFF&
pt.Line (l3 + dist, k - yb2)-(l4 + dist, k - yb3), &HFFFF&
pt.Line (l4 + dist, k - yb3)-(l5 + dist, k - yb4), &HFFFF&
pt.Line (l5 + dist, k - yb4)-(l6 + dist, k - yb5), &HFFFF&
pt.Line (l6 + dist, k - yb5)-(l7 + dist, k - yb6), &HFFFF&
pt.Line (l7 + dist, k - yb6)-(l8 + dist, k - yb7), &HFFFF&
pt.Line (l8 + dist, k - yb7)-(l9 + dist, k - yb8), &HFFFF&
pt.Line (l9 + dist, k - yb8)-(l10 + dist, k - yb9), &HFFFF&
pt.Line (l10 + dist, k - yb9)-(l11 + dist, k - yb10), &HFFFF&
pt.Line (l11 + dist, k - yb10)-(l12 + dist, k - yb11), &HFFFF&
pt.Line (l12 + dist, k - yb11)-(l13 + dist, k - yb12), &HFFFF&
pt.Line (l13 + dist, k - yb12)-(l14 + dist, k - yb13), &HFFFF&
pt.Line (l14 + dist, k - yb13)-(l15 + dist, k - yb14), &HFFFF&
pt.Line (l15 + dist, k - yb14)-(l16 + dist, k - yb15), &HFFFF&
pt.Line (l16 + dist, k - yb15)-(l17 + dist, k - yb16), &HFFFF&
pt.Line (l17 + dist, k - yb16)-(l18 + dist, k - yb17), &HFFFF&
pt.Line (l18 + dist, k - yb17)-(l18 + dist, k), &HFFFF&

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

'............... CA da ASA
For kcg = 0 To 1230 Step 0.25
        ca17.Visible = True
        ca17.Left = kcg - (ca17.Width / 2)
Next

escay = 2400 / Form10.sustentasa

x1 = ca17.Left + (ca17.Width / 2)
X2 = x1
altseta = 2520 - Abs(Form10.sustentasa * escay) 'Vertical
setesq = 2520 - (Abs(Form10.sustentasa * escay) * 0.8) 'Seta esquerda
setdir = 2520 - (Abs(Form10.sustentasa * escay) * 0.8) 'Seta direita
pt.Line (x1, 2520)-(X2, altseta), &HFFFF&
pt.Line (0.92 * x1, setesq)-(x1, altseta), &HFFFF&
pt.Line (1.08 * x1, setdir)-(x1, altseta), &HFFFF&
Form17.Lasa = Form10.sustentasa + " kgf"

'............... CA da Empenagem horizontal

If Form10.sustentaht = 0 Then Form10.sustentaht = 0.0001
escahty = 600 / Form10.sustentaht

If Val(Form10.cg) > Val(Form10.PN) Then
    For kcg = 0 To 6552 Step 0.25
        catail.Visible = True
        catail.Left = kcg - (catail.Width / 2)
    Next

            x1 = catail.Left + (catail.Width / 2)
            X2 = x1
            altsetail = 2520 - Abs(Form10.sustentaht * escahty) 'Vertical
            setesqtail = 2520 - (Abs(Form10.sustentaht * escahty) * 0.8) 'Seta esquerda
            setdirtail = 2520 - (Abs(Form10.sustentaht * escahty) * 0.8) 'Seta direita
            pt.Line (x1, 2520)-(X2, altsetail), &H80FF&
            pt.Line (0.99 * x1, setesqtail)-(x1, altsetail), &H80FF&
            pt.Line (1.01 * x1, setesqtail)-(x1, altsetail), &H80FF&
End If

If Val(Form10.cg) < Val(Form10.PN) Then

    For kcg = 0 To 6552 Step 0.25
        catail.Visible = True
        catail.Left = kcg - (catail.Width / 2)
    Next

            x1 = catail.Left + (catail.Width / 2)
            X2 = x1
            altsetail = 2520 + Abs(Form10.sustentaht * escahty) 'Vertical
            setesqtail = 2520 + (Abs(Form10.sustentaht * escahty) * 0.8) 'Seta esquerda
            setdirtail = 2520 + (Abs(Form10.sustentaht * escahty) * 0.8) 'Seta direita
            pt.Line (x1, 2520)-(X2, altsetail), &H80FF&
            pt.Line (0.99 * x1, setesqtail)-(x1, altsetail), &H80FF&
            pt.Line (1.01 * x1, setdirtail)-(x1, altsetail), &H80FF&
End If

If Form10.cg = 25 Then

    For kcg = 0 To 6552 Step 0.25
        catail.Visible = True
        catail.Left = kcg - (catail.Width / 2)
    Next

End If
Form17.Lht = Form10.sustentaht + " kgf"

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'............... CG

escapeso = 2400 / Form1.wo
locacg = 480 + (Form10.cg / 100 * 3000)
cg17.Visible = True
For kcg = 0 To locacg Step 0.05
    cg17.Left = kcg - (cg17.Width / 2)
Next

x1 = cg17.Left + (cg17.Width / 2)
X2 = x1
altseta = 2520 + Abs(Form1.wo * escapeso) 'Vertical
setesq = 2520 + (Abs(Form1.wo * escapeso) * 0.8) 'Seta esquerda
setdir = 2520 + (Abs(Form1.wo * escapeso) * 0.8) 'Seta direita

pt.Line (x1, 2520)-(X2, altseta), &HFF&
pt.Line (0.91 * x1, setesq)-(x1, altseta), &HFF&
pt.Line (1.09 * x1, setdir)-(x1, altseta), &HFF&

Form17.W17 = Form1.wo + " kgf"


'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

'............... PN
locapn = 480 + (Form10.PN / 100 * 3000)
pn17.Visible = True
For kcg = 0 To locapn Step 0.1
    pn17.Left = kcg - (pn17.Width / 2)
Next

x1 = pn17.Left + (pn17.Width / 2)
X2 = x1

pt.Line (x1, 0)-(X2, 5055), &HFFFFFF

Form17.PNlocal = Form10.PN + " %"

'xxxxxxxxxxxxxxx Mensagem xxxxxxxxxxxxxxxxxxxxx
If Val(Form10.cg) > Val(Form10.PN) Then Form17.mensagem = "Cuidado CG atrasado!"
If Val(Form10.cg) < Val(Form10.PN) And Val(Form10.cg) < 12 Then Form17.mensagem = "Cuidado CG adiantado!"
If Val(Form10.cg) < Val(Form10.PN) And Val(Form10.cg) > 12 And Val(Form10.cg) <= 15 Then Form17.mensagem = "Ok! Nariz pesado!"
If Val(Form10.cg) < Val(Form10.PN) And Val(Form10.cg) >= 15.01 And Val(Form10.cg) <= 18 Then Form17.mensagem = "Ok! Bom treinador!"
If Val(Form10.cg) < Val(Form10.PN) And Val(Form10.cg) >= 18.01 And Val(Form10.cg) <= 22 Then Form17.mensagem = "Ok! CG dentro da faixa!"
If Val(Form10.cg) < Val(Form10.PN) And Val(Form10.cg) >= 22.01 And Val(Form10.cg) <= 25 Then Form17.mensagem = "Ok! Bom acrobático!"
If Val(Form10.cg) < Val(Form10.PN) And Val(Form10.cg) >= 25.01 And Val(Form10.cg) <= 30 Then Form17.mensagem = "Ok! Verifique CG vazio!"
If Val(Form10.cg) < Val(Form10.PN) And Val(Form10.cg) >= 25.01 And Val(Form10.cg) < Val(Form10.PN) Then Form17.mensagem = "Ok! Verifique CG vazio!"
If Val(Form10.cg) = Val(Form10.PN) Then Form17.mensagem = "Cuidado! CG marginal!"

End Sub

Private Sub Command2_Click()

Form17.pt.Cls

catail.Visible = False
pn17.Visible = False
cg17.Visible = False
ca17.Visible = False
Form17.Visible = False
End Sub


Private Sub Command4_Click()
LL2.Height = 6105
End Sub

Private Sub Command5_Click()

pt.DrawWidth = 2

x1 = cg17.Left + (cg17.Width / 2)
X2 = pn17.Left + (cg17.Width / 2)

For mest = x1 To X2 Step 0.5

kmest = mest
 
    pt.Line (x1, 3240)-(kmest, 3240), &HFF00FF
    
Next

Form17.Margem17 = Form10.MARGEM + " %"
End Sub




Private Sub LL2_DblClick()
LL2.Height = 1230
End Sub
