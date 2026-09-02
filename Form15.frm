VERSION 5.00
Begin VB.Form Form15 
   BackColor       =   &H00808080&
   Caption         =   "Visualização - Peso e balanceamento"
   ClientHeight    =   7050
   ClientLeft      =   300
   ClientTop       =   465
   ClientWidth     =   8445
   ForeColor       =   &H00808080&
   Icon            =   "Form15.frx":0000
   LinkTopic       =   "Form15"
   ScaleHeight     =   7050
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   5415
      Left            =   240
      ScaleHeight     =   5355
      ScaleWidth      =   7845
      TabIndex        =   1
      Top             =   120
      Width           =   7905
      Begin VB.Frame Frame1 
         BackColor       =   &H00400000&
         Caption         =   "&Distancias na corda média da asa"
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   4200
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   3495
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cm"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   3120
            TabIndex        =   30
            Top             =   1800
            Width           =   210
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cm"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   3120
            TabIndex        =   29
            Top             =   1320
            Width           =   210
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cm"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   3120
            TabIndex        =   28
            Top             =   840
            Width           =   210
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cm"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   3120
            TabIndex        =   27
            Top             =   360
            Width           =   210
         End
         Begin VB.Label caba 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2640
            TabIndex        =   26
            Top             =   1800
            Width           =   45
         End
         Begin VB.Label pnba 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2640
            TabIndex        =   25
            Top             =   1320
            Width           =   45
         End
         Begin VB.Label cgvba 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2640
            TabIndex        =   24
            Top             =   840
            Width           =   45
         End
         Begin VB.Label cgtba 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2640
            TabIndex        =   23
            Top             =   360
            Width           =   45
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bordo de ataque ao Centro Aerod.:"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1800
            Width           =   2475
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bordo de ataque ao Ponto Neutro :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   2490
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bordo de ataque ao CG s/comb :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   2355
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bordo de ataque ao CGtotal :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   2070
         End
      End
      Begin VB.TextBox beqconv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   210
         Left            =   6840
         MouseIcon       =   "Form15.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   5
         ToolTipText     =   "Peso na Bequilha"
         Top             =   3735
         Width           =   735
      End
      Begin VB.TextBox beqtri 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   210
         Left            =   840
         MouseIcon       =   "Form15.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "Peso na Bequilha"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox re 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         MouseIcon       =   "Form15.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   3
         ToolTipText     =   "Peso na Roda Esquerda"
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox rd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   3240
         MouseIcon       =   "Form15.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   2
         ToolTipText     =   "Peso na Roda Direita"
         Top             =   2400
         Width           =   735
      End
      Begin VB.Shape cgvazio15 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape cgtotal15 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape pn15 
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   240
         Shape           =   3  'Circle
         Top             =   3720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label actext 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "actext"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   1745
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Line lnac 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   480
         X2              =   360
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label pntext 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "pntext"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label cgvtext 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "cgvtext"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label cgtext 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "cgtext"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Line lcgv 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   360
         X2              =   500
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line lpn 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   360
         X2              =   480
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line lcg 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   360
         X2              =   600
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line ln55 
         BorderColor     =   &H0000FFFF&
         X1              =   7080
         X2              =   6360
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line ln11 
         X1              =   6360
         X2              =   7080
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line ln33 
         BorderColor     =   &H0000FFFF&
         X1              =   7080
         X2              =   7080
         Y1              =   3840
         Y2              =   2640
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000FFFF&
         X1              =   2760
         X2              =   4080
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line ln5 
         BorderColor     =   &H0000FFFF&
         X1              =   2760
         X2              =   4080
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line ln2 
         BorderColor     =   &H0000FFFF&
         X1              =   2760
         X2              =   4080
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line ln3 
         BorderColor     =   &H0000FFFF&
         X1              =   4080
         X2              =   4080
         Y1              =   3840
         Y2              =   120
      End
      Begin VB.Line ln4 
         BorderColor     =   &H0000FFFF&
         X1              =   2760
         X2              =   2760
         Y1              =   3840
         Y2              =   120
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   7800
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line ln1 
         BorderColor     =   &H0000FFFF&
         Visible         =   0   'False
         X1              =   2760
         X2              =   4080
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Shape caerod 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   135
         Left            =   2700
         Shape           =   3  'Circle
         Top             =   1800
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Line lineac 
         BorderColor     =   &H0000FFFF&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   3120
         X2              =   3135
         Y1              =   3840
         Y2              =   3855
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   2  'Dash
         X1              =   360
         X2              =   360
         Y1              =   2640
         Y2              =   5040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   360
         X2              =   360
         Y1              =   5280
         Y2              =   0
      End
      Begin VB.Line Line8 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   5760
         X2              =   6840
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         Height          =   495
         Left            =   240
         Shape           =   2  'Oval
         Top             =   3600
         Width           =   6975
      End
      Begin VB.Line ln66 
         BorderColor     =   &H0000FFFF&
         X1              =   7080
         X2              =   7080
         Y1              =   3840
         Y2              =   4920
      End
      Begin VB.Line linecorda 
         BorderColor     =   &H0000FFFF&
         Visible         =   0   'False
         X1              =   2760
         X2              =   4080
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line ln77 
         BorderColor     =   &H0000FFFF&
         X1              =   6360
         X2              =   6360
         Y1              =   3840
         Y2              =   4920
      End
      Begin VB.Line ln44 
         BorderColor     =   &H0000FFFF&
         X1              =   6360
         X2              =   6360
         Y1              =   3840
         Y2              =   2640
      End
      Begin VB.Line ln22 
         BorderColor     =   &H0000FFFF&
         X1              =   6360
         X2              =   7080
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line ln7 
         BorderColor     =   &H0000FFFF&
         X1              =   2760
         X2              =   2760
         Y1              =   3840
         Y2              =   7560
      End
      Begin VB.Line ln6 
         BorderColor     =   &H0000FFFF&
         X1              =   4080
         X2              =   4080
         Y1              =   3840
         Y2              =   7560
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   5640
      Width           =   7935
      Begin VB.PictureBox Picture3 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   7635
         TabIndex        =   6
         Top             =   600
         Width           =   7695
         Begin VB.CommandButton Command3 
            Caption         =   "&Sair"
            Height          =   455
            Left            =   6000
            MouseIcon       =   "Form15.frx":0F32
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   0
            Width           =   1650
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Visualizar pontos"
            Height          =   455
            Left            =   3000
            MouseIcon       =   "Form15.frx":123C
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   0
            Width           =   3015
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Configurar o avião"
            Height          =   455
            Left            =   0
            MouseIcon       =   "Form15.frx":1546
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centro aerodinâmico"
         Height          =   195
         Left            =   6240
         TabIndex        =   16
         Top             =   175
         Width           =   1455
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   6000
         Shape           =   3  'Circle
         Top             =   205
         Width           =   135
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   4320
         Shape           =   3  'Circle
         Top             =   205
         Width           =   135
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   205
         Width           =   135
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   120
         Shape           =   3  'Circle
         Top             =   205
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ponto Neutro"
         Height          =   195
         Left            =   4560
         TabIndex        =   15
         Top             =   175
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CG sem combustível"
         Height          =   195
         Left            =   2400
         TabIndex        =   14
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CG com peso total"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   175
         Width           =   1305
      End
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'PARA A ASA

k = (1 - (Form1.Afilasa))
z = (ln2.X2 - ln1.x1)
ln2.x1 = ((k / 2) * z) + ln1.x1
ln2.X2 = ln2.x1 + (Form1.Afilasa * z)
ln4.X2 = ln2.x1
ln3.X2 = ln2.X2

ln5.x1 = ((k / 2) * z) + ln1.x1
ln5.X2 = ln5.x1 + (Form1.Afilasa * z)
ln6.X2 = ln5.X2
ln7.X2 = ln5.x1

'linha média
k1 = (1320 - (((1 + Form1.Afilasa) / 2) * (linecorda.X2 - linecorda.x1))) / 2

linecorda.Visible = True
linecorda.x1 = linecorda.x1 + k1
linecorda.X2 = linecorda.X2 - k1

'PARA O ESTABILIZADOR
k = (1 - (Form1.AfilaHt))
z = (ln22.X2 - ln11.x1)
ln22.x1 = ((k / 2) * z) + ln11.x1
ln22.X2 = ln22.x1 + (Form1.AfilaHt * z)
ln44.X2 = ln22.x1
ln33.X2 = ln22.X2

ln55.x1 = ((k / 2) * z) + ln11.x1
ln55.X2 = ln55.x1 + (Form1.AfilaHt * z)
ln66.X2 = ln55.X2
ln77.X2 = ln55.x1

'kline = ln2.X2 - ln2.X1
'k2line = (ln1.X2 - ln1.X1) * lda

'CEntro aerodinamico
lineac.Visible = True
lineac.y1 = 120
lineac.x1 = ln2.x1 + (0.277 * (ln2.X2 - ln2.x1))
caerod.Visible = True
caerod.Left = (linecorda.x1 + (0.25 * (linecorda.X2 - linecorda.x1))) - 30
actext.Visible = True
lnac.Visible = True
actext.Left = caerod.Left / 2
actext = Form10.LRCAasa * 100
lnac.X2 = caerod.Left
lnac.Visible = True
Command1.Enabled = False

End Sub

Private Sub Command2_Click()

distLRBAcma = Abs(Form10.LRCAasa * 100) - Abs((0.25 * Form7.CMaerodw * 100))

'Posição do PN
For kcg = 0 To linecorda.x1 - 35 Step 0.5
pn15.Visible = True
pn15.Left = kcg + ((Form14.PN / 100) * (linecorda.X2 - linecorda.x1))
Next

For kcg1 = 0 To 2100 Step 1
pn15.Top = 3840 - kcg1
lpn.Visible = True
lpn.X2 = pn15.Left
Next
pntext = Abs(((Form14.PN / 100) * Form7.CMaerodw) * 100) + Abs(distLRBAcma)
pntext = Format(pntext, "0.00")
pntext.Left = lpn.X2 / 2
pntext.Visible = True

'Posição do CGtotal
For kcg = 0 To linecorda.x1 - 35 Step 0.5
cgtotal15.Visible = True
cgtotal15.Left = kcg + ((Form14.CGestatico / 100) * (linecorda.X2 - linecorda.x1))
Next

For kcg1 = 0 To 2100 Step 1
cgtotal15.Top = 3840 - kcg1
lcg.Visible = True
lcg.X2 = cgtotal15.Left
Next
cgtext = Abs(((Form14.CGestatico / 100) * Form7.CMaerodw) * 100) + Abs(distLRBAcma)
cgtext = Format(cgtext, "0.00")
cgtext.Left = lcg.X2 / 2
cgtext.Visible = True


'Posição do CGvazio
For kcg = 0 To linecorda.x1 - 35 Step 0.5
cgvazio15.Visible = True
cgvazio15.Left = kcg + ((Form14.CGvazio / 100) * (linecorda.X2 - linecorda.x1))
Next

For kcg1 = 0 To 2100 Step 1
cgvazio15.Top = 3840 - kcg1
lcgv.Visible = True
lcgv.X2 = cgvazio15.Left
Next
cgvtext = Abs(((Form14.CGvazio / 100) * Form7.CMaerodw) * 100) + Abs(distLRBAcma)
cgvtext = Format(cgvtext, "0.00")
cgvtext.Left = lcgv.X2 / 2
cgvtext.Visible = True


cgtba = ((Form14.CGestatico / 100) * Form7.CMaerodw) * 100
cgtba = Format(cgtba, "0.00")

cgvba = ((Form14.CGvazio / 100) * Form7.CMaerodw) * 100
cgvba = Format(cgvba, "0.00")

pnba = ((Form14.PN / 100) * Form7.CMaerodw) * 100
pnba = Format(pnba, "0.00")

caba = (0.25 * Form7.CMaerodw) * 100
caba = Format(caba, "0.00")

Frame1.Visible = True
End Sub

Private Sub Command3_Click()
Frame1.Visible = False
Form15.Visible = False
Command1.Enabled = True
lineac.y1 = 3840
lineac.y2 = 3840
linecorda.Visible = False
lineac.Visible = True
caerod.Visible = False
caerod.Left = 0
lnac.Visible = False
'VOLTA AO ORIGINAL DA ASA
ln1.x1 = 2760
ln1.X2 = 4080
ln1.y1 = 3840
ln1.y2 = 3840

ln2.x1 = 2760
ln2.X2 = 4080
ln2.y1 = 120
ln2.y2 = 120

ln3.x1 = 4080
ln3.X2 = 4080
ln3.y1 = 3840
ln3.y2 = 120

ln4.x1 = 2760
ln4.X2 = 2760
ln4.y1 = 3840
ln4.y2 = 120

ln5.x1 = 2760
ln5.X2 = 4080
ln5.y1 = 7560
ln5.y2 = 7560

ln6.x1 = 4080
ln6.X2 = 4080
ln6.y1 = 3840
ln6.y2 = 7560

ln7.x1 = 2760
ln7.X2 = 2760
ln7.y1 = 3840
ln7.y2 = 7560

'VOLTA AO ORIGINAL DO ESTABILIZADOR
ln11.x1 = 6360
ln11.X2 = 7080
ln11.y1 = 3840
ln11.y2 = 3840

ln22.x1 = 6360
ln22.X2 = 7080
ln22.y1 = 2640
ln22.y2 = 2640

ln33.x1 = 7080
ln33.X2 = 7080
ln33.y1 = 3840
ln33.y2 = 2640

ln44.x1 = 6360
ln44.X2 = 6360
ln44.y1 = 3840
ln44.y2 = 2640

ln55.x1 = 7080
ln55.X2 = 6360
ln55.y1 = 4920
ln55.y2 = 4920

ln66.x1 = 7080
ln66.X2 = 7080
ln66.y1 = 3840
ln66.y2 = 4920

ln77.x1 = 6360
ln77.X2 = 6360
ln77.y1 = 3840
ln77.y2 = 4920

lineac.x1 = 3120
linecorda.x1 = 2760
linecorda.X2 = 4080

'Cgs e PN
cgtotal15.Visible = False
cgtotal15.Left = 240
cgtotal15.Top = 3720

cgvazio15.Visible = False
cgvazio15.Left = 240
cgvazio15.Top = 3720

pn15.Visible = False
pn15.Left = 240
pn15.Top = 3720

'linhas cg pn e cgv
lcg.X2 = 360
lcg.Visible = False

lcgv.X2 = 360
lcgv.Visible = False

lpn.X2 = 360
lpn.Visible = False

cgtext.Visible = False
cgtext.Left = 480
cgvtext.Visible = False
cgvtext.Left = 480
pntext.Visible = False
pntext.Left = 480

actext.Visible = False
lnac.Visible = False
actext.Left = 480
lnac.X2 = 360


End Sub
