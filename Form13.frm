VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form Form13 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros - estruturas"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10035
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10035
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   5100
         TabIndex        =   9
         Top             =   240
         Width           =   4815
         Begin VB.Frame Frame10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Observações"
            Height          =   2895
            Left            =   120
            TabIndex        =   14
            Top             =   3120
            Width           =   4575
            Begin VB.TextBox altlong 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00FF8080&
               Height          =   285
               Left            =   3480
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   2400
               Width           =   615
            End
            Begin VB.TextBox porclong 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1080
               TabIndex        =   18
               Text            =   "80"
               Top             =   2400
               Width           =   495
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   $"Form13.frx":030A
               ForeColor       =   &H00808080&
               Height          =   555
               Left            =   120
               TabIndex        =   43
               Top             =   1320
               Width           =   4275
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "2- Sómente a longarina da asa será calculada."
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   120
               TabIndex        =   42
               Top             =   840
               Width           =   3285
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1- A estrutura da asa será considerada cantilever."
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   120
               TabIndex        =   41
               Top             =   360
               Width           =   3510
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Altura da longarina :               cm"
               Height          =   195
               Left            =   2040
               TabIndex        =   19
               Top             =   2400
               Width           =   2295
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Novo valor :              %"
               Height          =   195
               Left            =   120
               TabIndex        =   17
               Top             =   2400
               Width           =   1620
            End
         End
         Begin VB.PictureBox Picture1 
            Height          =   435
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   4515
            TabIndex        =   11
            Top             =   6120
            Width           =   4575
            Begin VB.CommandButton Command2 
               BackColor       =   &H00C0C0C0&
               Caption         =   "&Calcular e Sair"
               Height          =   375
               Left            =   0
               MouseIcon       =   "Form13.frx":03AB
               MousePointer    =   99  'Custom
               Style           =   1  'Graphical
               TabIndex        =   0
               Top             =   0
               Width           =   4515
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Acelerações (Nº de Gs)"
            ForeColor       =   &H00000000&
            Height          =   2895
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   4575
            Begin VB.OptionButton Checkcgorig 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Nº de Gs calculado"
               Height          =   255
               Left            =   2520
               TabIndex        =   49
               Top             =   1680
               Width           =   1695
            End
            Begin VB.OptionButton checkg 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Modificar Nº de Gs"
               Height          =   255
               Left            =   240
               TabIndex        =   48
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox gsnegatual 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00FF8080&
               Height          =   285
               Left            =   3720
               TabIndex        =   16
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox gsatual 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00FF8080&
               Height          =   285
               Left            =   3720
               TabIndex        =   12
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Caso 1: manobra brusca em Velocidade maxima e Cl máximo Caso 2: manobra brusca em velocidade máx. e 80% Cl máx."
               ForeColor       =   &H00808080&
               Height          =   435
               Left            =   120
               TabIndex        =   24
               Top             =   2280
               Width           =   4275
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Caso 2- Nº de Gs (negativos) atuando no avião :"
               Height          =   195
               Left            =   240
               TabIndex        =   15
               Top             =   960
               Width           =   3435
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Caso 1- Nº de Gs (positivos) atuando no avião :"
               Height          =   195
               Left            =   240
               TabIndex        =   13
               Top             =   480
               Width           =   3360
            End
         End
         Begin VB.Label espcomp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "espcomp"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   840
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label esptrac 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "esptrac"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   600
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label alma 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "alma"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Visible         =   0   'False
            Width           =   330
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   6615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4935
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0E0E0&
            Height          =   1815
            Left            =   120
            TabIndex        =   28
            Top             =   4680
            Width           =   4695
            Begin VB.TextBox elasto 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00FF8080&
               Height          =   285
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox cizad 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00FF8080&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   31
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox sigtad 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00FF8080&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   30
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox sigcad 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00FF8080&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E = módulo de elasticidade :"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   120
               TabIndex        =   40
               Top             =   1320
               Width           =   1995
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tcza = tensão de cizalhamento  pp fib. :"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   120
               TabIndex        =   39
               Top             =   960
               Width           =   2835
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "kg/cm2"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   3720
               TabIndex        =   38
               Top             =   1320
               Width           =   555
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "kg/cm2"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   3720
               TabIndex        =   37
               Top             =   960
               Width           =   555
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "kg/cm2"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   3720
               TabIndex        =   36
               Top             =   600
               Width           =   555
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "kg/cm2"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   3720
               TabIndex        =   35
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tta = tensão admissivel a tração :"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   120
               TabIndex        =   34
               Top             =   600
               Width           =   2385
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tca = tensão admissivel a compressão :"
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   2835
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Seleciona o material para longarina"
            Height          =   2175
            Left            =   120
            TabIndex        =   4
            Top             =   2520
            Width           =   4695
            Begin VB.PictureBox Picture3 
               Height          =   315
               Left            =   120
               ScaleHeight     =   255
               ScaleWidth      =   2595
               TabIndex        =   45
               Top             =   1680
               Width           =   2655
               Begin VB.CommandButton Command4 
                  Caption         =   "&Alma"
                  Height          =   255
                  Left            =   0
                  MouseIcon       =   "Form13.frx":06B5
                  MousePointer    =   99  'Custom
                  TabIndex        =   47
                  Top             =   0
                  Width           =   2595
               End
            End
            Begin VB.PictureBox Picture2 
               Height          =   315
               Left            =   120
               ScaleHeight     =   255
               ScaleWidth      =   2595
               TabIndex        =   44
               Top             =   1200
               Width           =   2655
               Begin VB.CommandButton Command3 
                  Caption         =   "&Mesa"
                  Height          =   255
                  Left            =   0
                  MouseIcon       =   "Form13.frx":09BF
                  MousePointer    =   99  'Custom
                  TabIndex        =   46
                  Top             =   0
                  Width           =   2595
               End
            End
            Begin VB.TextBox matalma 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox matmesa 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Data Data1 
               Caption         =   "Data1"
               Connect         =   "Access"
               DatabaseName    =   ""
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   300
               Left            =   120
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   "Materiais"
               Top             =   2925
               Visible         =   0   'False
               Width           =   1065
            End
            Begin MSDBGrid.DBGrid DBGrid1 
               Bindings        =   "Form13.frx":0CC9
               Height          =   735
               Left            =   120
               OleObjectBlob   =   "Form13.frx":0CD9
               TabIndex        =   25
               Top             =   360
               Width           =   4455
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Selecione a seção transversal da longarina"
            Height          =   2285
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   4695
            Begin VB.OptionButton vigaC 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Viga perfil C"
               Height          =   255
               Left            =   2640
               TabIndex        =   8
               Top             =   1560
               Width           =   1215
            End
            Begin VB.OptionButton vigacaixa 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Viga caixa"
               Height          =   255
               Left            =   240
               TabIndex        =   7
               Top             =   1560
               Width           =   1095
            End
            Begin VB.OptionButton vigaI 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Viga perfil I"
               Height          =   255
               Left            =   2640
               TabIndex        =   6
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton vigaret 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Retangular"
               Height          =   255
               Left            =   240
               TabIndex        =   5
               Top             =   600
               Width           =   1095
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               X1              =   1800
               X2              =   1800
               Y1              =   1925
               Y2              =   1455
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               BorderWidth     =   2
               X1              =   1560
               X2              =   1560
               Y1              =   1925
               Y2              =   1455
            End
            Begin VB.Shape Shape14 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               Height          =   75
               Left            =   4080
               Top             =   1440
               Width           =   135
            End
            Begin VB.Shape Shape13 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               Height          =   75
               Left            =   4080
               Top             =   1860
               Width           =   135
            End
            Begin VB.Shape Shape5 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               Height          =   495
               Left            =   4200
               Top             =   1440
               Width           =   75
            End
            Begin VB.Shape Shape12 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               Height          =   495
               Left            =   4080
               Top             =   480
               Width           =   75
            End
            Begin VB.Shape Shape11 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               Height          =   75
               Left            =   3960
               Top             =   900
               Width           =   315
            End
            Begin VB.Shape Shape10 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               Height          =   75
               Left            =   3960
               Top             =   480
               Width           =   315
            End
            Begin VB.Shape Shape9 
               BorderColor     =   &H00404040&
               Height          =   735
               Left            =   2400
               Top             =   1320
               Width           =   2175
            End
            Begin VB.Shape Shape8 
               BorderColor     =   &H00404040&
               Height          =   735
               Left            =   2400
               Top             =   360
               Width           =   2175
            End
            Begin VB.Shape Shape7 
               BorderColor     =   &H00404040&
               Height          =   735
               Left            =   120
               Top             =   1320
               Width           =   2055
            End
            Begin VB.Shape Shape6 
               BorderColor     =   &H00404040&
               Height          =   735
               Left            =   120
               Top             =   360
               Width           =   2055
            End
            Begin VB.Shape Shape4 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               Height          =   135
               Left            =   1560
               Top             =   1800
               Width           =   255
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               Height          =   135
               Left            =   1560
               Top             =   1440
               Width           =   255
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               Height          =   495
               Left            =   1560
               Top             =   480
               Width           =   255
            End
         End
      End
   End
   Begin VB.Menu modifica 
      Caption         =   "&Modificar "
      Begin VB.Menu escmaterial 
         Caption         =   "Escolha do material "
      End
   End
   Begin VB.Menu filemat 
      Caption         =   "&Cadastrar"
      Begin VB.Menu cadmat 
         Caption         =   "&Materiais"
      End
      Begin VB.Menu excmat 
         Caption         =   "&Excluir materiais"
      End
   End
   Begin VB.Menu filefer13 
      Caption         =   "&Ferramentas"
      Begin VB.Menu filecalcon 
         Caption         =   "Ca&lculadora && Conversor"
      End
   End
   Begin VB.Menu filesai13 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cadmat_Click()
Form13.Visible = False
Form28.Frame1.Visible = False
Form28.Frame2.Visible = False
Form28.Frame3.Visible = True
Form28.Visible = True
Form28.n13.SetFocus
End Sub

Private Sub Checkcgorig_Click()
'Calculo do Nº de Gs no form13
dar = Form1.roar * 0.00194

Form13.gsatual = Form7.cl3dasa * ((dar / 2) * (Form2.Sw * 10.76) * ((Form8.vmax * 0.9113) ^ 2)) / (Form1.wo * 2.2)
Form13.gsatual = Format(Form13.gsatual, "0.00")

Form13.gsnegatual = 0.8 * Form7.cl3dasa * ((dar / 2) * (Form2.Sw * 10.76) * ((Form8.vmax * 0.9113) ^ 2)) / (Form1.wo * 2.2)
Form13.gsnegatual = Format(Form13.gsnegatual, "0.00")

Form13.gsnegatual.BackColor = &HE0E0E0
Form13.gsnegatual.Locked = True

Form13.gsatual.BackColor = &HE0E0E0
Form13.gsatual.Locked = True



End Sub

Private Sub checkg_Click()

If Form21.Data1.Recordset("Projeto") = Form20.p1 Then
    Form13.porclong = Form21.Data1.Recordset("Hlong")
    Form13.gsatual = Form21.Data1.Recordset("CGatual")
    Form13.gsnegatual = Form21.Data1.Recordset("CGnega")
End If

If Form21.Data1.Recordset("Projeto") <> Form20.p1 Then
    MsgBox "Entre com os valores desejados para NºGs"
    Form13.gsatual.SetFocus
End If

Form13.gsatual.BackColor = &H80000005
Form13.gsnegatual.BackColor = &H80000005

Form13.gsatual.Locked = False
Form13.gsnegatual.Locked = False

End Sub

Private Sub cizad_LostFocus()
cizad.Locked = True
cizad.BackColor = &HE0E0E0
End Sub

Private Sub Command2_Click()

If Form13.gsatual = Empty Or Form13.gsnegatual = Empty Then
MsgBox "Insirindo(ir) o Nº de Gs (positivo e negativo) ", vbCritical
Call checkg_Click
Exit Sub
End If

If vigaI.Value = True And matalma = Empty Then
MsgBox "Escolha o material para a alma da viga", vbInformation
Exit Sub
End If

If vigacaixa.Value = True And matalma = Empty Then
MsgBox "Escolha o material para a alma da viga", vbInformation
Exit Sub
End If

If vigaC.Value = True And matalma = Empty Then
MsgBox "Escolha o material para a alma da viga", vbInformation
Exit Sub
End If

If matalma = Empty And matmesa = Empty Then
MsgBox "Escolha o material da longarina", vbInformation
Exit Sub
End If

If gsatual = 0 Then
gsatual = 1
MsgBox "Nº de Gs deve ser diferente de zero ", vbCritical
End If

If gsnegatual = 0 Then
gsnegatual = 1
MsgBox "Nº de Gs deve ser diferente de zero ", vbCritical
End If

Form11.sigcizal = Form13.cizad
Form11.sigtrac = Form13.sigtad
Form11.sigcomp = Form13.sigcad

'Condição para Sigma (tração ou compressão)
If Val(Form13.sigtad) > Val(Form13.sigcad) Then sigproj = Val(Form13.sigcad)
If Val(Form13.sigtad) > Val(Form13.sigcad) Then sigproj = Val(Form13.sigtad)
If Val(Form13.sigtad) = Val(Form13.sigcad) Then sigproj = Val(Form13.sigcad)

'CALCULO DE VIGA I
If Form13.vigaI.Value = True Then

Form11.section = "Viga I"
Form11.npecas = "1 Pç"

h = ((Form2.espessura / 100) * (porclong / 100) * Form2.crw) * 100
h = Format(h, "0.00")

Form11.altmax = h
Form11.altmax = Format(Form11.altmax, "0.00")

a = (0.2 * h) + 0.1
a = Format(a, "0.00")
Form11.largamax = a
Form11.largamax = Format(Form11.largamax, "0.00")

'utiliza Gnegativo (gsneatual)- mesa inferior
Form13.esptrac = (2 * Form11.m2 * Form13.gsnegatual * 100) / (a * h * sigproj)
Form13.esptrac = Format(Form13.esptrac, "0.00")

'utiliza Gpositivo (gsatual)- mesa superior
Form13.espcomp = (2 * Form11.m2 * Form13.gsatual * 100) / (a * h * sigproj)
Form13.espcomp = Format(Form13.espcomp, "0.00")

Form11.mesatc = Form11.largamax + " x " + Form13.esptrac
Form11.mesacp = Form11.largamax + " x " + Form13.espcomp

alma = 2 * Form11.v2 * Form13.gsatual / (Form13.cizad * h)
alma = Format(alma, "0.00")
Form11.alma = alma + " x " + h

If Abs(Form13.esptrac) + (Form13.espcomp) > h Then
MsgBox "Altura das mesas > Altura total da viga. Escolha outro material ou calcule como Viga retangular", vbCritical
Form13.Visible = True
End If

If Abs(Form13.esptrac) + (Form13.espcomp) <= h Then
Form11.materialA = Form13.matalma
Form11.materialM = Form13.matmesa
Form13.Visible = False
End If

End If
'FIM VIGA I

'VIGA RETANGULAR MACIÇA
If Form13.vigaret.Value = True Then

Form11.npecas = "1 Pç"

'Condição para G
If Val(Form13.gsatual) > Val(Form13.gsnegatual) Then gproj = Val(Form13.gsatual)
If Val(Form13.gsatual) < Val(Form13.gsnegatual) Then gproj = Val(Form13.gsnegatual)
If Val(Form13.gsatual) = Val(Form13.gsnegatual) Then gproj = Val(Form13.gsatual)

Form11.section = "Retangular"

h = ((Form2.espessura / 100) * (porclong / 100) * Form2.crw) * 100
h = Format(h, "0.00")

Form11.altmax = h
Form11.altmax = Format(Form11.altmax, "0.00")

Form11.largamax = (6 * Form11.m2 * gproj * 100) / (sigproj * (h ^ 2))
Form11.largamax = Format(Form11.largamax, "0.00")

Form11.mesacp = "-----"

Form11.mesatc = "-----"

Form11.alma = "-----"

Form13.Visible = False

Form11.materialA = Form13.matalma
Form11.materialM = Form13.matmesa

End If
'FIM VIGA RETANGULAR

'VIGA CAIXA

If Form13.vigacaixa.Value = True Then

Form11.npecas = "2 Pçs"

If Val(Form13.gsatual) > Val(Form13.gsnegatual) Then gproj = Val(Form13.gsatual)
If Val(Form13.gsatual) < Val(Form13.gsnegatual) Then gproj = Val(Form13.gsnegatual)
If Val(Form13.gsatual) = Val(Form13.gsnegatual) Then gproj = Val(Form13.gsatual)

Form11.section = "Viga caixa"

h = ((Form2.espessura / 100) * (porclong / 100) * Form2.crw) * 100
h = Format(h, "0.00")

Form11.altmax = h
Form11.altmax = Format(Form11.altmax, "0.00")

a = (0.2 * h) + 0.1
a = Format(a, "0.00")

'v = (((6 * Form11.m2 * gproj * 100) * h) / ((sigproj / 0.55) * a)) ^ (1 / 3)
 
'Form13.esptrac = (Abs(h) - Abs(v)) / 2
'Form13.esptrac = Format(Form13.esptrac, "0.00")

'Form13.espcomp = (Abs(h) - Abs(v)) / 2
'Form13.espcomp = Format(Form13.espcomp, "0.00")

'utiliza Gnegativo (gsneatual)- mesa inferior
Form13.esptrac = (2 * Form11.m2 * Form13.gsnegatual * 100) / (a * h * sigproj)
Form13.esptrac = Format(Form13.esptrac, "0.00")

'utiliza Gpositivo (gsatual)- mesa superior
Form13.espcomp = (2 * Form11.m2 * Form13.gsatual * 100) / (a * h * sigproj)
Form13.espcomp = Format(Form13.espcomp, "0.00")

Form11.largamax = a

Form11.largamax = a
Form11.largamax = Format(Form11.largamax, "0.00")

Form11.mesatc = Form11.largamax + " x " + Form13.esptrac
Form11.mesacp = Form11.largamax + " x " + Form13.espcomp

alma = (2 * Form11.v2 * gproj / (Form13.cizad * h)) / 2
alma = Format(alma, "0.00")
Form11.alma = alma + " x " + h

If Abs(Form13.esptrac) + (Form13.espcomp) > h Then
MsgBox "Altura das mesas > Altura total da viga. Escolha outro material ou calcule como Viga retangular", vbCritical
Form13.Visible = True
End If

If Abs(Form13.esptrac) + (Form13.espcomp) <= h Then
Form11.materialA = Form13.matalma
Form11.materialM = Form13.matmesa
Form13.Visible = False
End If

End If

'FIM VIGA CAIXA

'VIGA C

If Form13.vigaC.Value = True Then

Form11.npecas = "1 Pç"

Form11.section = "Viga C"

h = ((Form2.espessura / 100) * (porclong / 100) * Form2.crw) * 100
h = Format(h, "0.00")

Form11.altmax = h
Form11.altmax = Format(Form11.altmax, "0.00")

a = (0.2 * h) + 0.1
a = Format(a, "0.00")
Form11.largamax = a
Form11.largamax = Format(Form11.largamax, "0.00")

'utiliza Gnegativo (gsneatual)- mesa inferior
Form13.esptrac = (2 * Form11.m2 * Form13.gsnegatual * 100) / (a * h * sigproj)
Form13.esptrac = Format(Form13.esptrac, "0.00")

'utiliza Gpositivo (gsatual)- mesa superior
Form13.espcomp = (2 * Form11.m2 * Form13.gsatual * 100) / (a * h * sigproj)
Form13.espcomp = Format(Form13.espcomp, "0.00")

Form11.mesatc = Form11.largamax + " x " + Form13.esptrac
Form11.mesacp = Form11.largamax + " x " + Form13.espcomp

alma = 2 * Form11.v2 * Form13.gsatual / (Form13.cizad * h)
alma = Format(alma, "0.00")
Form11.alma = alma + " x " + h

If Abs(Form13.esptrac) + (Form13.espcomp) > h Then
MsgBox "Altura das mesas > Altura total da viga. Escolha outro material ou calcule como Viga retangular", vbCritical
Form13.Visible = True
End If

If Abs(Form13.esptrac) + (Form13.espcomp) <= h Then
Form11.materialA = Form13.matalma
Form11.materialM = Form13.matmesa
Form13.Visible = False
End If


End If

End Sub

Private Sub Command3_Click()
If vigaret.Value = True Then
Call Command4_Click
Command4.Enabled = False
End If
sigcad = Form13.Data1.Recordset("tca")
sigtad = Form13.Data1.Recordset("tta")
matmesa = Form13.Data1.Recordset("Nome")
elasto = Form13.Data1.Recordset("e")

Command3.Enabled = False
End Sub

Private Sub Command4_Click()
cizad = Form13.Data1.Recordset("tcza")
matalma = Form13.Data1.Recordset("Nome")

Command4.Enabled = False
End Sub
Private Sub elasto_LostFocus()
sigcad.Locked = True
sigtad.Locked = True
elasto.Locked = True
sigcad.BackColor = &HE0E0E0
sigtad.BackColor = &HE0E0E0
elasto.BackColor = &HE0E0E0
End Sub

Private Sub escmaterial_Click()

Form13.Command3.Enabled = True
Form13.Command4.Enabled = True

variavelest = "Select * from materiais where nome=nome"
Form13.Data1.RecordSource = variavelest
Form13.Data1.Refresh
End Sub

Private Sub excmat_Click()
MsgBox "Verifique se existe algum projeto utilizando este material e então modifique-o.", vbInformation

If Command3.Enabled = False Or Command4.Enabled = False Then MsgBox " Clique Modificar para abrir a listagem de materiais", vbInformation

    If Command3.Enabled = True Or Command4.Enabled = True Then
    
    resp2 = MsgBox("Voce tem certeza de que deseja excluir o material selecionado?", vbOKCancel)

        If resp2 = vbOK Then
            Form13.Data1.Recordset.Delete
            Form13.Data1.Refresh
        End If
 
            If resp2 = vbCancel Then
                Form13.Visible = False
            End If
    End If
    
End Sub

Private Sub filecalcon_Click()
Form29.Show
End Sub

Private Sub filesai13_Click()
Form13.Visible = False
End Sub

Private Sub Form_Activate()
If vigaret.Value = True Then
Call Command4_Click
Command4.Enabled = False
End If
sigcad = Form13.Data1.Recordset("tca")
sigtad = Form13.Data1.Recordset("tta")
matmesa = Form13.Data1.Recordset("Nome")
elasto = Form13.Data1.Recordset("e")

Command3.Enabled = False

Call Command2_Click

End Sub

Private Sub Form_Load()
    Data1.DatabaseName = SR
    vigaret.Value = True

End Sub
Private Sub gsatual_GotFocus()
gsatual.SelStart = 0
gsatual.SelLength = Len(gsatual)
End Sub

Private Sub gsatual_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then gsnegatual.SetFocus
End Sub

Private Sub gsnegatual_LostFocus()
Form13.gsnegatual.BackColor = &HE0E0E0
Form13.gsnegatual.Locked = True
Form13.gsatual.BackColor = &HE0E0E0
Form13.gsatual.Locked = True
End Sub
Private Sub gsnegatual_GotFocus()
gsnegatual.SelStart = 0
gsnegatual.SelLength = Len(gsnegatual)
End Sub

Private Sub gsnegatual_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command2.SetFocus
End Sub


Private Sub Option20_Click()
Form11.section = "Retangular"
End Sub

Private Sub Option21_Click()
Form11.section = "Perfil I"
End Sub

Private Sub Option22_Click()
Form11.section = "Viga caixa"
End Sub

Private Sub Option23_Click()
Form11.section = "Tubular"
End Sub

Private Sub porclong_LostFocus()
'Calculo da altura da longarina para o form13
If Not IsNumeric(Form13.porclong) Or Form13.porclong = "0" Then
    MsgBox "Insira a relação altura da longarina/corda da asa em ""%""", vbOKOnly
    Form13.porclong.SetFocus
    Form13.porclong = 0
    Exit Sub
End If

Form13.altlong = (Form13.porclong / 100) * ((Form2.espessura / 100) * Form2.crw * 100)
Form13.altlong = Format(Form13.altlong, "0.00")
End Sub
Private Sub porclong_GotFocus()
porclong.SelStart = 0
porclong.SelLength = Len(gsatual)
End Sub

Private Sub porclong_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command2.SetFocus
End Sub

Private Sub vigaC_Click()
If vigaret.Value = False Then Command4.Enabled = True
End Sub

Private Sub vigacaixa_Click()
If vigaret.Value = False Then Command4.Enabled = True
End Sub

Private Sub vigaI_Click()
If vigaret.Value = False Then Command4.Enabled = True
End Sub

Private Sub vigaret_Click()
If matmesa <> Empty Then
matalma = matmesa
Call Command3_Click
End If
End Sub
