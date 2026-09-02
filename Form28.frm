VERSION 5.00
Begin VB.Form Form28 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro: Motores - Perfis - Materiais"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7275
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form28.frx":0000
   LinkTopic       =   "Form28"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&MATERIAIS PARA LONGARINA"
      ForeColor       =   &H00008000&
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6975
      Begin VB.PictureBox Picture3 
         Height          =   375
         Left            =   360
         ScaleHeight     =   315
         ScaleWidth      =   6195
         TabIndex        =   45
         Top             =   2040
         Width           =   6255
         Begin VB.CommandButton Command4 
            Caption         =   "&Cadastrar novo material"
            Enabled         =   0   'False
            Height          =   315
            Left            =   0
            MouseIcon       =   "Form28.frx":030A
            MousePointer    =   99  'Custom
            TabIndex        =   46
            Top             =   0
            Width           =   6195
         End
      End
      Begin VB.TextBox e13 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5520
         TabIndex        =   44
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox tcz13 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5400
         TabIndex        =   42
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox tt13 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   40
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox tc13 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   38
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox n13 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modulo de elasticidade E"
         Height          =   195
         Left            =   3600
         TabIndex        =   43
         Top             =   960
         Width           =   1785
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tensão de cizalhamento                kgf/cm2"
         Height          =   195
         Left            =   3600
         TabIndex        =   41
         Top             =   1560
         Width           =   3060
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tensão de tração                           kgf/cm2"
         Height          =   195
         Left            =   360
         TabIndex        =   39
         Top             =   1560
         Width           =   3075
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tensão de compressão                  kgf/cm2"
         Height          =   195
         Left            =   360
         TabIndex        =   37
         Top             =   960
         Width           =   3075
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do material"
         Height          =   195
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&PERFIS AERODINÂMICOS"
      ForeColor       =   &H000000FF&
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6975
      Begin VB.TextBox cmmax 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4200
         TabIndex        =   48
         Top             =   1080
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   360
         ScaleHeight     =   315
         ScaleWidth      =   6195
         TabIndex        =   32
         Top             =   2040
         Width           =   6255
         Begin VB.CommandButton Command3 
            Caption         =   "Estabilizador vertical"
            Height          =   315
            Left            =   4080
            MouseIcon       =   "Form28.frx":0614
            MousePointer    =   99  'Custom
            TabIndex        =   35
            Top             =   0
            Width           =   2115
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Asa"
            Height          =   315
            Left            =   0
            MouseIcon       =   "Form28.frx":091E
            MousePointer    =   99  'Custom
            TabIndex        =   34
            Top             =   0
            Width           =   2055
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Estabilizador horizontal"
            Height          =   315
            Left            =   2040
            MouseIcon       =   "Form28.frx":0C28
            MousePointer    =   99  'Custom
            TabIndex        =   33
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.TextBox esp28 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5880
         TabIndex        =   31
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox util28 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3840
         TabIndex        =   29
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox rey28 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox alfam28 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5880
         TabIndex        =   25
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox alfa028 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5880
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox cd28 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4200
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox cl28 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4200
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Nperf 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cm 2 graus"
         Height          =   195
         Left            =   3360
         TabIndex        =   47
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espessura"
         Height          =   195
         Left            =   5040
         TabIndex        =   30
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Utilização"
         Height          =   195
         Left            =   3000
         TabIndex        =   28
         Top             =   1560
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Reynolds"
         Height          =   195
         Left            =   360
         TabIndex        =   26
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alfa Clmáx."
         Height          =   195
         Left            =   5040
         TabIndex        =   24
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alfa Clzero"
         Height          =   195
         Left            =   5040
         TabIndex        =   22
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cd 2 graus"
         Height          =   195
         Left            =   3360
         TabIndex        =   20
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cl máx."
         Height          =   195
         Left            =   3360
         TabIndex        =   18
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do perfil"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&MOTORES"
      ForeColor       =   &H00FF0000&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   360
         ScaleHeight     =   315
         ScaleWidth      =   6195
         TabIndex        =   15
         Top             =   2040
         Width           =   6255
         Begin VB.CommandButton Command10 
            Caption         =   "&Cadastrar novo motor"
            Enabled         =   0   'False
            Height          =   315
            Left            =   0
            MouseIcon       =   "Form28.frx":0F32
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   0
            Width           =   6195
         End
      End
      Begin VB.TextBox t1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox pc1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5880
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox php1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox pm1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4080
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox prpm1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5880
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do motor "
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Polegadas cúbicas (ci)"
         Height          =   195
         Left            =   4200
         TabIndex        =   13
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Potencia (Hp)"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peso do motor (kg)"
         Height          =   195
         Left            =   2640
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rpm"
         Height          =   195
         Left            =   5400
         TabIndex        =   10
         Top             =   1320
         Width           =   330
      End
   End
   Begin VB.Menu filesai28 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alfa028_GotFocus()
If Not IsNumeric(cd28) Or cd28 = "0" Then
    MsgBox "Insira o valor de Cd conforme recomendado", vbOKOnly
    cd28.SetFocus
    cd28 = 0
    Exit Sub
End If

alfa028.SelStart = 0
alfa028.SelLength = Len(alfa028)
End Sub
Private Sub alfa028_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.alfam28.SetFocus
End Sub

Private Sub alfam28_GotFocus()
alfam28.SelStart = 0
alfam28.SelLength = Len(cd28)
End Sub
Private Sub alfam28_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.cmmax.SetFocus
End Sub
Private Sub cd28_GotFocus()
If Not IsNumeric(cl28) Or cl28 = "0" Then
    MsgBox "Insira o valor de Cl conforme recomendado", vbOKOnly
    cl28.SetFocus
    cl28 = 0
    Exit Sub
End If

cd28.SelStart = 0
cd28.SelLength = Len(cd28)
End Sub
Private Sub cd28_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.alfa028.SetFocus
End Sub
Private Sub cl28_GotFocus()
cl28.SelStart = 0
cl28.SelLength = Len(cl28)
End Sub
Private Sub cl28_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.cd28.SetFocus
End Sub


Private Sub cmmax_GotFocus()
If Not IsNumeric(alfam28) Or alfam28 = "0" Then
    MsgBox "Insira o valor de Alfa conforme recomendado", vbOKOnly
    alfam28.SetFocus
    alfam28 = 0
    Exit Sub
End If

cmmax.SelStart = 0
cmmax.SelLength = Len(cmmax)
End Sub
Private Sub cmmax_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.esp28.SetFocus
End Sub


Private Sub Command1_Click()

variavelw = "Select * from Aerofolios where Perfil=Perfil"
Form4.Data1.RecordSource = variavelw
Form4.Data1.Refresh


If Form28.Nperf = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.cl28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.cd28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.cmmax = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.alfa028 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.alfam28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.esp28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.util28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.rey28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If

If Val(cd28) > 0.05 Then
MsgBox "O Cd está muito alto, diminua-o."
Exit Sub
End If

Form4.Data1.Recordset.AddNew

Form4.Data1.Recordset("Perfil") = Form28.Nperf
Form4.Data1.Recordset("clmax") = Form28.cl28
Form4.Data1.Recordset("cdmax") = Form28.cd28
Form4.Data1.Recordset("cmo") = Form28.cmmax
Form4.Data1.Recordset("angulo") = Form28.alfam28
Form4.Data1.Recordset("angulo0") = Form28.alfa028
Form4.Data1.Recordset("reynolds") = Form28.rey28
Form4.Data1.Recordset("espessura") = Form28.esp28
Form4.Data1.Recordset("utilizado para") = Form28.util28

Form4.Data1.Recordset.Update
    Form4.Data1.Refresh

Form4.Data1.Recordset.MoveLast

Form28.Visible = False
Form4.Visible = True
End Sub

Private Sub Command10_Click()
If Form28.prpm1 > 30000 Then
MsgBox "O Nºde rotações deve ser menor ou igual a 30000Rpm", vbCritical
Exit Sub
End If
        

If Not IsNumeric(prpm1) Or prpm1 = "0" Then
    MsgBox "Insira o valor de Rpm conforme recomendado", vbOKOnly
    prpm1.SetFocus
    prpm1 = 0
    Exit Sub
End If



If php1 <= 15 Then
    If Form28.t1 = "" Then
        MsgBox "Preencha todos os dados", vbInformation
        Exit Sub
    End If
    If Form28.pc1 = "" Then
        MsgBox "Preencha todos os dados", vbInformation
        Exit Sub
    End If
    If Form28.php1 = "" Then
        MsgBox "Preencha todos os dados", vbInformation
        Exit Sub
    End If
    If Form28.pm1 = "" Then
        MsgBox "Preencha todos os dados", vbInformation
        Exit Sub
    End If
    If Form28.prpm1 = "" Then
        MsgBox "Preencha todos os dados", vbInformation
        Exit Sub
    End If

    

End If

Form28.Visible = False

End Sub

Private Sub Command2_Click()

variavelht = "Select * from PerfilHT where Perfil=Perfil"
Form4.Data2.RecordSource = variavelht
Form4.Data2.Refresh

If Form28.Nperf = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.cl28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.cd28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.cmmax = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.alfa028 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.alfam28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.esp28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.util28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.rey28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If

If cd28 > 0.05 Then
MsgBox "O Cd está muito alto, diminua-o."
Exit Sub
End If

Form4.Data2.Recordset.AddNew

Form4.Data2.Recordset("Perfil") = Form28.Nperf
Form4.Data2.Recordset("clmax") = Form28.cl28
Form4.Data2.Recordset("cdmax") = Form28.cd28
Form4.Data2.Recordset("cmo") = Form28.cmmax
Form4.Data2.Recordset("angulo") = Form28.alfam28
Form4.Data2.Recordset("angulo0") = Form28.alfa028
Form4.Data2.Recordset("reynolds") = Form28.rey28
Form4.Data2.Recordset("espessura") = Form28.esp28
Form4.Data2.Recordset("utilizado para") = Form28.util28

Form4.Data2.Recordset.Update
    Form4.Data2.Refresh

Form4.Data1.Recordset.MoveLast

Form28.Visible = False
Form4.Visible = True
End Sub

Private Sub Command3_Click()

variavelvt = "Select * from PerfilVT where Perfil=Perfil"
Form4.Data3.RecordSource = variavelvt
Form4.Data3.Refresh

If Form28.Nperf = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.cl28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.cd28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.cmmax = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.alfa028 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.alfam28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.esp28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.util28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.rey28 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If

If cd28 > 0.05 Then
MsgBox "O Cd está muito alto, diminua-o."
Exit Sub
End If


Form4.Data3.Recordset.AddNew

Form4.Data3.Recordset("Perfil") = Form28.Nperf
Form4.Data3.Recordset("clmax") = Form28.cl28
Form4.Data3.Recordset("cdmax") = Form28.cd28
Form4.Data3.Recordset("cmo") = Form28.cmmax
Form4.Data3.Recordset("angulo") = Form28.alfam28
Form4.Data3.Recordset("angulo0") = Form28.alfa028
Form4.Data3.Recordset("reynolds") = Form28.rey28
Form4.Data3.Recordset("espessura") = Form28.esp28
Form4.Data3.Recordset("utilizado para") = Form28.util28

Form4.Data3.Recordset.Update
    Form4.Data3.Refresh

Form4.Data3.Recordset.MoveLast

Form28.Visible = False
Form4.Visible = True
End Sub

Private Sub Command4_Click()

variavelest = "Select * from materiais where nome=nome"
Form13.Data1.RecordSource = variavelest
Form13.Data1.Refresh

If Not IsNumeric(tcz13) Or tcz13 = "0" Then
    MsgBox "Insira o valor da tensão conforme recomendado", vbOKOnly
    tcz13.SetFocus
    tcz13 = 0
    Exit Sub
End If

If Form28.n13 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.tc13 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.tt13 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.tcz13 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If
If Form28.e13 = "" Then
    MsgBox "Preencha todos os dados", vbInformation
    Exit Sub
End If

Form13.Data1.Recordset.AddNew

Form13.Data1.Recordset("Nome") = Form28.n13
Form13.Data1.Recordset("tca") = Form28.tc13
Form13.Data1.Recordset("tta") = Form28.tt13
Form13.Data1.Recordset("tcza") = Form28.tcz13
Form13.Data1.Recordset("E") = Form28.e13

Form13.Data1.Recordset.Update
    Form13.Data1.Refresh

Form13.Data1.Recordset.MoveLast

Form13.Visible = True
Form28.Visible = False

End Sub


Private Sub e13_GotFocus()
If Not IsNumeric(tc13) Or tc13 = "0" Then
    MsgBox "Insira o valor da tensão conforme recomendado", vbOKOnly
    tc13.SetFocus
    tc13 = 0
    Exit Sub
End If

e13.SelStart = 0
e13.SelLength = Len(e13)
End Sub
Private Sub e13_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.tt13.SetFocus
End Sub

Private Sub esp28_GotFocus()
If Not IsNumeric(cmmax) Or cmmax = "0" Then
    MsgBox "Insira o valor de Cm conforme recomendado", vbOKOnly
    cmmax.SetFocus
    cmmax = 0
    Exit Sub
End If

esp28.SelStart = 0
esp28.SelLength = Len(esp28)
End Sub
Private Sub esp28_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.rey28.SetFocus
End Sub
Private Sub filesai28_Click()
Form28.Visible = False
End Sub



Private Sub n13_GotFocus()
n13.SelStart = 0
n13.SelLength = Len(n13)
End Sub
Private Sub n13_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.tc13.SetFocus
End Sub
Private Sub Nperf_GotFocus()
Nperf.SelStart = 0
Nperf.SelLength = Len(Nperf)
End Sub
Private Sub Nperf_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.cl28.SetFocus
End Sub

Private Sub rey28_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.util28.SetFocus
End Sub
Private Sub rey28_GotFocus()
If Not IsNumeric(esp28) Or esp28 = "0" Then
    MsgBox "Insira o valor da espessura conforme recomendado", vbOKOnly
    esp28.SetFocus
    esp28 = 0
    Exit Sub
End If

rey28.SelStart = 0
rey28.SelLength = Len(rey28)
End Sub
Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.pc1.SetFocus
End Sub
Private Sub t1_GotFocus()
t1.SelStart = 0
t1.SelLength = Len(t1)
End Sub
Private Sub pc1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.php1.SetFocus
End Sub
Private Sub pc1_GotFocus()
pc1.SelStart = 0
pc1.SelLength = Len(pc1)
End Sub

Private Sub php1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.pm1.SetFocus
End Sub
Private Sub php1_GotFocus()
If Not IsNumeric(pc1) Or pc1 = "0" Then
    MsgBox "Insira o valor de Ci conforme recomendado", vbOKOnly
    pc1.SetFocus
    pc1 = 0
    Exit Sub
End If

php1.SelStart = 0
php1.SelLength = Len(php1)
End Sub
Private Sub pm1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.prpm1.SetFocus
End Sub
Private Sub pm1_GotFocus()
If Not IsNumeric(php1) Or php1 = "0" Then
    MsgBox "Insira o valor de Hp conforme recomendado", vbOKOnly
    php1.SetFocus
    php1 = 0
    Exit Sub
End If

pm1.SelStart = 0
pm1.SelLength = Len(pm1)
End Sub
Private Sub prpm1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then
Form28.Command10.Enabled = True
Form28.Command10.SetFocus
End If
End Sub
Private Sub prpm1_GotFocus()
If Not IsNumeric(pm1) Or pm1 = "0" Then
    MsgBox "Insira o valor do Peso conforme recomendado", vbOKOnly
    pm1.SetFocus
    pm1 = 0
    Exit Sub
End If
prpm1.SelStart = 0
prpm1.SelLength = Len(php1)
End Sub

Private Sub tc13_GotFocus()
tc13.SelStart = 0
tc13.SelLength = Len(tc13)
End Sub
Private Sub tc13_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.e13.SetFocus
End Sub

Private Sub tcz13_GotFocus()
If Not IsNumeric(tt13) Or tt13 = "0" Then
    MsgBox "Insira o valor da tensão conforme recomendado", vbOKOnly
    tt13.SetFocus
    tt13 = 0
    Exit Sub
End If

tcz13.SelStart = 0
tcz13.SelLength = Len(tcz13)
End Sub
Private Sub tcz13_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then
Form28.Command4.Enabled = True
Form28.Command4.SetFocus
End If
End Sub

Private Sub tt13_GotFocus()
If Not IsNumeric(e13) Or e13 = "0" Then
    MsgBox "Insira o valor do módulo conforme recomendado", vbOKOnly
    e13.SetFocus
    e13 = 0
    Exit Sub
End If

tt13.SelStart = 0
tt13.SelLength = Len(tt13)
End Sub
Private Sub tt13_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Form28.tcz13.SetFocus
End Sub

Private Sub util28_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then KeyAscii = 13
End Sub
Private Sub util28_GotFocus()
If Not IsNumeric(rey28) Or rey28 = "0" Then
    MsgBox "Insira o valor de Reynolds conforme recomendado", vbOKOnly
    rey28.SetFocus
    rey28 = 0
    Exit Sub
End If

util28.SelStart = 0
util28.SelLength = Len(util28)
End Sub

