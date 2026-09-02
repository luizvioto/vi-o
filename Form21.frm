VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form Form21 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Projetos"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form21.frx":0000
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form21.frx":030A
      Height          =   1335
      Left            =   240
      OleObjectBlob   =   "Form21.frx":031A
      TabIndex        =   2
      Top             =   360
      Width           =   6615
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   1920
      Width           =   6555
      Begin VB.CommandButton Command3 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   3240
         MouseIcon       =   "Form21.frx":0EA5
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   0
         Width           =   3255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Excluir projeto"
         Height          =   375
         Left            =   2160
         MouseIcon       =   "Form21.frx":11AF
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   0
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Salvar projeto"
         Height          =   375
         Left            =   0
         MouseIcon       =   "Form21.frx":14B9
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Abrir projeto"
         Height          =   375
         Left            =   0
         MouseIcon       =   "Form21.frx":17C3
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   3255
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Salvar"
      Top             =   360
      Width           =   1140
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'On Error GoTo Final

Requisitos.Show
Requisitos.Refresh

Form1.Command1.Enabled = True
Form1.Command3.Enabled = True
Form1.Command4.Enabled = True
Form1.Command5.Enabled = True
Form1.Command6.Enabled = True
Form1.Command7.Enabled = True
Form1.Command8.Enabled = True
Form1.Command15.Enabled = True
Form1.Command16.Enabled = True
Form1.Command9.Enabled = True
'Ajudas

Form1.Command10.Enabled = True
Form1.Command11.Enabled = True
Form1.Command12.Enabled = True
Form1.Command13.Enabled = True
Form1.Command14.Enabled = True

'ZERANDO TUDO

'zerar inclusive o form2.sw
Form2.Sw = Empty

Form1.motor = 0
Form1.wo = 0

Form1.posibaixa.Value = False
Form1.posimedia.Value = False
Form1.posialta.Value = False

Form1.arw = 0
Form1.ARht = 0
Form1.ARvt = 0

Form1.Afilasa = 0
Form1.AfilaHt = 0
Form1.Afilavt = 0

Form1.clmaximo = 0
Form1.clmaxFLAP = 0
Form1.clmaxht = 0

Form1.tremconv.Value = False
Form1.tremtric.Value = False
Form1.Checkretratil.Value = Unchecked

'ZERANDO FORM4
Form4.perfil1 = Empty
Form4.perfil2 = Empty
Form4.perfil3 = Empty

Form4.semflap.Value = True
Form4.clsflap = Empty
Form4.semflap.Value = True
Form4.flap.x1 = 1800
Form4.flap.y1 = 480
Form4.flap.X2 = 2160
Form4.flap.y2 = 480
Form4.clsflap = Empty
Form4.clmaxfinal = Empty
Form4.clmaxfinal = Empty



'ZERANDO O FORM27
Form27.Option4.Value = False

variavelw = "Select * from Aerofolios where Perfil=Perfil"
Form4.Data1.RecordSource = variavelw
Form4.Data1.Refresh

variavelht = "Select * from PerfilHT where Perfil=Perfil"
Form4.Data2.RecordSource = variavelht
Form4.Data2.Refresh

variavelvt = "Select * from PerfilVT where Perfil=Perfil"
Form4.Data3.RecordSource = variavelvt
Form4.Data3.Refresh

Form4.DBGrid1.Height = 2040
Form4.DBGrid2.Height = 695
Form4.DBGrid3.Height = 695

Form4.Frame5.Visible = False
Form4.Command28.Enabled = True
Form4.Command3.Enabled = True
Form4.Command4.Enabled = True

Form4.DBGrid1.Enabled = True
Form4.DBGrid2.Enabled = True
Form4.DBGrid3.Enabled = True
Form4.Command23.Enabled = False
'FIM FORM4


Form1.hsfuse.Value = 100
Form1.hsalto.Value = 100
Form1.hslargo.Value = 100
Form1.hsvolumeht.Value = 100
Form1.hsvolumevt.Value = 100
Form1.hsnariz.Value = 100
Form1.adicionarea = 100
Form1.hsmargem.Value = 25
Form1.hsvstol.Value = 0

Form1.fuse = 0
Form1.largo = 0
Form1.alto = 0
Form1.volumeht = 0
Form1.volumevt = 0
Form1.nariz = 0
Form1.adicionarea = 0
Form1.vstol = 0

'ZERANDO O FORM13- planilha estruturas
Form13.vigaret.Value = True
Form13.matmesa = Empty
Form13.matalma = Empty
Form13.sigcad = Empty
Form13.sigtad = Empty
Form13.cizad = Empty
Form13.elasto = Empty
Form13.gsatual = Empty
Form13.gsnegatual = Empty
Form13.checkg.Value = Unchecked
Form13.Checkcgorig.Value = Unchecked
Form13.Command3.Enabled = True
Form13.Command4.Enabled = True
variavelest = "Select * from materiais where nome=nome"
Form13.Data1.RecordSource = variavelest
Form13.Data1.Refresh
'FIM FORM13



'ZERANDO FORM14

Form14.p1 = 0
Form14.d1 = 0

Form14.p2 = 0
Form14.d2 = 0

Form14.p11 = 0
Form14.d11 = 0

Form14.p12 = 0
Form14.d12 = 0

Form14.p5 = 0
Form14.d5 = 0

Form14.p6 = 0
Form14.d6 = 0

Form14.p7 = 0
Form14.d7 = 0

Form14.p10 = 0
Form14.d10 = 0

Form14.p4 = 0
Form14.d4 = 0

Form14.p3 = 0
Form14.d3 = 0

Form14.p13 = 0
Form14.d13 = 0

Form14.p14 = 0
Form14.d14 = 0

Form14.p15 = 0
Form14.d15 = 0
'FIM FORM14

'ZERANDO O FORM 27
Form27.clafaw = 0
Form27.clalfah = 0
Form27.PN27 = 0
Form27.xcg = 0
Form27.cmfuse = 0
Form27.nih = 0
Form27.shsw = 0
Form27.xachcg = 0
Form27.cecp = 0
Form27.drodf = 0
Form27.Ew = 0
Form27.dedalfa = 0
Form27.testecl = 0

Form27.Picture1.AutoRedraw = True
Form27.Picture1.Cls
Form27.kde1.Visible = False
Form27.kde2.Visible = False
Form27.kde3.Visible = False
Form27.kde4.Visible = False
Form27.kde11.Visible = False
Form27.kde22.Visible = False
Form27.kde33.Visible = False

Form27.clponto.Visible = False

Form27.cmg0.Visible = False
Form27.cmg1.Visible = False
Form27.cmg2.Visible = False
Form27.cmg3.Visible = False
Form27.cmg11.Visible = False
Form27.cmg22.Visible = False
Form27.cmg33.Visible = False
Form27.defgrau.Visible = False

Form27.cmcgok1.Visible = False

Form27.Frame3.Visible = True
'Fim form27

'ZERANDO O DESENHO

Form18.Picture = LoadPicture()
    OpenFile = ""
' Desenho

Form1.Command8.Visible = True
Form1.Command15.Visible = True
Form1.Frame13.Visible = True

Form1.alterapeso.Value = 0

'FIM ZERANDO


'aaaaaaaaaaaa INICIO DA ABERTURA DO PROJETO aaaaaaaaaaaaaaa

Form20.p1 = Form21.Data1.Recordset("Projeto")
Form20.p2 = Form21.Data1.Recordset("Projetista")
Form20.datatual = Form21.Data1.Recordset("Data")

If Form21.Data1.Recordset("Projeto") = Form20.p1 Then
Form14.Command2.Enabled = True 'Botão de Atualizar
End If

'Tipo
ktipo = Form21.Data1.Recordset("Tipo")
If ktipo = 1 Then Tipo.Option1.Value = True
If ktipo = 2 Then Tipo.Option2.Value = True
If ktipo = 3 Then Tipo.Option3.Value = True
If ktipo = 4 Then Tipo.Option4.Value = True
If ktipo = 5 Then Tipo.Option5.Value = True

'Requisitos
kmaterial = Form21.Data1.Recordset("Matcon")
If kmaterial = 1 Then Requisitos.tubo.Value = True
If kmaterial = 2 Then Requisitos.madeira.Value = True
If kmaterial = 3 Then Requisitos.metal.Value = True
If kmaterial = 4 Then Requisitos.composto.Value = True
If kmaterial = 5 Then Requisitos.Tubalum.Value = True

'Abrindo a Aeronave Base
Tipo.av1 = Form21.Data1.Recordset("aeronavebase")
Requisitos.anvbase = Form21.Data1.Recordset("aeronavebase")
Requisitos.Option1.Value = True
Requisitos.bag = Form21.Data1.Recordset("bagagem")
Requisitos.HScroll2.Value = Val(Requisitos.bag)

Requisitos.alcR = Form21.Data1.Recordset("alcR")
Requisitos.tempR = Form21.Data1.Recordset("tempR")
Requisitos.hpR = Form21.Data1.Recordset("hprR")
Requisitos.consmotczR = Form21.Data1.Recordset("consczR")
Requisitos.consmotltR = Form21.Data1.Recordset("consltR")
Requisitos.NIcruzR = Form21.Data1.Recordset("niczR")
Requisitos.NIloiterR = Form21.Data1.Recordset("niltR")
Requisitos.VcmaxR = Form21.Data1.Recordset("VmaxR")
Requisitos.vstallR = Form21.Data1.Recordset("VstallR")

kalc = Form21.Data1.Recordset("alcool")
If kalc = 1 Then Requisitos.alcool.Value = Unchecked
If kalc = 2 Then Requisitos.alcool.Value = Checked
'fim REQUISITOS

pp = Form21.Data1.Recordset("pushpull")
If pp = 1 Then Form1.pull.Value = True
If pp = 2 Then Form1.pusher.Value = True

'ABRINDO AREA DA ASA
Form2.Sw = Form21.Data1.Recordset("Areasa")
    
                
' FIM DOS DADOS DE MOTORES
Form1.motor = Form21.Data1.Recordset("Motor")
    
    'Gerais do form1
    If Data1.Recordset("Tipoasa") = 1 Then
    Form1.posialta.Value = True
    Form18.VSasamed = Form21.Data1.Recordset("pmed")
    End If
    
    If Data1.Recordset("Tipoasa") = 2 Then
    Form1.posibaixa.Value = True
    Form18.VSasamed = Form21.Data1.Recordset("pmed")
    End If
    
    If Data1.Recordset("Tipoasa") = 3 Then
    Form1.posimedia.Value = True
    End If
    
    If Data1.Recordset("Tipotrem") = 1 Then Form1.tremconv.Value = True
    If Data1.Recordset("Tipotrem") = 2 Then Form1.tremtric.Value = True
    
    If Data1.Recordset("Retratil") = 1 Then Form1.Checkretratil.Value = Checked
    If Data1.Recordset("Retratil") = 0 Then Form1.Checkretratil.Value = Unchecked
       
    Form1.arw = Data1.Recordset("ARasa")
    Form1.ARht = Data1.Recordset("ARht")
    Form1.ARvt = Data1.Recordset("ARvt")
     
    Form1.Afilasa = Data1.Recordset("AfilamW")
    Form1.AfilaHt = Data1.Recordset("Afilamht")
    Form1.Afilavt = Data1.Recordset("Afilamvt")


     
 '..................... PERFIS ..............................
     'DADOS ASA
Form4.perfil1 = Form21.Data1.Recordset("Perfilasa")
Kasa = "Select * from Aerofolios where Perfil='" & Form4.perfil1.Text + "'"
Form4.Data1.RecordSource = Kasa
    Form4.Data1.Refresh
    
        Form1.clmaximo = Form4.Data1.Recordset("clmax")
        Form1.clmaximo = Format(Form1.clmaximo, "0.00")
        Form7.alfasamax = Form4.Data1.Recordset("angulo")
        Form7.alfasazero = Form4.Data1.Recordset("angulo0")
        Form7.cdperfasa = Form4.Data1.Recordset("cdmax")
        Form2.espessura = Form4.Data1.Recordset("Espessura")
        Form2.perfasa = Form4.perfil1
        Form1.clmaxFLAP = Form21.Data1.Recordset("ClasaFlap")
        
     'Estabilizador Horizontal
Form4.perfil2 = Form21.Data1.Recordset("Perfilht")
Kht = "Select * from PerfilHT where Perfil='" & Form4.perfil2.Text + "'"
Form4.Data2.RecordSource = Kht
    Form4.Data2.Refresh
    
        Form1.clmaxht = Form4.Data2.Recordset("clmax")
        Form1.clmaxht = Format(Form1.clmaxht, "0.00")
        Form7.alfahtmax = Form4.Data2.Recordset("angulo")
        Form7.alfahtzero = Form4.Data2.Recordset("angulo0")
        Form7.cdperfht = Form4.Data2.Recordset("cdmax")
        Form2.espesht = Form4.Data2.Recordset("Espessura")
        Form2.perfht = Form4.perfil2
            
            
     'Estabilizador Vertical
Form4.perfil3 = Form21.Data1.Recordset("Perfilvt")
kvt = "Select * from PerfilVT where Perfil='" & Form4.perfil3.Text + "'"
Form4.Data3.RecordSource = kvt
    Form4.Data3.Refresh
    
        Form2.espesvt = Form4.Data3.Recordset("Espessura")
        Form2.perfvt = Form4.perfil2
        Form7.cdperfvt = Form4.Data2.Recordset("cdmax")
        
        
    Form4.origcail = Form21.Data1.Recordset("ailersalvC")
    Form4.origbail = Form21.Data1.Recordset("ailersalvB")
    Form4.origcEV = Form21.Data1.Recordset("Ldirsalv")
    Form4.origcEH = Form21.Data1.Recordset("Profsalv")
    
    Form4.hsaileron.Value = Form4.origcail
    Form4.hsbail.Value = Form4.origbail
    Form4.hsestver.Value = Form4.origcEV
    Form4.hsesthor = Form4.origcEH
        
'xxxxxxxxxxxxx ABRINDO BARRAS DE ROLAGEM XXXXXXXXXXXXXXXX
Form1.valfuse = Form21.Data1.Recordset("modfus")
Form1.hsfuse.Value = Form1.valfuse

Form1.valfuseL = Form21.Data1.Recordset("modlfus")
Form1.hslargo.Value = Form1.valfuseL

Form1.valfuseH = Form21.Data1.Recordset("modhfus")
Form1.hsalto.Value = Form1.valfuseH

Form1.volempemH = Form21.Data1.Recordset("volumH")
Form1.hsvolumeht = Form1.volempemH

Form1.volempemV = Form21.Data1.Recordset("volumV")
Form1.hsvolumevt = Form1.volempemV

Form1.marg = Form21.Data1.Recordset("margest")
Form1.hsmargem = Form1.marg
        
Form1.distanasa = Form21.Data1.Recordset("distasa")
Form1.hsnariz = Form1.distanasa

Form1.stallsalva = Form21.Data1.Recordset("stallSV")
Form1.vstolanexo = Form21.Data1.Recordset("stallanex")
Form1.vstol = Form1.vstolanexo
Form1.vstol = Format(Form1.vstol, "0.00")

Form1.Command15.Enabled = True
Form1.Command15.Visible = False

Form1.adicionarea = Form21.Data1.Recordset("aarea")
 
Form1.espbeq = Form21.Data1.Recordset("Ebequilha")
    
Form1.diambeq = Form21.Data1.Recordset("Dbequilha")
    
Form1.esptp = Form21.Data1.Recordset("Etp")
    
Form1.diamtp = Form21.Data1.Recordset("Dtp")

Form1.y2 = Form21.Data1.Recordset("altBEQ")
Form1.VScroll1.Value = 1 - (y2 * 100)
    
Form1.y1 = Form21.Data1.Recordset("altTP")
Form1.VScroll1.Value = 1 - (y1 * 100)
' xx ABRINDO PESO XXX

Form1.wo = Data1.Recordset("PESO")
Form1.wo = Format(Form1.wo, "0.00")


' .........ABRINDO FORM4 ..............
   kflap = Form21.Data1.Recordset("flap")
    
    If kflap = 1 Then Form4.semflap.Value = True
    If kflap = 2 Then Form4.plainflap.Value = True
    If kflap = 3 Then Form4.splitflap.Value = True
    If kflap = 4 Then Form4.fowlerflap.Value = True
    
   Form4.valflap = Form21.Data1.Recordset("modflap")
   Form4.VSflap = Form4.valflap
   Form4.clsflap = Form21.Data1.Recordset("cls4")
   
If Form4.semflap.Value = True Then Form4.clmaxfinal = Form4.clsflap
        

' ooooooooooooo ABRINDO PESO E BALANCEAMENTO ooooooooooooooo


If Form21.Data1.Recordset("peso1") <> Empty Then
Form14.p1 = Form21.Data1.Recordset("peso1")
Form14.d1 = Form21.Data1.Recordset("dist1")

Form14.p2 = Form21.Data1.Recordset("peso2")
Form14.d2 = Form21.Data1.Recordset("dist2")

 Form14.p11 = Form21.Data1.Recordset("peso3")
 Form14.d11 = Form21.Data1.Recordset("dist3")

 Form14.p12 = Form21.Data1.Recordset("peso4")
 Form14.d12 = Form21.Data1.Recordset("dist4")

 Form14.p5 = Form21.Data1.Recordset("peso5")
 Form14.d5 = Form21.Data1.Recordset("dist5")

 Form14.p6 = Form21.Data1.Recordset("peso6")
 Form14.d6 = Form21.Data1.Recordset("dist6")

 Form14.p7 = Form21.Data1.Recordset("peso7")
 Form14.d7 = Form21.Data1.Recordset("dist7")

 Form14.p10 = Form21.Data1.Recordset("peso8")
 Form14.d10 = Form21.Data1.Recordset("dist8")

 Form14.p4 = Form21.Data1.Recordset("peso9")
 Form14.d4 = Form21.Data1.Recordset("dist9")

 Form14.p3 = Form21.Data1.Recordset("peso10")
 Form14.d3 = Form21.Data1.Recordset("dist10")

 Form14.p13 = Form21.Data1.Recordset("peso11")
 Form14.d13 = Form21.Data1.Recordset("dist11")

Form14.p14 = Form21.Data1.Recordset("peso12")
Form14.d14 = Form21.Data1.Recordset("dist12")

Form14.p15 = Form21.Data1.Recordset("peso13")
Form14.d15 = Form21.Data1.Recordset("dist13")
End If

' eeeeeeee ABRINDO ESTRUTURAS eeeeeeeeeeeeeee
300:
If Form21.Data1.Recordset("mesa") = Empty Then
Form21.Visible = False
Exit Sub
End If

If Form21.Data1.Recordset("mesa") <> Empty Then

If Form21.Data1.Recordset("tipvig") = 1 Then Form13.vigaret.Value = True
If Form21.Data1.Recordset("tipvig") = 2 Then Form13.vigaC.Value = True
If Form21.Data1.Recordset("tipvig") = 3 Then Form13.vigacaixa.Value = True
If Form21.Data1.Recordset("tipvig") = 4 Then Form13.vigaI.Value = True

Form13.matmesa = Form21.Data1.Recordset("mesa")
Form13.matalma = Form21.Data1.Recordset("alma")

kestrut = "Select * from Materiais where Nome ='" & Form13.matmesa.Text + "'"
            Form13.Data1.RecordSource = kestrut
        Form13.Data1.Refresh
        
Form13.sigtad = Form13.Data1.Recordset("tta")
Form13.sigcad = Form13.Data1.Recordset("tca")
        Form13.Data1.Refresh

kestrut2 = "Select * from Materiais where Nome ='" & Form13.matalma.Text + "'"
            Form13.Data1.RecordSource = kestrut2
Form13.cizad = Form13.Data1.Recordset("tcza")
Form13.Data1.Refresh

Form13.elasto = Form13.Data1.Recordset("E")

Form13.Command3.Enabled = False
Form13.Command4.Enabled = False

Form1.Refresh
Form21.Refresh


Form21.Visible = False

End If


' Abrindo eficiencia HT
Form4.valefHT = Form21.Data1.Recordset("EFFHt")
Form4.Vseficht.Value = Form4.valefHT

Form4.valefHT = Form4.Vseficht.Value


' Abrindo escala de Desenho
Form18.esc18 = Form21.Data1.Recordset("escala18")
Form18.posilinha = Form21.Data1.Recordset("posiline18")

' xxxxxxxxxxxxxxxx FIM ABRIR xxxxxxxxxxxxxxxxxxxxx

'Fixando o nome do projeto
manternome = "Select * from Salvar where Projeto ='" & Form20.p1.Text & "'"
                Form21.Data1.RecordSource = manternome
                    Form21.Data1.Refresh
    Exit Sub
'Final:
    'MsgBox "Escolha um projeto existente", 16, "Erro na escolha": Exit Sub
End Sub

Private Sub Command2_Click()
If Form21.Data1.Recordset("Projeto") = "ME-01" Then
    MsgBox "Projeto exemplo não pode ser excluido"
    Exit Sub
End If

On Error Resume Next
If Form21.Data1.Recordset.RecordCount < 1 Then
    MsgBox "Não existem projetos para serem excluidos", vbInformation
    Exit Sub
End If

If Form21.Data1.Recordset.RecordCount > 0 Then
        resp21 = MsgBox("Voce tem certeza de que deseja excluir o projeto selecionado?", vbOKCancel)

        If resp21 = vbOK Then
            
            ' ' 'Form21.Data1.Recordset("Projeto")
            
            Dim LocalSave As String
            Dim LocalDir As String
            LocalSave = App.Path + "\" + Form21.Data1.Recordset("Projeto") + "\"
            LocalDir = App.Path + "\" + Form21.Data1.Recordset("Projeto")
            Kill LocalSave + "*.*"
            RmDir LocalDir
                       
            ' ' '
            
            Form21.Data1.Recordset.Delete
'Acrescentei aqui porque dava bug quando excluia pois, o Data1 ficava vazio
            variavel = "Select * from Salvar where Projeto=Projeto"
            Form21.Data1.RecordSource = variavel
            Form21.Data1.Refresh
            Form21.Data1.Recordset.MoveFirst
            Form21.Visible = False
        End If
 
        If resp21 = vbCancel Then
            Form21.Visible = False
        End If
        
    End If
    
End Sub

Private Sub Command3_Click()
On Error Resume Next
Form21.Data1.Recordset.MoveFirst
Form21.Visible = False
End Sub

Private Sub Command4_Click()

'If LIBER <> "KFAR0852" Then
    'Set TB = DB.OpenRecordset("Select * from Salvar")
    'TB.MoveLast
    'TB.MoveFirst
    'If TB.RecordCount >= 50 Then
        'MsgBox "A rotina para salvar projetos sómente estará disponível após a liberação permanente", 16, "Cancelamento de impressão": Exit Sub
    'End If
'End If
If Form20.p1 = "ME-01" Then
    MsgBox "O projeto exemplo não pode ser modificado, modifique o nome do projeto para salva-lo."
    Form20.Show
    Form20.p1.SetFocus
    Exit Sub
End If

If Form13.matalma = Empty Or Form13.matmesa = Empty Then
    MsgBox "O projeto deve ser totalmente calculado!", vbInformation
    Form21.Visible = False
    Exit Sub
End If

procproj = "projeto= '" & Form20.p1 & "'"
Form21.Data1.Recordset.FindFirst procproj


If Form21.Data1.Recordset.NoMatch Then
Else
    respja = MsgBox(" Projeto já existente, deseja atualiza-lo???", vbYesNo)
        If respja = vbYes Then
            If Form13.matalma = Empty Or Form13.gsatual = Empty Then
                MsgBox "O projeto deve ser totalmente calculado!", vbInformation
                Form21.Visible = False
                Exit Sub
            End If
            GoTo Salva
        End If
End If

If respja = vbNo Then
    Form20.Visible = True
    Form20.p1.SetFocus
    MsgBox "Modifique o nome do projeto ou, acrescente subscrito; ex. ""Modelo-1A"""
    Exit Sub
End If
            
procproj = "projeto= '" & Form20.p1 & "'"
Form21.Data1.Recordset.FindFirst procproj

If Form21.Data1.Recordset.NoMatch Then
    Form21.Data1.Recordset.AddNew
 Else
Salva:
    Form21.Data1.Recordset.Edit
End If

            ' ...... SALVANDO .......
            
            '  xxxx SALVANDO DADOS INICIAIS xxx
    Form21.Data1.Recordset("Projeto") = Form20.p1
    Form21.Data1.Recordset("Projetista") = Form20.p2
    Form21.Data1.Recordset("Data") = Form20.datatual
    Form21.Data1.Recordset("Motor") = Form1.motor
        
' Tipo
If Tipo.Option1.Value = True Then ktipo = 1
If Tipo.Option2.Value = True Then ktipo = 2
If Tipo.Option3.Value = True Then ktipo = 3
If Tipo.Option4.Value = True Then ktipo = 4
If Tipo.Option5.Value = True Then ktipo = 5
Form21.Data1.Recordset("Tipo") = ktipo

'Requisitos
Form21.Data1.Recordset("aeronavebase") = Requisitos.anvbase
Form21.Data1.Recordset("bagagem") = Requisitos.bag

If Requisitos.tubo.Value = True Then kmaterial = 1
If Requisitos.madeira.Value = True Then kmaterial = 2
If Requisitos.metal.Value = True Then kmaterial = 3
If Requisitos.composto.Value = True Then kmaterial = 4
If Requisitos.Tubalum.Value = True Then kmaterial = 5
Form21.Data1.Recordset("Matcon") = kmaterial

If Requisitos.alcool.Value = Unchecked Then kalc = 1
If Requisitos.alcool.Value = Checked Then kalc = 2
Form21.Data1.Recordset("alcool") = kalc

Form21.Data1.Recordset("alcR") = Requisitos.alcR
Form21.Data1.Recordset("tempR") = Requisitos.tempR
Form21.Data1.Recordset("hprR") = Requisitos.hpR
Form21.Data1.Recordset("consczR") = Requisitos.consmotczR
Form21.Data1.Recordset("consltR") = Requisitos.consmotltR
Form21.Data1.Recordset("niczR") = Requisitos.NIcruzR
Form21.Data1.Recordset("niltR") = Requisitos.NIloiterR
Form21.Data1.Recordset("VmaxR") = Requisitos.VcmaxR
Form21.Data1.Recordset("VstallR") = Requisitos.vstallR
If Form1.pull.Value = True Then pp = 1
If Form1.pusher.Value = True Then pp = 2
Form21.Data1.Recordset("pushpull") = pp
            
'SALVANDO A AREA DA ASA
    Form21.Data1.Recordset("Areasa") = Form2.Sw
            
    
    If Form1.posialta.Value = True Then
    Pasa = 1
    Form21.Data1.Recordset("pmed") = 0
    End If
    
    If Form1.posibaixa.Value = True Then
    Pasa = 2
    Form21.Data1.Recordset("pmed") = 0
    End If
    
    If Form1.posimedia.Value = True Then
    Pasa = 3
    If Form18.VSasamed <> Empty Then Form21.Data1.Recordset("pmed") = Form18.VSasamed
    If Form18.VSasamed = Empty Then Form21.Data1.Recordset("pmed") = 6
    End If
    
    Form21.Data1.Recordset("Tipoasa") = Pasa
    
    If Form1.tremconv.Value = True Then ttrem = 1
    If Form1.tremtric.Value = True Then ttrem = 2
    Form21.Data1.Recordset("Tipotrem") = ttrem
    
    If Form1.Checkretratil.Value = Checked Then fixret = 1
    If Form1.Checkretratil.Value = Unchecked Then fixret = 0
    Form21.Data1.Recordset("Retratil") = fixret
    
     Form21.Data1.Recordset("ARasa") = Form1.arw
     Form21.Data1.Recordset("ARht") = Form1.ARht
     Form21.Data1.Recordset("ARvt") = Form1.ARvt
     
     Form21.Data1.Recordset("AfilamW") = Form1.Afilasa
     Form21.Data1.Recordset("Afilamht") = Form1.AfilaHt
     Form21.Data1.Recordset("Afilamvt") = Form1.Afilavt
     
     '  xxx SALVANDO DADOS DE ASA/Ht/Vt - FORM4 XXX
    Form21.Data1.Recordset("Perfilasa") = Form4.perfil1
    Form21.Data1.Recordset("Perfilht") = Form4.perfil2
    Form21.Data1.Recordset("Perfilvt") = Form4.perfil3
    
    Form21.Data1.Recordset("Clasa") = Form1.clmaximo
    Form21.Data1.Recordset("Clasaflap") = Form1.clmaxFLAP
    Form21.Data1.Recordset("ClHT") = Form1.clmaxht
    
    'xxxxxxxxx SALVANDO AS BARRASA DE ROLAGENS XXXXXXXXXXX
    Form21.Data1.Recordset("modfus") = Form1.valfuse
    
    Form21.Data1.Recordset("modlfus") = Form1.valfuseL
    
    Form21.Data1.Recordset("modhfus") = Form1.valfuseH
    
    Form21.Data1.Recordset("volumH") = Form1.volempemH
    
    Form21.Data1.Recordset("volumV") = Form1.volempemV
    
    Form21.Data1.Recordset("margest") = Form1.marg
    
    Form21.Data1.Recordset("PESO") = Form1.wo
    
    Form21.Data1.Recordset("distasa") = Form1.distanasa
    
    Form21.Data1.Recordset("stallSV") = Form1.stallsalva
    Form21.Data1.Recordset("stallanex") = Form1.vstolanexo
    
    Form21.Data1.Recordset("aarea") = Form1.adicionarea
    
    Form21.Data1.Recordset("Ebequilha") = Form1.espbeq
    
    Form21.Data1.Recordset("Dbequilha") = Form1.diambeq
    
    Form21.Data1.Recordset("Etp") = Form1.esptp
    
    Form21.Data1.Recordset("Dtp") = Form1.diamtp
    
    Form21.Data1.Recordset("altBEQ") = Form1.y2
    
    Form21.Data1.Recordset("altTP") = Form1.y1
    
    '..........SALVANDO FORM4 .............
    If Form4.semflap.Value = True Then kflap = 1
    If Form4.plainflap.Value = True Then kflap = 2
    If Form4.splitflap.Value = True Then kflap = 3
    If Form4.fowlerflap.Value = True Then kflap = 4
    
    Form21.Data1.Recordset("ailersalvC") = Form4.origcail
    Form21.Data1.Recordset("ailersalvB") = Form4.origbail
    Form21.Data1.Recordset("Ldirsalv") = Form4.origcEV
    Form21.Data1.Recordset("Profsalv") = Form4.origcEH
    
    
    Form21.Data1.Recordset("flap") = kflap
    Form21.Data1.Recordset("modflap") = Form4.valflap
    Form21.Data1.Recordset("cls4") = Form4.clsflap
    
     'Salvando eficiencia do HT
    Form21.Data1.Recordset("EFFHt") = Form4.valefHT
    
    
    'sssssssssssss SALVANDO TABELA DE SERVOS SSSSSSSSSSSSSSSS
    

    
Form21.Data1.Recordset("peso1") = Form14.p1
Form21.Data1.Recordset("dist1") = Form14.d1

Form21.Data1.Recordset("peso2") = Form14.p2
Form21.Data1.Recordset("dist2") = Form14.d2

Form21.Data1.Recordset("peso3") = Form14.p11
Form21.Data1.Recordset("dist3") = Form14.d11

Form21.Data1.Recordset("peso4") = Form14.p12
Form21.Data1.Recordset("dist4") = Form14.d12

Form21.Data1.Recordset("peso5") = Form14.p5
Form21.Data1.Recordset("dist5") = Form14.d5

Form21.Data1.Recordset("peso6") = Form14.p6
Form21.Data1.Recordset("dist6") = Form14.d6

Form21.Data1.Recordset("peso7") = Form14.p7
Form21.Data1.Recordset("dist7") = Form14.d7

Form21.Data1.Recordset("peso8") = Form14.p10
Form21.Data1.Recordset("dist8") = Form14.d10

Form21.Data1.Recordset("peso9") = Form14.p4
Form21.Data1.Recordset("dist9") = Form14.d4

Form21.Data1.Recordset("peso10") = Form14.p3
Form21.Data1.Recordset("dist10") = Form14.d3

Form21.Data1.Recordset("peso11") = Form14.p13
Form21.Data1.Recordset("dist11") = Form14.d13

Form21.Data1.Recordset("peso12") = Form14.p14
Form21.Data1.Recordset("dist12") = Form14.d14

Form21.Data1.Recordset("peso13") = Form14.p15
Form21.Data1.Recordset("dist13") = Form14.d15




'If Form14.Option1.Value = True Then
'Form21.Data1.Recordset("Oauto") = 1
'Form21.Data1.Recordset("Omanual") = 0
'End If

'If Form14.Option2.Value = True Then
'Form21.Data1.Recordset("Oauto") = 0
'Form21.Data1.Recordset("Omanual") = 1
'End If

'SERVOS


'eeeeeeeeeeeeeeeeeee SALVANDO ESTRUTURAS eeeeeeeeeeeeeeeeee
If Form13.vigaret.Value = True Then Form21.Data1.Recordset("tipvig") = 1
If Form13.vigaC.Value = True Then Form21.Data1.Recordset("tipvig") = 2
If Form13.vigacaixa.Value = True Then Form21.Data1.Recordset("tipvig") = 3
If Form13.vigaI.Value = True Then Form21.Data1.Recordset("tipvig") = 4

If Form13.Checkcgorig.Value = True Then Form21.Data1.Recordset("checkCG") = 1
If Form13.checkg.Value = True Then Form21.Data1.Recordset("checkCG") = 0

Form21.Data1.Recordset("CGatual") = Form13.gsatual
Form21.Data1.Recordset("CGnega") = Form13.gsnegatual
Form21.Data1.Recordset("Hlong") = Form13.porclong

Form21.Data1.Recordset("mesa") = Form13.matmesa
Form21.Data1.Recordset("alma") = Form13.matalma

'SALVANDO ESCALA DO DESENHO e Configurações
Form21.Data1.Recordset("escala18") = Form18.esc18
Form21.Data1.Recordset("posiline18") = Form18.posilinha

If Form18.hsasaval = Empty Then Form18.hsasaval = 0
If Form18.hsasaval <> Empty Then Form21.Data1.Recordset("f18hs1") = Form18.hsasaval

If Form18.hshtval = Empty Then Form18.hshtval = 0
If Form18.hshtval <> Empty Then Form21.Data1.Recordset("f18hs2") = Form18.hshtval

Form21.Data1.Recordset("f18hs3") = Form18.gradvt
Form21.Data1.Recordset("f18hs4") = Form18.vt

Form21.Data1.Recordset("Fo18hs4") = Form18.hs4

    Form21.Data1.Recordset.Update
    Form21.Data1.Refresh
    
varsalv = "Select * from Salvar where Projeto=Projeto"
Form21.Data1.RecordSource = varsalv
Form21.Data1.Refresh
    
End Sub

Private Sub Form_Load()
Command1.Left = 0
Command2.Left = 0
Command4.Left = 0
Command3.Left = 3240
    Data1.DatabaseName = SR
End Sub
