VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form34 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hélice"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   ControlBox      =   0   'False
   Icon            =   "Form34.frx":0000
   LinkTopic       =   "Form34"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar calcbar 
      Height          =   135
      Left            =   240
      TabIndex        =   48
      Top             =   5880
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   7785
      TabIndex        =   38
      Top             =   5280
      Width           =   7845
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar escolha"
         Height          =   435
         Left            =   1800
         TabIndex        =   43
         Top             =   0
         Width           =   1800
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sair"
         Height          =   435
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   0
         Width           =   1420
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calcular"
         Height          =   435
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   0
         Width           =   1800
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Enviar para projeto e sair"
         Height          =   435
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   0
         Width           =   2760
      End
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H00400000&
      Height          =   4995
      Left            =   240
      ScaleHeight     =   4935
      ScaleWidth      =   7800
      TabIndex        =   0
      Top             =   240
      Width           =   7855
      Begin VB.PictureBox Picture12 
         BackColor       =   &H00400000&
         Height          =   3450
         Left            =   360
         ScaleHeight     =   3390
         ScaleWidth      =   4755
         TabIndex        =   22
         Top             =   1320
         Width           =   4815
         Begin VB.TextBox angpa 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   3240
            TabIndex        =   42
            Text            =   "0"
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox vcrit 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "0"
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox phe 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3240
            TabIndex        =   29
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox rpmhe 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "0"
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox vhe 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "0"
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox dhe 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3240
            TabIndex        =   26
            Text            =   "0"
            Top             =   360
            Width           =   735
         End
         Begin VB.PictureBox Picture14 
            Height          =   255
            Left            =   1680
            ScaleHeight     =   195
            ScaleWidth      =   1035
            TabIndex        =   24
            Top             =   1080
            Width           =   1095
            Begin VB.HScrollBar HScroll1 
               Height          =   215
               Left            =   0
               Max             =   5
               TabIndex        =   25
               Top             =   0
               Value           =   3
               Width           =   1045
            End
         End
         Begin VB.TextBox mg 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3240
            TabIndex        =   23
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Angulo da pá da hélice :                                                  º"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   3000
            Width           =   4035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Velocidade na ponta da hélice :                                      m/s"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   2640
            Width           =   4230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Entradas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1800
            TabIndex        =   44
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RPM para a velocidade acima :                                      rpm"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   2280
            Width           =   4200
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Velocidade ótima para hélice (existente) :                       km/h"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   1920
            Width           =   4305
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "O passo da hélice existente é de :                                  poleg."
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   4365
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "O diâmetro da hélice existente é de :                              poleg."
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   4365
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Faixa de tolerancia :"
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   1425
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resultados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1680
            TabIndex        =   30
            Top             =   1680
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00400000&
         Caption         =   "Variação Rpm x Velocidade"
         ForeColor       =   &H0000FFFF&
         Height          =   4410
         Left            =   5400
         TabIndex        =   1
         Top             =   360
         Width           =   2175
         Begin VB.Label b10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   21
            Top             =   3600
            Width           =   45
         End
         Begin VB.Label b9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   20
            Top             =   3240
            Width           =   45
         End
         Begin VB.Label b8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   19
            Top             =   2880
            Width           =   45
         End
         Begin VB.Label b7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   18
            Top             =   2520
            Width           =   45
         End
         Begin VB.Label b6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   17
            Top             =   2160
            Width           =   45
         End
         Begin VB.Label b5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   16
            Top             =   1800
            Width           =   45
         End
         Begin VB.Label b4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   15
            Top             =   1440
            Width           =   45
         End
         Begin VB.Label b3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   14
            Top             =   1080
            Width           =   45
         End
         Begin VB.Label b2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   13
            Top             =   720
            Width           =   45
         End
         Begin VB.Label b1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   360
            Width           =   45
         End
         Begin VB.Label a1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   11
            Top             =   360
            Width           =   45
         End
         Begin VB.Label a2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   10
            Top             =   720
            Width           =   45
         End
         Begin VB.Label a3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   9
            Top             =   1080
            Width           =   45
         End
         Begin VB.Label a4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   8
            Top             =   1440
            Width           =   45
         End
         Begin VB.Label a5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   7
            Top             =   1800
            Width           =   45
         End
         Begin VB.Label a6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   6
            Top             =   2160
            Width           =   45
         End
         Begin VB.Label a7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   5
            Top             =   2520
            Width           =   45
         End
         Begin VB.Label a8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   4
            Top             =   2880
            Width           =   45
         End
         Begin VB.Label a9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   3
            Top             =   3240
            Width           =   45
         End
         Begin VB.Label a10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   1440
            TabIndex        =   2
            Top             =   3600
            Width           =   45
         End
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form34.frx":030A
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
         Height          =   780
         Left            =   360
         TabIndex        =   37
         Top             =   480
         Width           =   4770
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O usuário já possuí uma hélice e deseja utiliza-la."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   36
         Top             =   120
         Width           =   4245
      End
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Form34.dhe = 0
Form34.phe = 0
Form34.vhe = 0
Form34.rpmhe = 0
Form34.vcrit = 0
Form34.angpa = 0

a1 = "-"
a2 = "-"
a3 = "-"
a4 = "-"
a5 = "-"
a6 = "-"
a7 = "-"
a8 = "-"
a9 = "-"
a10 = "-"
b1 = "-"
b2 = "-"
b3 = "-"
b4 = "-"
b5 = "-"
b6 = "-"
b7 = "-"
b8 = "-"
b9 = "-"
b10 = "-"

Form6.temhel.Value = False
Form6.Command6.Enabled = False
Form6.optavest.Value = Unchecked
Form6.Vestimada = "0,00"

If Form6.temhel.Value = True Then
    Form6.Command6.Enabled = True
    Form6.hsrpm.Enabled = False
End If

If Form6.temhel.Value = False Then
    Form6.Command6.Enabled = False
    Form6.hsrpm.Enabled = True
End If

Form6.hsrpm.Value = Form6.hsrpm.Max

End Sub

Private Sub Command4_Click()

If Not IsNumeric(phe) Or phe = "0" Then
    MsgBox "Insira o passo da hélice corretamente", vbOKOnly
    phe.SetFocus
    phe = 0
    Exit Sub
End If

If Not IsNumeric(dhe) Or dhe = "0" Then
    MsgBox "Insira o diametro da hélice corretamente", vbOKOnly
    dhe.SetFocus
    dhe = 0
    Exit Sub
End If


' Começo do cálculo

'kvelok = (Int((137 * (((Form6.hporiginal * Form6.NI) / (Form2.Sw * 10.76)) ^ (1 / 3))) * 1.61)) * 2.5
'kvelok = Format(kvelok, "0.00")

kvelok = 300

calcbar.Visible = True

'LOOP
For rotacao = 100 To Form6.rpmoriginal Step 100
For velocidade = 20 To kvelok Step 1

calcbar.Max = Form6.rpmoriginal
calcbar.Value = rotacao

    passo = ((velocidade / 1.61) / (rotacao / 1056))
    passo = Format(passo, "0.0")
    diametro = (Form8.khel * ((Form6.hporiginal / ((rotacao ^ 2) * velocidade)) ^ 0.25) * (30.5 / 2.54))
    diametro = Format(diametro, "0.0")
    
'Label de teste no form16
'va = kvelok
'da = diametro
'pa = passo

If passo > Val(Form34.phe) And passo <= (Val(Form34.phe) + mg) And diametro > Val(Form34.dhe) And diametro <= (Val(Form34.dhe) + mg) Then
    
    Form34.vhe = velocidade
    rpmhe = rotacao

            
b1 = 1 * Form6.rpmoriginal
b2 = 0.9 * Form6.rpmoriginal
b3 = 0.8 * Form6.rpmoriginal
b4 = 0.7 * Form6.rpmoriginal
b5 = 0.6 * Form6.rpmoriginal
b6 = 0.5 * Form6.rpmoriginal
b7 = 0.4 * Form6.rpmoriginal
b8 = 0.3 * Form6.rpmoriginal
b9 = 0.2 * Form6.rpmoriginal
b10 = 0.1 * Form6.rpmoriginal
b1 = Format(b1, "0.00")
b2 = Format(b2, "0.00")
b3 = Format(b3, "0.00")
b4 = Format(b4, "0.00")
b5 = Format(b5, "0.00")
b6 = Format(b6, "0.00")
b7 = Format(b7, "0.00")
b8 = Format(b8, "0.00")
b9 = Format(b9, "0.00")
b10 = Format(b10, "0.00")

a1 = (passo * b1 / 1056) * 1.61
a2 = (passo * b2 / 1056) * 1.61
a3 = (passo * b3 / 1056) * 1.61
a4 = (passo * b4 / 1056) * 1.61
a5 = (passo * b5 / 1056) * 1.61
a6 = (passo * b6 / 1056) * 1.61
a7 = (passo * b7 / 1056) * 1.61
a8 = (passo * b8 / 1056) * 1.61
a9 = (passo * b9 / 1056) * 1.61
a10 = (passo * b10 / 1056) * 1.61
a1 = Format(a1, "0.00")
a2 = Format(a2, "0.00")
a3 = Format(a3, "0.00")
a4 = Format(a4, "0.00")
a5 = Format(a5, "0.00")
a6 = Format(a6, "0.00")
a7 = Format(a7, "0.00")
a8 = Format(a8, "0.00")
a9 = Format(a9, "0.00")
a10 = Format(a10, "0.00")

'Velocidade critica da hélice
Form34.vcrit = ((((3.1415 * Form34.dhe * 0.08333 * Form34.rpmhe / 60) ^ 2) + ((Form34.vhe * 0.91134) ^ 2)) ^ 0.5) * 0.3048
Form34.vcrit = Format(Form34.vcrit, "0.00")

If vcrit >= 304.8 Then MsgBox "Velocidade crítica na ponta da hélice!", vbCritical
calcbar.Visible = False

angpa = (Atn(Form34.phe / (Form34.dhe * 3.141516))) * (180 / 3.1415)
angpa = Format(angpa, "0.00")

            Exit Sub
End If

Next
Next
        If passo <> Val(Form34.phe) Or diametro <> Val(Form34.dhe) Then
            MsgBox "Aumente a tolerancia ou escolha outra hélice!", vbInformation
        End If
        

End Sub

Private Sub Command5_Click()
If vhe = 0 And rpmhe = 0 And angpa = 0 And vcrit = 0 Then
    MsgBox "Os parametros da hélice devem ser calculados. Clique calcular.", vbInformation
    Exit Sub
End If

If Not IsNumeric(phe) Or phe = "0" Then
    MsgBox "Insira o passo da hélice corretamente", vbOKOnly
    phe.SetFocus
    phe = 0
    Exit Sub
End If

If Not IsNumeric(dhe) Or dhe = "0" Then
    MsgBox "Insira o diametro da hélice corretamente", vbOKOnly
    dhe.SetFocus
    dhe = 0
    Exit Sub
End If

Form6.Visible = True
Form34.Visible = False
Form8.Visible = False
Form16.Visible = False
Form1.Visible = True
Form6.optavest.Value = Checked
Form6.Vestimada = Form34.vhe
Form6.Vestimada.Locked = True
Form6.rpm19 = Form34.rpmhe
Form6.hsrpm.Value = Form34.rpmhe
End Sub
Private Sub dhe_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then phe.SetFocus
End Sub
Private Sub dhe_GotFocus()
dhe.SelStart = 0
dhe.SelLength = Len(dhe)
End Sub

Private Sub Form_Load()
Form34.HScroll1.Value = 1
Form34.HScroll1.Value = 2
End Sub

Private Sub HScroll1_Change()
mg = Form34.HScroll1.Value / 10
End Sub

Private Sub phe_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command4.SetFocus
End Sub
Private Sub phe_GotFocus()

If Not IsNumeric(dhe) Or dhe = "0" Then
    MsgBox "Insira o diametro da hélice corretamente", vbOKOnly
    dhe.SetFocus
    dhe = 0
    Exit Sub
End If

phe.SelStart = 0
phe.SelLength = Len(phe)

End Sub

Private Sub Command6_Click()
cancela = MsgBox("Deseja manter a hélice escolhida?", vbYesNo)

            If cancela = vbYes Then
                If vhe = 0 And rpmhe = 0 And angpa = 0 And vcrit = 0 Then
                    MsgBox "Os parametros da hélice devem ser calculados. Clique calcular.", vbInformation
                    Exit Sub
                End If
                
                If vhe <> 0 And rpmhe <> 0 And angpa <> 0 And vcrit <> 0 Then Call Command5_Click
            End If
            
            If cancela = vbNo Then
                Call Command1_Click
                Form34.Visible = False
                calcbar.Visible = False

                a1 = "-"
                a2 = "-"
                a3 = "-"
                a4 = "-"
                a5 = "-"
                a6 = "-"
                a7 = "-"
                a8 = "-"
                a9 = "-"
                a10 = "-"
                b1 = "-"
                b2 = "-"
                b3 = "-"
                b4 = "-"
                b5 = "-"
                b6 = "-"
                b7 = "-"
                b8 = "-"
                b9 = "-"
                b10 = "-"
             End If
            
'If dhe = 0 And phe <> 0 Then
    'dhe = 0
    'phe = 0
'End If

'If dhe <> 0 And phe = 0 Then
    'dhe = 0
    'phe = 0
'End If

'If Form21.Data1.Recordset("Projeto") <> Form20.p1 Then
    'If dhe = 0 And phe = 0 Then
        'Form6.temhel.Value = False
        'Form6.Command6.Enabled = False
        'Form6.hsrpm.Enabled = True
    'End If
        'If dhe <> 0 And phe <> 0 Then
            'Form6.temhel.Value = True
            'Form6.Command6.Enabled = True
            'Form6.hsrpm.Enabled = False
        'End If
'End If

'If Form21.Data1.Recordset("Projeto") = Form20.p1 And Form21.Data1.Recordset("temhel") = 1 Then
    'If dhe = 0 And phe = 0 Then
        'Form6.temhel.Value = False
        'Form6.Command6.Enabled = False
        'Form6.hsrpm.Enabled = True
    'End If
        'If dhe <> 0 And phe <> 0 Then
            'Form6.temhel.Value = True
            'Form6.Command6.Enabled = True
            'Form6.hsrpm.Enabled = False
        'End If
'End If

'If Form21.Data1.Recordset("Projeto") = Form20.p1 And Form21.Data1.Recordset("temhel") = 0 Then
    'If dhe = 0 And phe = 0 Then
        'Form6.temhel.Value = False
        'Form6.Command6.Enabled = False
        'Form6.hsrpm.Enabled = True
    'End If
        
        'If dhe <> 0 And phe <> 0 Then
            'Form6.temhel.Value = True
            'Form6.Command6.Enabled = True
            'Form6.hsrpm.Enabled = False
        'End If
'End If

'Form34.Visible = False
'calcbar.Visible = False

'a1 = "-"
'a2 = "-"
'a3 = "-"
'a4 = "-"
'a5 = "-"
'a6 = "-"
'a7 = "-"
'a8 = "-"
'a9 = "-"
'a10 = "-"
'b1 = "-"
'b2 = "-"
'b3 = "-"
'b4 = "-"
'b5 = "-"
'b6 = "-"
'b7 = "-"
'b8 = "-"
'b9 = "-"
'b10 = "-"


End Sub

