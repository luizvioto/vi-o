VERSION 5.00
Begin VB.Form Form29 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Conversor & Calculadora"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8535
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form29.frx":0000
   LinkTopic       =   "Form29"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3495
      Left            =   240
      TabIndex        =   29
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command20 
         Height          =   255
         Left            =   2520
         MouseIcon       =   "Form29.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Form29.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Transfere valores convertidos para nova conversão"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton limpar 
         Caption         =   "C"
         Height          =   255
         Left            =   3360
         MouseIcon       =   "Form29.frx":091E
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Limpa tudo"
         Top             =   960
         Width           =   735
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Unidades de :"
         Height          =   1215
         Left            =   180
         TabIndex        =   39
         Top             =   1200
         Width           =   3975
         Begin VB.OptionButton velocidade 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Velocidade"
            Height          =   375
            Left            =   240
            TabIndex        =   45
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton volume 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Volume"
            Height          =   375
            Left            =   1800
            TabIndex        =   44
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton força 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Força"
            Height          =   375
            Left            =   3000
            TabIndex        =   43
            Top             =   720
            Width           =   735
         End
         Begin VB.OptionButton massa 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Massa"
            Height          =   255
            Left            =   3000
            TabIndex        =   42
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton area 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Área"
            Height          =   255
            Left            =   1800
            TabIndex        =   41
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton comprimento 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Comprimento"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   180
         ScaleHeight     =   435
         ScaleWidth      =   3915
         TabIndex        =   35
         Top             =   2520
         Width           =   3975
         Begin VB.CommandButton Command27 
            Caption         =   "&Para rec. parafuso"
            Height          =   435
            Left            =   1920
            MouseIcon       =   "Form29.frx":0C28
            MousePointer    =   99  'Custom
            TabIndex        =   50
            ToolTipText     =   "Envia valores para recuperação de parafusos"
            Top             =   0
            Width           =   1995
         End
         Begin VB.CommandButton Command21 
            Caption         =   "&Converter"
            Height          =   435
            Left            =   0
            MouseIcon       =   "Form29.frx":0F32
            MousePointer    =   99  'Custom
            TabIndex        =   49
            ToolTipText     =   "Converte unidades"
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.TextBox aconvert 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   32
         Text            =   "0,000"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox convert 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0,000"
         Top             =   720
         Width           =   855
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   645
         ItemData        =   "Form29.frx":123C
         Left            =   2520
         List            =   "Form29.frx":123E
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Converter :"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resulta :"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   630
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   3495
      Left            =   4800
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.TextBox difpeso 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   54
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Rc"
         Height          =   375
         Left            =   1440
         MouseIcon       =   "Form29.frx":1240
         MousePointer    =   99  'Custom
         TabIndex        =   47
         ToolTipText     =   "Reclama a memória"
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton Command25 
         Caption         =   "M"
         Height          =   375
         Left            =   840
         MouseIcon       =   "Form29.frx":154A
         MousePointer    =   99  'Custom
         TabIndex        =   46
         ToolTipText     =   "Insere na memória"
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Cvs"
         Height          =   375
         Left            =   240
         MouseIcon       =   "Form29.frx":1854
         MousePointer    =   99  'Custom
         TabIndex        =   38
         ToolTipText     =   "Envia valores para o conversor de unidades"
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton Command23 
         Caption         =   "1/x"
         Height          =   375
         Left            =   2760
         MouseIcon       =   "Form29.frx":1B5E
         MousePointer    =   99  'Custom
         TabIndex        =   37
         ToolTipText     =   "Inverso"
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton Command22 
         Caption         =   "+/-"
         Height          =   375
         Left            =   2160
         MouseIcon       =   "Form29.frx":1E68
         MousePointer    =   99  'Custom
         TabIndex        =   36
         ToolTipText     =   "Inverte o sinal"
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton Command18 
         Caption         =   "%"
         Height          =   375
         Left            =   2760
         MouseIcon       =   "Form29.frx":2172
         MousePointer    =   99  'Custom
         TabIndex        =   27
         ToolTipText     =   "Porcentagem"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command28 
         Caption         =   "C"
         Height          =   375
         Left            =   2760
         MouseIcon       =   "Form29.frx":247C
         MousePointer    =   99  'Custom
         TabIndex        =   26
         ToolTipText     =   "Limpa tudo"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox visor1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004000&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox visor2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton Command19 
         Caption         =   "+"
         Height          =   375
         Left            =   2160
         MouseIcon       =   "Form29.frx":2786
         MousePointer    =   99  'Custom
         TabIndex        =   20
         ToolTipText     =   "Adição"
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command17 
         Caption         =   "-"
         Height          =   375
         Left            =   2160
         MouseIcon       =   "Form29.frx":2A90
         MousePointer    =   99  'Custom
         TabIndex        =   19
         ToolTipText     =   "Subtração"
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Command16 
         Caption         =   "x^2"
         Height          =   375
         Left            =   2760
         MouseIcon       =   "Form29.frx":2D9A
         MousePointer    =   99  'Custom
         TabIndex        =   18
         ToolTipText     =   "Eleva ao quadrado"
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Command15 
         Caption         =   "x"
         Height          =   375
         Left            =   2160
         MouseIcon       =   "Form29.frx":30A4
         MousePointer    =   99  'Custom
         TabIndex        =   17
         ToolTipText     =   "Multiplicação"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Raiz"
         Height          =   375
         Left            =   2760
         MouseIcon       =   "Form29.frx":33AE
         MousePointer    =   99  'Custom
         TabIndex        =   16
         ToolTipText     =   "Raiz quadrada"
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command13 
         Caption         =   "/"
         Height          =   375
         Left            =   2160
         MouseIcon       =   "Form29.frx":36B8
         MousePointer    =   99  'Custom
         TabIndex        =   15
         ToolTipText     =   "Divisão"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command12 
         Caption         =   "="
         Height          =   375
         Left            =   1440
         MouseIcon       =   "Form29.frx":39C2
         MousePointer    =   99  'Custom
         TabIndex        =   14
         ToolTipText     =   "Igual"
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   ","
         Height          =   375
         Left            =   840
         MouseIcon       =   "Form29.frx":3CCC
         MousePointer    =   99  'Custom
         TabIndex        =   13
         ToolTipText     =   "Separador (ponto/virgula)"
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "0"
         Height          =   375
         Left            =   240
         MouseIcon       =   "Form29.frx":3FD6
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "3"
         Height          =   375
         Left            =   1440
         MouseIcon       =   "Form29.frx":42E0
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "2"
         Height          =   375
         Left            =   840
         MouseIcon       =   "Form29.frx":45EA
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "1"
         Height          =   375
         Left            =   240
         MouseIcon       =   "Form29.frx":48F4
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "6"
         Height          =   375
         Left            =   1440
         MouseIcon       =   "Form29.frx":4BFE
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "5"
         Height          =   375
         Left            =   840
         MouseIcon       =   "Form29.frx":4F08
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "4"
         Height          =   375
         Left            =   240
         MouseIcon       =   "Form29.frx":5212
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "9"
         Height          =   375
         Left            =   1440
         MouseIcon       =   "Form29.frx":551C
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "8"
         Height          =   375
         Left            =   840
         MouseIcon       =   "Form29.frx":5826
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         Height          =   375
         Left            =   240
         MouseIcon       =   "Form29.frx":5B30
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox memo 
         Alignment       =   2  'Center
         Height          =   195
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença de Peso ="
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   5280
      TabIndex        =   21
      Top             =   480
      Width           =   1815
      Begin VB.OptionButton porcento 
         Caption         =   "porcento"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton multiplica 
         Caption         =   "multiplica"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton divide 
         Caption         =   "divide"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton soma 
         Caption         =   "soma"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton subtrai 
         Caption         =   "subtrai"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Menu filesai29 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "Form29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aconvert_GotFocus()
aconvert.SelStart = 0
aconvert.SelLength = Len(aconvert)
End Sub

Private Sub aconvert_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then Command21.SetFocus
End Sub

Private Sub Command1_Click()
visor1 = visor1 + "7"

End Sub

Private Sub Command10_Click()
visor1 = visor1 + "0"

End Sub

Private Sub Command11_Click()
visor1 = visor1 + ","
Command11.Enabled = False
End Sub

Private Sub Command12_Click()
If visor1 = "," Or visor1 = ",-" Or visor1 = "-," Then
MsgBox "entre com o Número"
visor1 = Empty
Call Command28_Click
Exit Sub
End If

If visor1 = Empty Then visor1 = 0
If visor2 = Empty Then visor2 = 0

If soma.Value = True Then visor1 = visor1 + CDbl(visor2)
If subtrai.Value = True Then visor1 = visor2 - CDbl(visor1)
If multiplica.Value = True Then visor1 = visor1 * visor2

If divide.Value And visor1 = 0 Then
    MsgBox "Divisão por zero", vbCritical
    Exit Sub
End If

If divide.Value = True And visor1 > 0 Then visor1 = visor2 / visor1

If porcento.Value = True Then visor1 = visor1 * (visor2 / 100)

End Sub

Private Sub Command13_Click()
divide.Value = True

multiplica.Value = False
soma.Value = False
subtrai.Value = False
porcento.Value = False

visor2 = visor1
visor1 = Empty

Command11.Enabled = True
End Sub

Private Sub Command14_Click()

If visor1 = "," Or visor1 = ",-" Or visor1 = "-," Then
MsgBox "entre com o Número"
visor1 = Empty
Call Command28_Click
Exit Sub
End If


If visor1 = Empty Then visor1 = 0
If visor1 < 0 Then
    MsgBox "O valor deve ser positivo", vbCritical
    Exit Sub
End If
If visor1 >= 0 Then
    visor1 = visor1 ^ 0.5
End If
Command11.Enabled = True
End Sub

Private Sub Command15_Click()
multiplica.Value = True

divide.Value = False
soma.Value = False
subtrai.Value = False
porcento.Value = False

visor2 = visor1
visor1 = Empty

Command11.Enabled = True
End Sub

Private Sub Command16_Click()
If visor1 = "," Or visor1 = ",-" Or visor1 = "-," Then
MsgBox "entre com o Número"
visor1 = Empty
Call Command28_Click
Exit Sub
End If

visor1 = visor1 ^ 2
Command11.Enabled = True
End Sub

Private Sub Command17_Click()
subtrai.Value = True

multiplica.Value = False
soma.Value = False
divide.Value = False
porcento.Value = False

visor2 = visor1
visor1 = Empty

Command11.Enabled = True
End Sub



Private Sub Command18_Click()
porcento.Value = True

multiplica.Value = False
divide.Value = False
subtrai.Value = False
soma.Value = False

visor2 = visor1
visor1 = Empty

Command11.Enabled = True
End Sub

Private Sub Command19_Click()
soma.Value = True

multiplica.Value = False
divide.Value = False
subtrai.Value = False
porcento.Value = False

visor2 = visor1
visor1 = Empty

Command11.Enabled = True
End Sub

Private Sub Command2_Click()
visor1 = visor1 + "8"

End Sub

Private Sub Command20_Click()
aconvert = convert
convert = "0,000"
End Sub

Private Sub Command21_Click()
If velocidade.Value = False And massa.Value = False And area.Value = False And comprimento.Value = False And força.Value = False And volume.Value = False Then
MsgBox "Faça a opção em ""Unidades de :"""
Exit Sub
End If

'UNIDADES DE COMPRIMENTO
If comprimento.Value = True Then
List1.Refresh
If List1.ListIndex = 0 Then convert = aconvert * 3.2808
If List1.ListIndex = 1 Then convert = aconvert / 3.2808
If List1.ListIndex = 2 Then convert = aconvert * 0.3937
If List1.ListIndex = 3 Then convert = aconvert / 0.3937
If List1.ListIndex = 4 Then convert = aconvert * 0.6214
If List1.ListIndex = 5 Then convert = aconvert / 0.6214

If List1.ListIndex = 6 Then convert = aconvert * 39.37
If List1.ListIndex = 7 Then convert = aconvert / 39.37
If List1.ListIndex = 8 Then convert = aconvert * 100
If List1.ListIndex = 9 Then convert = aconvert / 100

End If

'UNIDADES DE ÁREA
If area.Value = True Then
List1.Refresh
If List1.ListIndex = 0 Then convert = aconvert * 10.7639
If List1.ListIndex = 1 Then convert = aconvert / 10.7639
If List1.ListIndex = 2 Then convert = aconvert * 0.155
If List1.ListIndex = 3 Then convert = aconvert / 0.155
If List1.ListIndex = 4 Then convert = aconvert * 144
If List1.ListIndex = 5 Then convert = aconvert / 144

If List1.ListIndex = 6 Then convert = aconvert * 0.001076
If List1.ListIndex = 7 Then convert = aconvert / 0.001076
If List1.ListIndex = 8 Then convert = aconvert * 10000
If List1.ListIndex = 9 Then convert = aconvert / 10000
End If

'UNIDADES DE MASSA
If massa.Value = True Then
List1.Refresh
If List1.ListIndex = 0 Then convert = aconvert * 2.204
If List1.ListIndex = 1 Then convert = aconvert / 2.204
If List1.ListIndex = 2 Then convert = aconvert * 35.274
If List1.ListIndex = 3 Then convert = aconvert / 35.274
If List1.ListIndex = 4 Then convert = aconvert * 0.03527
If List1.ListIndex = 5 Then convert = aconvert / 0.03527

If List1.ListIndex = 6 Then convert = aconvert * 0
If List1.ListIndex = 7 Then convert = aconvert / 0
If List1.ListIndex = 8 Then convert = aconvert * 0
If List1.ListIndex = 9 Then convert = aconvert / 0
End If

'UNIDADES DE VELOCIDADE
If velocidade.Value = True Then
List1.Refresh
If List1.ListIndex = 0 Then convert = aconvert * 0.6213
If List1.ListIndex = 1 Then convert = aconvert / 0.6213
If List1.ListIndex = 2 Then convert = aconvert * 3.2808
If List1.ListIndex = 3 Then convert = aconvert / 3.2808
If List1.ListIndex = 4 Then convert = aconvert * 0.2778
If List1.ListIndex = 5 Then convert = aconvert / 0.2778

If List1.ListIndex = 6 Then convert = aconvert * 1.4667
If List1.ListIndex = 7 Then convert = aconvert / 1.4667
If List1.ListIndex = 8 Then convert = aconvert * 0.534
If List1.ListIndex = 9 Then convert = aconvert / 0.534
End If

'UNIDADES DE FORÇA
If força.Value = True Then
List1.Refresh
If List1.ListIndex = 0 Then convert = aconvert * 2.204
If List1.ListIndex = 1 Then convert = aconvert / 2.204
If List1.ListIndex = 2 Then convert = aconvert * 9.806
If List1.ListIndex = 3 Then convert = aconvert / 9.806
If List1.ListIndex = 4 Then convert = aconvert * 0.002204
If List1.ListIndex = 5 Then convert = aconvert / 0.002204

If List1.ListIndex = 6 Then convert = aconvert * 0
If List1.ListIndex = 7 Then convert = aconvert / 0
If List1.ListIndex = 8 Then convert = aconvert * 0
If List1.ListIndex = 9 Then convert = aconvert / 0
End If

'UNIDADES DE VOLUME
If volume.Value = True Then
List1.Refresh
If List1.ListIndex = 0 Then convert = aconvert * 35.3146
If List1.ListIndex = 1 Then convert = aconvert / 35.3146
If List1.ListIndex = 2 Then convert = aconvert * 0.06102
If List1.ListIndex = 3 Then convert = aconvert / 0.06102
If List1.ListIndex = 4 Then convert = aconvert * 1728
If List1.ListIndex = 5 Then convert = aconvert / 1728

If List1.ListIndex = 6 Then convert = aconvert * 1000000
If List1.ListIndex = 7 Then convert = aconvert / 1000000
If List1.ListIndex = 8 Then convert = aconvert * 0
If List1.ListIndex = 9 Then convert = aconvert / 0
End If


convert = Format(convert, " 0.000")
End Sub

Private Sub Command22_Click()
visor1 = "-" + visor1
Command11.Enabled = True
End Sub

Private Sub Command23_Click()
If visor1 = "," Or visor1 = ",-" Or visor1 = "-," Then
MsgBox "entre com o Número"
visor1 = Empty
Call Command28_Click
Exit Sub
End If

If visor1 = Empty Then visor1 = 0
If visor1 = 0 Then
    MsgBox "Divisão por zero", vbCritical
    Exit Sub
End If

visor1 = 1 / visor1
Command11.Enabled = True
End Sub

Private Sub Command24_Click()
aconvert = visor1
aconvert = Format(aconvert, "0.000")
Command11.Enabled = True
End Sub

Private Sub Command25_Click()
memo = visor1
visor1 = 0
End Sub

Private Sub Command26_Click()
visor1 = memo
End Sub

Private Sub Command27_Click()
Form30.sombra = Form29.visor1
Form30.Visible = True
Form29.Visible = False
End Sub

Private Sub Command28_Click()
visor1 = Empty
visor2 = Empty
Command11.Enabled = True
End Sub

Private Sub Command29_Click()
visor1 = Empty
End Sub

Private Sub Command3_Click()
visor1 = visor1 + "9"

End Sub

Private Sub Command4_Click()
visor1 = visor1 + "4"

End Sub

Private Sub Command5_Click()
visor1 = visor1 + "5"

End Sub

Private Sub Command6_Click()
visor1 = visor1 + "6"

End Sub

Private Sub Command7_Click()
visor1 = visor1 + "1"

End Sub

Private Sub Command8_Click()
visor1 = visor1 + "2"

End Sub

Private Sub Command9_Click()
visor1 = visor1 + "3"

End Sub

Private Sub comprimento_Click()
Form29.List1.Refresh
List1.AddItem "metro para pé", 0
List1.AddItem "pé para metro", 1
List1.AddItem "cm para polg.", 2
List1.AddItem "polg. para cm", 3
List1.AddItem "km para milha", 4
List1.AddItem "milha para km", 5

List1.AddItem "metro para polg.", 6
List1.AddItem "polg. para metro", 7
List1.AddItem "metro para cm.", 8
List1.AddItem "cm. para metro", 9

End Sub

Private Sub area_Click()
Form29.List1.Refresh
List1.AddItem "m2 para pé2", 0
List1.AddItem "pé2 para m2", 1
List1.AddItem "cm2 para polg2.", 2
List1.AddItem "polg2. para cm2", 3
List1.AddItem "pé2 para polg2", 4
List1.AddItem "polg2 para pé2", 5

List1.AddItem "cm2 para pé2", 6
List1.AddItem "pé2 para cm2", 7
List1.AddItem "m2 para cm2", 8
List1.AddItem "cm2 para m2", 9

End Sub

Private Sub filesai29_Click()
Form29.Visible = False
Form29.Label3.Visible = False
Form29.difpeso.Visible = False
End Sub

Private Sub força_Click()
Form29.List1.Refresh
List1.AddItem "kgf para lbf", 0
List1.AddItem "lbf para kgf", 1
List1.AddItem "kgf para N.", 2
List1.AddItem "N para kgf", 3
List1.AddItem "gr.f para lbf", 4
List1.AddItem "lbf para gr.f", 5

List1.AddItem "-", 6
List1.AddItem "-", 7
List1.AddItem "-", 8
List1.AddItem "-", 9
End Sub

Private Sub limpar_Click()
aconvert = "0,000"
convert = "0,000"
comprimento.Value = False
massa.Value = False
velocidade.Value = False
força.Value = False
area.Value = False
volume.Value = False
End Sub

Private Sub massa_Click()
Form29.List1.Refresh
List1.AddItem "kg para lb", 0
List1.AddItem "lb para kg", 1
List1.AddItem "kg para onça", 2
List1.AddItem "onça para kg", 3
List1.AddItem "grama para onça", 4
List1.AddItem "onça para grama", 5

List1.AddItem "kg para grama", 6
List1.AddItem "grama para kg", 7
List1.AddItem "lb para onça", 8
List1.AddItem "onça para lb", 9
End Sub

Private Sub volume_Click()
Form29.List1.Refresh
List1.AddItem "m3 para pé3", 0
List1.AddItem "pé3 para m3", 1
List1.AddItem "cm3 para polg3.", 2
List1.AddItem "polg3. para cm3", 3
List1.AddItem "pé3 para polg3", 4
List1.AddItem "polg3 para pé3", 5

List1.AddItem "m3 para cm3", 6
List1.AddItem "cm3 para m3", 7
List1.AddItem "-", 8
List1.AddItem "-", 9
End Sub

Private Sub velocidade_Click()
Form29.List1.Refresh
List1.AddItem "km/h para mph", 0
List1.AddItem "mph para km/h", 1
List1.AddItem "m/s para pé/s.", 2
List1.AddItem "pé/s. para m/s", 3
List1.AddItem "km/h para m/s", 4
List1.AddItem "m/s para km/h", 5

List1.AddItem "mph para pé/s", 6
List1.AddItem "pé/s para mph", 7
List1.AddItem "km/h para knot", 8
List1.AddItem "knot para km/h", 9
End Sub
