VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form Tipo 
   Caption         =   "Form35"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   11880
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form35"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      ForeColor       =   &H00FF0000&
      Height          =   8055
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   11355
      Begin VB.TextBox av1 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   4440
         Width           =   2775
      End
      Begin VB.TextBox av2 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   4800
         Width           =   2775
      End
      Begin VB.TextBox av3 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   5160
         Width           =   2775
      End
      Begin VB.TextBox av4 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   5520
         Width           =   2775
      End
      Begin VB.TextBox av5 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   5880
         Width           =   2775
      End
      Begin VB.TextBox av6 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   6240
         Width           =   2775
      End
      Begin VB.TextBox av7 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   6600
         Width           =   2775
      End
      Begin VB.TextBox av8 
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   6960
         Width           =   2775
      End
      Begin VB.TextBox av9 
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   4440
         Width           =   2775
      End
      Begin VB.TextBox av10 
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   4800
         Width           =   2775
      End
      Begin VB.Frame Frame3 
         Caption         =   "Categoria de aeronave - Monomotores"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   280
         TabIndex        =   21
         Top             =   720
         Width           =   10815
         Begin VB.OptionButton Option1 
            Caption         =   "Treinador Primário"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Treinador Básico"
            Height          =   255
            Left            =   2040
            TabIndex        =   25
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Treinador Primário Básico"
            Height          =   255
            Left            =   3840
            TabIndex        =   24
            Top             =   360
            Width           =   2175
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Treinador  Básico Avançado"
            Height          =   255
            Left            =   6240
            TabIndex        =   23
            Top             =   360
            Width           =   2415
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Treinador Avançado"
            Height          =   255
            Left            =   8880
            TabIndex        =   22
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Biblioteca de aeronaves"
         ForeColor       =   &H00FF0000&
         Height          =   2655
         Left            =   280
         TabIndex        =   18
         Top             =   1560
         Width           =   10815
         Begin VB.PictureBox Picture2 
            Height          =   2100
            Left            =   360
            ScaleHeight     =   2040
            ScaleWidth      =   9975
            TabIndex        =   19
            Top             =   360
            Width           =   10035
            Begin MSDBGrid.DBGrid DBGrid1 
               Bindings        =   "Tipo.frx":0000
               Height          =   2055
               Left            =   0
               OleObjectBlob   =   "Tipo.frx":0014
               TabIndex        =   20
               Top             =   0
               Width           =   9975
            End
            Begin VB.Data Data1 
               Caption         =   "Data1"
               Connect         =   "Access"
               DatabaseName    =   "E:\Arquivos de programas\PCA\TRmil.mdb"
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   345
               Left            =   1080
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   "Plan1"
               Top             =   1440
               Width           =   1140
            End
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   435
         Left            =   480
         ScaleHeight     =   375
         ScaleWidth      =   10395
         TabIndex        =   13
         Top             =   7440
         Width           =   10455
         Begin VB.CommandButton Command1 
            Caption         =   "OK !"
            Height          =   375
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   2775
         End
         Begin VB.CommandButton Command111 
            Caption         =   "Cancelar escolha "
            Height          =   375
            Left            =   2760
            TabIndex        =   16
            Top             =   0
            Width           =   2535
         End
         Begin VB.CommandButton Command112 
            Caption         =   "Salvar"
            Height          =   375
            Left            =   5280
            TabIndex        =   15
            Top             =   0
            Width           =   2535
         End
         Begin VB.CommandButton Command113 
            Caption         =   "Remover todas aeronaves"
            Height          =   375
            Left            =   7800
            TabIndex        =   14
            Top             =   0
            Width           =   2610
         End
      End
      Begin VB.TextBox av11 
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   5160
         Width           =   2775
      End
      Begin VB.TextBox av12 
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   5520
         Width           =   2775
      End
      Begin VB.TextBox av14 
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   6240
         Width           =   2775
      End
      Begin VB.TextBox av13 
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   5880
         Width           =   2775
      End
      Begin VB.TextBox av15 
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   6600
         Width           =   2775
      End
      Begin VB.TextBox av16 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8160
         TabIndex        =   7
         Top             =   6960
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   3480
         TabIndex        =   5
         Top             =   5040
         Width           =   4455
         Begin VB.HScrollBar Barra 
            DragMode        =   1  'Automatic
            Height          =   255
            Left            =   120
            Max             =   16
            TabIndex        =   39
            Top             =   240
            Width           =   4215
         End
         Begin VB.TextBox escolha 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            Height          =   285
            Left            =   1800
            TabIndex        =   6
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Nome do Projeto"
         Height          =   855
         Left            =   3480
         TabIndex        =   3
         Top             =   6240
         Width           =   4455
         Begin VB.OptionButton PCA 
            Caption         =   "PCA"
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox av17 
            Height          =   285
            Left            =   1440
            TabIndex        =   4
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   315
         Left            =   3480
         ScaleHeight     =   255
         ScaleWidth      =   4395
         TabIndex        =   1
         Top             =   4560
         Width           =   4455
         Begin VB.CommandButton Command2 
            Caption         =   "Adicionar barras de rolagem"
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   4400
         End
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "E:\Arquivos de programas\PCA\ANVUSP.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3600
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   $"Tipo.frx":4377
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   8715
      End
   End
   Begin VB.Menu filefechar 
      Caption         =   "Fechar"
   End
   Begin VB.Menu fileir 
      Caption         =   "Ir para"
      Begin VB.Menu filealc 
         Caption         =   "Alcance"
      End
   End
End
Attribute VB_Name = "Tipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub av1_GotFocus()
av1.SelStart = 0
av1.SelLength = Len(av1)
End Sub
Private Sub av2_GotFocus()
av2.SelStart = 0
av2.SelLength = Len(av2)
End Sub
Private Sub av3_GotFocus()
av3.SelStart = 0
av3.SelLength = Len(av3)
End Sub
Private Sub av4_GotFocus()
av4.SelStart = 0
av4.SelLength = Len(av4)
End Sub
Private Sub av5_GotFocus()
av5.SelStart = 0
av5.SelLength = Len(av5)
End Sub
Private Sub av6_GotFocus()
av6.SelStart = 0
av6.SelLength = Len(av6)
End Sub
Private Sub av7_GotFocus()
av7.SelStart = 0
av7.SelLength = Len(av7)
End Sub
Private Sub av8_GotFocus()
av8.SelStart = 0
av8.SelLength = Len(av8)
End Sub
Private Sub av9_GotFocus()
av9.SelStart = 0
av9.SelLength = Len(av9)
End Sub
Private Sub av10_GotFocus()
av10.SelStart = 0
av10.SelLength = Len(av10)
End Sub
Private Sub av11_GotFocus()
av11.SelStart = 0
av11.SelLength = Len(av11)
End Sub
Private Sub av12_GotFocus()
av12.SelStart = 0
av12.SelLength = Len(av12)
End Sub
Private Sub av13_GotFocus()
av13.SelStart = 0
av13.SelLength = Len(av13)
End Sub
Private Sub av14_GotFocus()
av14.SelStart = 0
av14.SelLength = Len(av14)
End Sub
Private Sub av15_GotFocus()
av15.SelStart = 0
av15.SelLength = Len(av15)
End Sub
Private Sub av16_GotFocus()
av16.SelStart = 0
av16.SelLength = Len(av16)
End Sub

Private Sub barra_Change()
escolha = Barra.Value
End Sub

Private Sub Command1_Click()
Tipo.Hide
Requisitos.Show
Requisitos.Option1.Value = True

End Sub

Private Sub Command111_Click()
If av1.SelLength Then av1 = Empty
If av2.SelLength Then av2 = Empty
If av3.SelLength Then av3 = Empty
If av4.SelLength Then av4 = Empty
If av5.SelLength Then av5 = Empty
If av6.SelLength Then av6 = Empty
If av7.SelLength Then av7 = Empty
If av8.SelLength Then av8 = Empty
If av9.SelLength Then av9 = Empty
If av10.SelLength Then av10 = Empty
If av11.SelLength Then av11 = Empty
If av12.SelLength Then av12 = Empty
If av13.SelLength Then av13 = Empty
If av14.SelLength Then av14 = Empty
If av15.SelLength Then av15 = Empty
If av16.SelLength Then av16 = Empty

Barra.Value = Barra.Value - 1
End Sub

Private Sub Command113_Click()
av1 = Empty
av2 = Empty
av3 = Empty
av4 = Empty
av5 = Empty
av6 = Empty
av7 = Empty
av8 = Empty
av9 = Empty
av10 = Empty
av11 = Empty
av12 = Empty
av13 = Empty
av14 = Empty
av15 = Empty
av16 = Empty


WS = Empty


Rlinear = Empty
Rlinearc = Empty
Rlinearct = Empty

Barra.Value = 0

End Sub

Private Sub Command2_Click()
DBGrid1.Splits(0).ScrollBars = 3
Tipo.Data1.Recordset.MoveNext
End Sub




Private Sub DBGrid1_DblClick()
If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False And Option5.Value = False Then
MsgBox " Escolha uma categoria de aeronaves", vbInformation
Exit Sub
End If

If av1.TabStop And av1 = Empty Then
    If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
    'Continua > Or Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 1
av1 = Data1.Recordset("aeronave")
av2.SetFocus
Exit Sub
End If

If av2.TabStop And av2 = Empty Then
    If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
    'continua Or Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
    
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 2
av2 = Data1.Recordset("aeronave")
av3.SetFocus
Exit Sub
End If

If av3.TabStop And av3 = Empty Then
    If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
    'continua Or Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
    
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 3
av3 = Data1.Recordset("aeronave")
av4.SetFocus
Exit Sub
End If

If av4.TabStop And av4 = Empty Then
    If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
    'continua Or Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
    
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 4
av4 = Data1.Recordset("aeronave")
av5.SetFocus
Exit Sub
End If

If av5.TabStop And av5 = Empty Then
    If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
    'continua Or Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
    
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 5
av5 = Data1.Recordset("aeronave")
av6.SetFocus
Exit Sub
End If

If av6.TabStop And av6 = Empty Then
   If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
   'continua Or Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
   
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 6
av6 = Data1.Recordset("aeronave")
av7.SetFocus
Exit Sub
End If

If av7.TabStop And av7 = Empty Then
    If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
    'continua Or Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
    
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 7
av7 = Data1.Recordset("aeronave")
av8.SetFocus
Exit Sub
End If

If av8.TabStop And av8 = Empty Then
   If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
   'continua Or Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
   
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 8
av8 = Data1.Recordset("aeronave")
av9.SetFocus
Exit Sub
End If

If av9.TabStop And av9 = Empty Then
   If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
   'continuaOr Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
   
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 9
av9 = Data1.Recordset("aeronave")
av10.SetFocus
Exit Sub
End If

If av10.TabStop And av10 = Empty Then
   If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
   'continuaOr Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
   
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 10
av10 = Data1.Recordset("aeronave")
av11.SetFocus
Exit Sub
End If

If av11.TabStop And av11 = Empty Then
   If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
   'continua Or Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
   
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 11
av11 = Data1.Recordset("aeronave")
av12.SetFocus
Exit Sub
End If

If av12.TabStop And av12 = Empty Then
    If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
    'continuaOr Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
    
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 12
av12 = Data1.Recordset("aeronave")
av13.SetFocus
Exit Sub
End If

If av13.TabStop And av13 = Empty Then
    If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
    'continuaOr Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
    
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 13
av13 = Data1.Recordset("aeronave")
av14.SetFocus
Exit Sub
End If

If av14.TabStop And av14 = Empty Then
    If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
    'continua Or Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
    
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 14
av14 = Data1.Recordset("aeronave")
av15.SetFocus
Exit Sub
End If

If av15.TabStop And av15 = Empty Then
    If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
    'continuaOr Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
    
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 15
av15 = Data1.Recordset("aeronave")
av16.SetFocus
Exit Sub
End If

If av16.TabStop And av16 = Empty Then
    If Data1.Recordset("sw") = 0 Or Data1.Recordset("bw") = 0 Or Data1.Recordset("sht") = 0 Or Data1.Recordset("svt") = 0 Then
    'continua Or Data1.Recordset("sailer") = 0 Or Data1.Recordset("sprof") = 0 Or Data1.Recordset("sldirec") = 0 Or Data1.Recordset("sflap") = 0 Or Data1.Recordset("lfuse") = 0 Or Data1.Recordset("vcabine") = 0 Or Data1.Recordset("vstall") = 0 Or Data1.Recordset("vcruz") = 0 Or Data1.Recordset("vmax") = 0 Or Data1.Recordset("alcance") = 0 Or Data1.Recordset("climb") = 0 Or Data1.Recordset("ddecol") = 0 Or Data1.Recordset("daterra") = 0 Or Data1.Recordset("teto") = 0 Or Data1.Recordset("diedro") = 0 Then
    
        MsgBox "Alguns dados não estão disponíveis"
        Exit Sub
    End If
Barra.Value = 16
av16 = Data1.Recordset("aeronave")
Exit Sub
End If

End Sub


Private Sub filealc_Click()
'Entradas.Show
'Entradas.pic.Visible = True
'Entradas.Picture2.Visible = True
'Tipo.Hide

End Sub

Private Sub filefechar_Click()
Tipo.Hide
End Sub

Private Sub Form_Load()
    Data1.DatabaseName = App.Path + "\TRmil.mdb"
    Data2.DatabaseName = App.Path + "\TRmil.mdb"
    PCA.Value = True
End Sub

Private Sub Option1_Click()
    Data1.RecordSource = ("Select * from Plan1 where (Categoria = 'TRB') and (Ocupante = 2) order by PesoT")
    Data1.Refresh
    If Data1.Recordset.RecordCount < 1 Then Exit Sub
    Data1.Recordset.MoveLast
    Data1.Recordset.MoveFirst
    Requisitos.HScroll1.Value = 2
End Sub

Private Sub Option2_Click()
    Data1.RecordSource = ("Select * from Plan1 where (Categoria = 'TRP') and (Ocupante = 2) and (len(Trim(Origem)) > 0) order by PesoT")
    Data1.Refresh
    If Data1.Recordset.RecordCount < 1 Then Exit Sub
    
    Data1.Recordset.MoveLast
    Data1.Recordset.MoveFirst
    Requisitos.HScroll1.Value = 2
End Sub

Private Sub Option3_Click()
    Data1.RecordSource = ("Select * from Plan1 where (Categoria = 'TRPB') and (Ocupante = 2) and (len(Trim(Origem)) > 0) order by PesoT")
    Data1.Refresh
    If Data1.Recordset.RecordCount < 1 Then Exit Sub

    Data1.Recordset.MoveLast
    Data1.Recordset.MoveFirst
    Requisitos.HScroll1.Value = 2
End Sub

Private Sub Option4_Click()
    Data1.RecordSource = ("Select * from Plan1 where (Categoria = 'TRBA') and (Ocupante = 2) and (len(Trim(Origem)) > 0) order by PesoT")
    Data1.Refresh
    If Data1.Recordset.RecordCount < 1 Then Exit Sub

    Data1.Recordset.MoveLast
    Data1.Recordset.MoveFirst
    Requisitos.HScroll1.Value = 2
End Sub

Private Sub Option5_Click()
Data1.RecordSource = ("Select * from Plan1 where (Categoria = 'TRA') and (Ocupante = 2) order by PesoT")
    Data1.Refresh
    If Data1.Recordset.RecordCount < 1 Then Exit Sub

    Data1.Recordset.MoveLast
    Data1.Recordset.MoveFirst
End Sub

Private Sub PCA_Click()
av17 = "Pca - usp"
End Sub
