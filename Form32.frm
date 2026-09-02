VERSION 5.00
Begin VB.Form Form32 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liberação"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5190
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form32.frx":0000
   LinkTopic       =   "Form32"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   4635
      TabIndex        =   8
      Top             =   4200
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "&Liberar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   0
         MouseIcon       =   "Form32.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   0
         Width           =   4635
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Digite a chave para liberação"
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   4695
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Chave do controle"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   4695
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   3105
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data limite"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   4695
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   3105
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controle"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   3105
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Width           =   1140
   End
   Begin VB.Menu menu1 
      Caption         =   "Ajustar"
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SenhaC As String
Private Sub Command1_Click()
On Error Resume Next
    If SenhaC <> Mid(Text1, 1, 4) Then MsgBox "Você não está Autorizado", 16, "Erro de Acesso": End
    Data3.RecordSource = ("Select * from Acesso")
    Data3.Refresh
    
    If Data3.Recordset.RecordCount > 0 Then
        Data3.Recordset.MoveLast
        Dta = Data3.Recordset("Data")
    Else
        Dta = Date
    End If
    Mais = IIf(IsNumeric(Mid(Text1, 5, 2)), Mid(Text1, 5, 2), 15)
    Dta = Dta + Mais
    Data3.Recordset.AddNew
    Data3.Recordset("Data") = Dta
    Data3.Recordset("DtAce") = Date
    Data3.Recordset("Tipo") = Mid(Text1, 7, 8)
    Data3.Recordset("Tps") = IIf(IsNumeric(Mid(Text1, 15, 1)), Mid(Text1, 15, 1), 0)
    Data3.Recordset.Update
    Data3.Refresh
    End
End Sub

Private Sub Form_Load()
    Data1.DatabaseName = SR
    Data2.DatabaseName = SR
    Data3.DatabaseName = SR
    
    Exit Sub
    
    Data1.RecordSource = ("Select * from Protege")
    Data1.Refresh
    If Data1.Recordset.RecordCount < 1 Then Exit Sub
    Label1.Caption = Data1.Recordset("HD")
    
    Data2.RecordSource = ("Select * from Acesso")
    Data2.Refresh
    If Data2.Recordset.RecordCount < 1 Then Exit Sub
    Data2.Recordset.MoveLast
    Label4 = Data2.Recordset("Data")
    Inise = Val(Minute(Time)) + Val(Second(Time)) + Val(Date)
    Div# = IIf(Right(Inise, 1) <> 0, Right(Inise, 1), 1)
    SenhaC = Format(Mid(Int((Inise / Div#) + Inise), 1, 4), "0000")
    Inise = Format$(Inise, "0000")
    Label6 = Trim(Trim(Mid(VarDisc, 1, 9)) + Trim(Inise))
                
End Sub

Private Sub menu1_Click()
    resp = InputBox("Digite senha de acesso" + Chr(13) + "Passe o numero para o programador" + Chr(13) + Label6, "Senha de Acesso")
    If SenhaC <> resp Then MsgBox "Você não está Autorizado", 16, "Erro de Acesso": End
    If Dir("C:\D3dbf.dll") <> "" Then Kill ("C:\D3dbf.dll")
    WR.Close
    Set WR = DBEngine.Workspaces(0)
    Set DB = OpenDatabase(App.Path + "\Alvaro.mdb", True, False, ";PWD=AF33216028")
    DB.Execute "Drop Table Protege"
    DB.Execute "Drop Table Acesso"
    End
End Sub

