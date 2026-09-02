VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Iniciar o projeto"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form20.frx":0000
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2760
      TabIndex        =   10
      Top             =   1080
      Width           =   3975
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Projeto Conceitual de Aeronave - PCA"
         Height          =   225
         Left            =   360
         TabIndex        =   12
         Top             =   180
         Value           =   -1  'True
         Width           =   3195
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   435
      Left            =   3960
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   8
      Top             =   1800
      Width           =   2835
      Begin VB.CommandButton Command3 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   0
         MouseIcon       =   "Form20.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.TextBox datatual 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "    /    /2005"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox p2 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   5055
   End
   Begin VB.TextBox p1 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   5055
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   1800
      Width           =   2835
      Begin VB.CommandButton Command2 
         Caption         =   "&Ok - salvar"
         Height          =   375
         Left            =   0
         MouseIcon       =   "Form20.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   -360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Iniciar"
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data :"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do projetista :"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do projeto :"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Form1.Command1.Enabled = True
    Form20.Visible = False
End Sub

Private Sub Command2_Click()
    Form20.Visible = False
    Command2.Visible = False
    Form21.Visible = True
    Form21.Command1.Visible = False
    Form21.Command4.Visible = True
    Form21.Command2.Visible = False
End Sub

Private Sub Command3_Click()
    Form20.Visible = False
End Sub
Private Sub datatual_GotFocus()
datatual = Date
End Sub
Private Sub datatual_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub p1_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then p2.SetFocus
End Sub
Private Sub p1_GotFocus()
p1.SelStart = 0
p1.SelLength = Len(p1)
End Sub
Private Sub p2_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44
If KeyAscii = 13 Then datatual.SetFocus
End Sub
Private Sub p2_GotFocus()
p2.SelStart = 0
p2.SelLength = Len(p2)
End Sub

