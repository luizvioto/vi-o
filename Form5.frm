VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form Form5 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Motores"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7005
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame motor 
      Caption         =   "&Motores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   6855
      Left            =   25
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   4560
         ScaleHeight     =   315
         ScaleWidth      =   2115
         TabIndex        =   5
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton Command4 
            Caption         =   "&Modificar escolha"
            Height          =   315
            Left            =   0
            MouseIcon       =   "Form5.frx":030A
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   0
            Width           =   2115
         End
      End
      Begin VB.TextBox motescol 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   5595
         Left            =   180
         ScaleHeight     =   5535
         ScaleWidth      =   6495
         TabIndex        =   1
         Top             =   960
         Width           =   6555
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "Form5.frx":0614
            Height          =   5535
            Left            =   0
            OleObjectBlob   =   "Form5.frx":0624
            TabIndex        =   2
            Top             =   0
            Width           =   6495
         End
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   240
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Motores"
            Top             =   4320
            Visible         =   0   'False
            Width           =   1065
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Motor escolhido :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   1530
      End
   End
   Begin VB.Menu filecadastro5 
      Caption         =   "&Cadastrar "
      Begin VB.Menu filecadmot 
         Caption         =   "Motor"
      End
   End
   Begin VB.Menu filedelete5 
      Caption         =   "&Excluir "
      Begin VB.Menu fileexcmot 
         Caption         =   "Motor"
      End
   End
   Begin VB.Menu filesai5 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
