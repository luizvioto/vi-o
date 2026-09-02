VERSION 5.00
Begin VB.Form Form22 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "O que é alongamento?"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form22.frx":0000
   LinkTopic       =   "Form22"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame DEFINE 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alongamento - ( Aspect Ratio )  AR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3495
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   8300
      Begin VB.PictureBox Picture6 
         Height          =   375
         Left            =   4920
         ScaleHeight     =   315
         ScaleWidth      =   3075
         TabIndex        =   22
         Top             =   3000
         Width           =   3135
         Begin VB.CommandButton Command3 
            Caption         =   "&Vantagens e desvantagens"
            Height          =   315
            Left            =   0
            MouseIcon       =   "Form22.frx":030A
            MousePointer    =   99  'Custom
            TabIndex        =   23
            Top             =   0
            Width           =   3075
         End
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form22.frx":0614
         ForeColor       =   &H00800000&
         Height          =   3135
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   8055
      End
   End
   Begin VB.Frame VANTAGEM 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vantagens e Desvantagens"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3495
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   8295
      Begin VB.PictureBox Picture7 
         Height          =   375
         Left            =   4920
         ScaleHeight     =   315
         ScaleWidth      =   3075
         TabIndex        =   24
         Top             =   3000
         Width           =   3135
         Begin VB.CommandButton Command4 
            Caption         =   "&Voltar"
            Height          =   315
            Left            =   0
            MouseIcon       =   "Form22.frx":0D45
            MousePointer    =   99  'Custom
            TabIndex        =   25
            Top             =   0
            Width           =   3075
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   75
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form22.frx":104F
         ForeColor       =   &H00800000&
         Height          =   2055
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   7935
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1515
      ScaleWidth      =   8235
      TabIndex        =   14
      Top             =   3720
      Width           =   8295
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   855
         Left            =   1920
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   5400
      Width           =   8295
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00400000&
         Height          =   855
         Left            =   1320
         ScaleHeight     =   795
         ScaleWidth      =   5595
         TabIndex        =   9
         Top             =   120
         Width           =   5655
         Begin VB.Label Sw 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   480
            Width           =   45
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Área da asa "
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Corda média"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   2280
            TabIndex        =   13
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Envergadura"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   4440
            TabIndex        =   12
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Cm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   2280
            TabIndex        =   11
            Top             =   480
            Width           =   45
         End
         Begin VB.Label bw 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   4440
            TabIndex        =   10
            Top             =   480
            Width           =   45
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   5655
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   360
            Max             =   12
            Min             =   5
            TabIndex        =   7
            Top             =   240
            Value           =   5
            Width           =   2175
         End
         Begin VB.TextBox AR 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4440
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ALONGAMENTO :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2760
            TabIndex        =   8
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1395
         ScaleWidth      =   1035
         TabIndex        =   3
         Top             =   120
         Width           =   1095
         Begin VB.CommandButton Command1 
            Caption         =   "&Enviar e sair"
            Height          =   1395
            Left            =   0
            MouseIcon       =   "Form22.frx":1481
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   0
            Width           =   1035
         End
      End
      Begin VB.PictureBox Picture5 
         Height          =   1455
         Left            =   7080
         ScaleHeight     =   1395
         ScaleWidth      =   1035
         TabIndex        =   1
         Top             =   120
         Width           =   1095
         Begin VB.CommandButton Command2 
            Caption         =   "&Sair"
            Height          =   1395
            Left            =   0
            MouseIcon       =   "Form22.frx":178B
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   0
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.arw = Form22.AR
Form22.Visible = False
HScroll1.Value = 5
End Sub

Private Sub Command2_Click()
Form22.Visible = False
HScroll1.Value = 5
End Sub

Private Sub Command3_Click()
DEFINE.Visible = False
VANTAGEM.Visible = True
End Sub

Private Sub Command4_Click()
DEFINE.Visible = True
VANTAGEM.Visible = False
End Sub


Private Sub HScroll1_Change()

a = 154

If HScroll1.Value = 5 Then
Shape1.Width = 4215 * 1
Shape1.Height = 855
End If

If HScroll1.Value = 6 Then
Shape1.Width = 4215 * 1.13
Shape1.Height = 0.88 * 855
Shape1.Left = 1920 - a
End If

If HScroll1.Value = 7 Then
Shape1.Width = 4215 * 1.23
Shape1.Height = 0.815 * 855
Shape1.Left = 1920 - (a * 2)
End If

If HScroll1.Value = 8 Then
Shape1.Width = 4215 * 1.3
Shape1.Height = 0.765 * 855
Shape1.Left = 1920 - (a * 3)
End If

If HScroll1.Value = 9 Then
Shape1.Width = 4215 * 1.38
Shape1.Height = 0.724 * 855
Shape1.Left = 1920 - (a * 4)
End If

If HScroll1.Value = 10 Then
Shape1.Width = 4215 * 1.46
Shape1.Height = 0.688 * 855
Shape1.Left = 1920 - (a * 5)
End If

If HScroll1.Value = 11 Then
Shape1.Width = 4215 * 1.54
Shape1.Height = 0.65 * 855
Shape1.Left = 1920 - (a * 6)
End If

If HScroll1.Value = 12 Then
Shape1.Width = 4215 * 1.56
Shape1.Height = 0.626 * 855
Shape1.Left = 1920 - (a * 7.5)
End If

AR = HScroll1.Value
Sw = "10.00 m2"

bw1 = (10 * HScroll1.Value) ^ 0.5
bw1 = Format(bw1, "0.00")
bw = bw1 + " metros"

Cm1 = 10 / bw1
Cm1 = Format(Cm1, "0.00")
Cm = Cm1 + " metros"
End Sub


