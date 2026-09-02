VERSION 5.00
Begin VB.Form Form26 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "O que é CL ?"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form26.frx":0000
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00400000&
      Height          =   2175
      Left            =   4320
      ScaleHeight     =   2115
      ScaleWidth      =   4155
      TabIndex        =   9
      Top             =   3840
      Width           =   4215
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         X1              =   1320
         X2              =   1320
         Y1              =   2040
         Y2              =   120
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   840
         Shape           =   3  'Circle
         Top             =   1800
         Width           =   135
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FF8080&
         X1              =   2160
         X2              =   2520
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF8080&
         X1              =   840
         X2              =   2160
         Y1              =   1920
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FFFF&
         X1              =   360
         X2              =   3840
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   8235
      TabIndex        =   2
      Top             =   6120
      Width           =   8295
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   11
         TabIndex        =   5
         Top             =   600
         Width           =   5895
      End
      Begin VB.PictureBox Picture6 
         Height          =   735
         Left            =   6240
         ScaleHeight     =   675
         ScaleWidth      =   1755
         TabIndex        =   4
         Top             =   120
         Width           =   1815
         Begin VB.CommandButton Command2 
            Caption         =   "&Sair"
            Height          =   680
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   1755
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00400000&
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   5835
         TabIndex        =   3
         Top             =   120
         Width           =   5895
         Begin VB.Label stolou 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STALL !!!!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   4800
            TabIndex        =   21
            Top             =   0
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label cl 
            BackStyle       =   0  'Transparent
            Caption         =   "0,20"
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   3840
            TabIndex        =   20
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cl (admensional) ="
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   2400
            TabIndex        =   19
            Top             =   0
            Width           =   1290
         End
         Begin VB.Label alfa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-6,00"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   1440
            TabIndex        =   18
            Top             =   0
            Width           =   360
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alfa (graus) ="
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   360
            TabIndex        =   17
            Top             =   0
            Width           =   930
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0E0FF&
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2115
      ScaleWidth      =   3915
      TabIndex        =   1
      Top             =   3840
      Width           =   3975
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2040
         Shape           =   3  'Circle
         Top             =   240
         Width           =   135
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00FF0000&
         X1              =   2160
         X2              =   2280
         Y1              =   360
         Y2              =   480
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00FF0000&
         X1              =   2280
         X2              =   2400
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00008080&
         X1              =   1440
         X2              =   1800
         Y1              =   1080
         Y2              =   1680
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   " Cl1 = Clmáximo "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Width           =   1245
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   840
         Shape           =   3  'Circle
         Top             =   1560
         Width           =   135
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FF0000&
         X1              =   720
         X2              =   2040
         Y1              =   1920
         Y2              =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "  Slope = inclinação da curva "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   180
         Left            =   1560
         TabIndex        =   16
         Top             =   1200
         Width           =   2160
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cl0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   600
         TabIndex        =   15
         Top             =   1440
         Width           =   270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cl1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a0"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   13
         Top             =   1800
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a1"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2040
         TabIndex        =   12
         Top             =   1800
         Width           =   210
      End
      Begin VB.Line Line19 
         BorderStyle     =   3  'Dot
         X1              =   2160
         X2              =   1080
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line18 
         BorderStyle     =   3  'Dot
         X1              =   2100
         X2              =   2100
         Y1              =   360
         Y2              =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cl"
         Height          =   195
         Left            =   960
         TabIndex        =   11
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   10
         Top             =   1560
         Width           =   120
      End
      Begin VB.Line Line11 
         X1              =   360
         X2              =   3360
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line10 
         X1              =   1200
         X2              =   1200
         Y1              =   120
         Y2              =   2040
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3555
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   120
      Width           =   8300
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form26.frx":030A
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   8055
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form26.frx":04F3
         ForeColor       =   &H00800000&
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   8175
      End
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form26.Visible = False
End Sub

Private Sub Form_Load()
HScroll1.Value = 0
End Sub

Private Sub HScroll1_Change()
If HScroll1.Value = 0 Then
Shape3.Left = 840
Shape3.Top = 1800
alfa = -6#
alfa = Format(alfa, "0.00")
cl = -0.2
cl = Format(cl, "0.00")
End If

If HScroll1.Value = 1 Then
Shape3.Left = 960
Shape3.Top = 1620
alfa = -4#
alfa = Format(alfa, "0.00")
cl = 0#
cl = Format(cl, "0.00")
End If

If HScroll1.Value = 2 Then
Shape3.Left = 1080
Shape3.Top = 1480
alfa = -2#
alfa = Format(alfa, "0.00")
cl = 0.2
cl = Format(cl, "0.00")
End If

If HScroll1.Value = 3 Then
Shape3.Left = 1250
Shape3.Top = 1320
alfa = 0#
alfa = Format(alfa, "0.00")
cl = 0.4
cl = Format(cl, "0.00")
End If

If HScroll1.Value = 4 Then
Shape3.Left = 1440
Shape3.Top = 1080
alfa = 2#
alfa = Format(alfa, "0.00")
cl = 0.6
cl = Format(cl, "0.00")
End If

If HScroll1.Value = 5 Then
Shape3.Left = 1560
Shape3.Top = 960
alfa = 4#
alfa = Format(alfa, "0.00")
cl = 0.8
cl = Format(cl, "0.00")
End If

If HScroll1.Value = 6 Then
Shape3.Left = 1680
Shape3.Top = 840
alfa = 6#
alfa = Format(alfa, "0.00")
cl = 1#
cl = Format(cl, "0.00")
End If

If HScroll1.Value = 7 Then
Shape3.Left = 1760
Shape3.Top = 720
alfa = 8#
alfa = Format(alfa, "0.00")
cl = 1.2
cl = Format(cl, "0.00")
End If

If HScroll1.Value = 8 Then
Shape3.Left = 1880
Shape3.Top = 600
alfa = 10
alfa = Format(alfa, "0.00")
cl = 1.4
cl = Format(cl, "0.00")
End If

If HScroll1.Value = 9 Then
Shape3.Left = 1980
Shape3.Top = 480
alfa = 12#
alfa = Format(alfa, "0.00")
cl = 1.5
cl = Format(cl, "0.00")
End If

If HScroll1.Value = 10 Then
Shape3.Left = 2100
Shape3.Top = 360
alfa = 14#
alfa = Format(alfa, "0.00")
cl = 1.6
cl = Format(cl, "0.00")
End If

If HScroll1.Value = 11 Then
stolou.Visible = True
Form26.Refresh
Shape3.Left = 2280
Shape3.Top = 480
Shape3.Top = i
For i = 480 To 2500 Step 0.025
Shape3.Top = i
Next
alfa = 16#
alfa = Format(alfa, "0.00")
cl = 1.4
cl = Format(cl, "0.00")
Shape3.Top = 1800
Shape3.Left = 840
HScroll1.Value = 0
stolou.Visible = False
End If

End Sub

