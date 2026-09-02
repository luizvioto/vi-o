VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "O que é afilamento?"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00E0E0E0&
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3435
      ScaleWidth      =   8235
      TabIndex        =   19
      Top             =   120
      Width           =   8295
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form3.frx":030A
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   7935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form3.frx":0470
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   8055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form3.frx":07E4
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   8100
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "O Afilamento (     ) - Taper Ratio"
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
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Porque Afilar a Asa?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "l"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1680
         TabIndex        =   20
         Top             =   120
         Width           =   135
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   8235
      TabIndex        =   1
      Top             =   5400
      Width           =   8295
      Begin VB.PictureBox Picture5 
         Height          =   1455
         Left            =   7080
         ScaleHeight     =   1395
         ScaleWidth      =   1035
         TabIndex        =   5
         Top             =   120
         Width           =   1095
         Begin VB.CommandButton Command2 
            Caption         =   "&Sair"
            Height          =   1395
            Left            =   0
            MouseIcon       =   "Form3.frx":08A4
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   0
            Width           =   1035
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1395
         ScaleWidth      =   1035
         TabIndex        =   4
         Top             =   120
         Width           =   1095
         Begin VB.CommandButton Command1 
            Caption         =   "&Enviar e Sair"
            Height          =   1395
            Left            =   0
            MouseIcon       =   "Form3.frx":0BAE
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   0
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   1320
         TabIndex        =   3
         Top             =   960
         Width           =   5655
         Begin VB.TextBox lda 
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
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   360
            Max             =   10
            Min             =   3
            TabIndex        =   8
            Top             =   240
            Value           =   10
            Width           =   2175
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AFILAMENTO :"
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
            Left            =   3000
            TabIndex        =   16
            Top             =   240
            Width           =   1305
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00400000&
         Height          =   855
         Left            =   1320
         ScaleHeight     =   795
         ScaleWidth      =   5595
         TabIndex        =   2
         Top             =   120
         Width           =   5655
         Begin VB.Label bw 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   4200
            TabIndex        =   15
            Top             =   480
            Width           =   45
         End
         Begin VB.Label cp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   2280
            TabIndex        =   14
            Top             =   480
            Width           =   45
         End
         Begin VB.Label cr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Left            =   360
            TabIndex        =   13
            Top             =   480
            Width           =   45
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Envergadura"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   4200
            TabIndex        =   12
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Corda da Ponta"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   2280
            TabIndex        =   11
            Top             =   120
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Corda da Raiz"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   360
            TabIndex        =   10
            Top             =   120
            Width           =   1005
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1515
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   3720
      Width           =   8295
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "Cponta"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   720
         TabIndex        =   18
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "Craiz"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   3840
         TabIndex        =   17
         Top             =   720
         Width           =   345
      End
      Begin VB.Line ln5 
         BorderColor     =   &H0000FFFF&
         X1              =   1320
         X2              =   1320
         Y1              =   360
         Y2              =   1200
      End
      Begin VB.Line ln2 
         BorderColor     =   &H0000FFFF&
         X1              =   6600
         X2              =   6600
         Y1              =   360
         Y2              =   1200
      End
      Begin VB.Line ln7 
         BorderColor     =   &H0000FFFF&
         X1              =   3960
         X2              =   1320
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line ln6 
         BorderColor     =   &H0000FFFF&
         X1              =   3960
         X2              =   1320
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line ln4 
         BorderColor     =   &H0000FFFF&
         X1              =   3960
         X2              =   6600
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line ln3 
         BorderColor     =   &H0000FFFF&
         X1              =   3960
         X2              =   6600
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line ln1 
         BorderColor     =   &H0000FFFF&
         X1              =   3960
         X2              =   3960
         Y1              =   360
         Y2              =   1200
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.afilasa = Form3.lda
Form3.Visible = False
HScroll1.Value = 10
End Sub

Private Sub Command2_Click()

'ASA

ln2.x1 = 6600
ln2.X2 = 6600
ln2.y1 = 360
ln2.y2 = 1200

ln3.x1 = 3960
ln3.X2 = 6600
ln3.y1 = 1200
ln3.y2 = 1200

ln4.x1 = 3960
ln4.X2 = 6600
ln4.y1 = 360
ln4.y2 = 360

ln5.x1 = 1320
ln5.X2 = 1320
ln5.y1 = 360
ln5.y2 = 1200

ln6.x1 = 3960
ln6.X2 = 1320
ln6.y1 = 1200
ln6.y2 = 1200

ln7.x1 = 3960
ln7.X2 = 1320
ln7.y1 = 360
ln7.y2 = 360

ln1.Visible = True
ln2.Visible = True
ln3.Visible = True
ln4.Visible = True
ln5.Visible = True
ln6.Visible = True
ln7.Visible = True

HScroll1.Value = 10
Form3.Visible = False

End Sub

Private Sub Form_Load()
HScroll1.Value = 10
End Sub

Private Sub HScroll1_Change()
'Asa

a = 42

If HScroll1.Value = 10 Then
cr = 1
bw = "10.00 metros"
lda = 1
cp = lda * cr

ln2.y1 = 360
ln2.y2 = 1200

ln3.y1 = 1200
ln3.y2 = 1200

ln4.y1 = 360
ln4.y2 = 360

ln5.y1 = 360
ln5.y2 = 1200

ln6.y1 = 1200
ln6.y2 = 1200

ln7.y1 = 360
ln7.y2 = 360

End If

If HScroll1.Value = 9 Then
cr = 1
bw = "10.00 metros"
lda = 0.9
cp = lda * cr

ln2.y1 = 360 + a
ln2.y2 = 1200 - a

End If

If HScroll1.Value = 8 Then
cr = 1
bw = 100
lda = 0.8
cp = lda * cr

ln2.y1 = 360 + (2 * a)
ln2.y2 = 1200 - (2 * a)

End If

If HScroll1.Value = 7 Then
cr = 1
bw = 100
lda = 0.7
cp = lda * cr

ln2.y1 = 360 + (3 * a)
ln2.y2 = 1200 - (3 * a)

End If

If HScroll1.Value = 6 Then
cr = 1
bw = "10.00 metros"
lda = 0.6
cp = lda * cr

ln2.y1 = 360 + (4 * a)
ln2.y2 = 1200 - (4 * a)

End If

If HScroll1.Value = 5 Then
cr = 1
bw = 100
lda = 0.5
cp = lda * cr

ln2.y1 = 360 + (5 * a)
ln2.y2 = 1200 - (5 * a)

End If

If HScroll1.Value = 4 Then
cr = 1
bw = 100
lda = 0.4
cp = lda * cr

ln2.y1 = 360 + (6 * a)
ln2.y2 = 1200 - (6 * a)

End If

If HScroll1.Value = 3 Then
cr = 1
bw = 100
lda = 0.3
cp = lda * cr

ln2.y1 = 360 + (7 * a)
ln2.y2 = 1200 - (7 * a)

End If


cp = Format(cp, "0.00")
cp = cp + " metros."

cr = Format(cr, "0.00")
cr = cr + " metros."


bw = "10.00 metros"

ln3.y2 = ln2.y2
ln4.y2 = ln2.y1

ln5.y1 = ln2.y1
ln5.y2 = ln2.y2

ln6.y2 = ln3.y2
ln7.y2 = ln4.y2

End Sub
