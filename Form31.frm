VERSION 5.00
Begin VB.Form Form31 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form31.frx":0000
   LinkTopic       =   "Form31"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   3240
         ScaleHeight     =   315
         ScaleWidth      =   915
         TabIndex        =   6
         Top             =   2640
         Width           =   975
         Begin VB.CommandButton Command1 
            Caption         =   "&Ok!"
            Height          =   315
            Left            =   0
            MouseIcon       =   "Form31.frx":030A
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1935
         Left            =   3600
         Max             =   2250
         SmallChange     =   50
         TabIndex        =   5
         Top             =   600
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Max             =   3000
         SmallChange     =   50
         TabIndex        =   4
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox altura 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox largura 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   2640
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   2250
         Left            =   240
         ScaleHeight     =   2190
         ScaleWidth      =   2940
         TabIndex        =   1
         Top             =   240
         Width           =   3000
         Begin VB.Line Line1 
            BorderColor     =   &H00FF00FF&
            X1              =   120
            X2              =   120
            Y1              =   0
            Y2              =   2280
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF00FF&
            X1              =   0
            X2              =   3000
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Shape desenho 
            BackColor       =   &H0080FFFF&
            BorderColor     =   &H00C00000&
            BorderStyle     =   3  'Dot
            FillColor       =   &H00C00000&
            FillStyle       =   5  'Downward Diagonal
            Height          =   1350
            Left            =   180
            Top             =   180
            Width           =   1800
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   2175
            Left            =   30
            Top             =   15
            Width           =   2895
         End
      End
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If LIBER <> "KFAR0852" Then
    MsgBox "A rotina de impressão sómente estará disponível após a liberação permanente", 16, "Cancelamento de impressão": Exit Sub
End If

Dim alt, larg As Double

Dim resp18 As Double
If Form18.CommonDialog1.filename = "" Then
resp18 = MsgBox("Para ser impresso o desenho deverá ser salvo", vbYesNo)
        If resp18 = vbYes Then
            If OpenFile <> "" Then
                SavePicture Image, OpenFile
            End If
        End If


         If resp18 = vbNo Then
            Exit Sub
         End If
End If

alt = Form31.altura
larg = Form31.largura

Printer.Orientation = 2
Printer.PaintPicture LoadPicture(Form18.CommonDialog1.filename), alt, larg, 12000, 9000 'Largura(12000) e altura(9000) do FORM18
Printer.EndDoc

Form31.Visible = False
End Sub

Private Sub Form_Load()
' Dividir o value max de X Y por 4 para saber abaixo
Form31.HScroll1.Value = 600
Form31.VScroll1.Value = 300

End Sub

Private Sub HScroll1_Change()
largura = Form31.HScroll1 * 4
Line1.x1 = largura / 4
Line1.X2 = largura / 4
desenho.Left = Line1.x1 + 60
End Sub

Private Sub VScroll1_Change()
altura = Form31.VScroll1.Value * 4
Line2.y1 = altura / 4
Line2.y2 = altura / 4
desenho.Top = Line2.y1 + 60
End Sub
