VERSION 5.00
Begin VB.Form Form30 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperação de parafuso"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8535
   ForeColor       =   &H00E0E0E0&
   Icon            =   "Form30.frx":0000
   LinkTopic       =   "Form30"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4095
      Begin VB.TextBox sombra 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form30.frx":030A
         ForeColor       =   &H00404040&
         Height          =   795
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3795
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form30.frx":03C9
         ForeColor       =   &H00404040&
         Height          =   1215
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form30.frx":04EF
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   3855
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
   End
   Begin VB.PictureBox Picture2 
      Height          =   415
      Left            =   4320
      ScaleHeight     =   360
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   3000
      Width           =   975
      Begin VB.CommandButton Command1 
         Caption         =   "&Calcular "
         Height          =   355
         Left            =   0
         MouseIcon       =   "Form30.frx":0602
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   2655
      Left            =   4440
      Picture         =   "Form30.frx":090C
      ScaleHeight     =   2595
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Menu filesai30 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "Form30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Form30.sombra = Empty Then
        MsgBox " Entre com a área sombreada", vbInformation
        Exit Sub
    End If

If Form30.sombra <> Empty Then

If Form2.svt = Empty Or Form10.MARGEM = Empty Then
MsgBox "O projeto deve ser totalmente calculado.", vbCritical
Exit Sub
End If

fatorP = Abs(Form2.svt) - Abs(Form30.sombra)

fracaoP1 = 0.25 * Form2.svt
fracaoP2 = 0.3 * Form2.svt
fracaoP3 = 0.35 * Form2.svt
fracaoP4 = 0.4 * Form2.svt

    If Form10.MARGEM <= 5 And fatorP <= fracaoP1 Then
        Text2 = "O avião dificilmente sairá do parafuso"
    End If
    
    If Form10.MARGEM > 5 And fatorP <= fracaoP1 Then
        Text2 = "Péssimas condições de recuperação"
    End If

        If Form10.MARGEM > 5 Then
            If fatorP > fracaoP1 And fatorP <= fracaoP2 Then
                Text2 = "Péssimas condições de recuperação"
            End If

                If fatorP > fracaoP2 And fatorP <= fracaoP3 Then
                    Text2 = "Razoaveis condições de recuperação"
                End If

                    If fatorP > fracaoP3 And fatorP <= fracaoP4 Then
                        Text2 = "Boas condições de recuperação"
                    End If

                        If fatorP > fracaoP4 Then
                            Text2 = "Excelentes condições de recuperação"
                        End If
        End If
End If
End Sub

Private Sub filesai30_Click()
Form30.Visible = False
Text2 = Empty
End Sub
