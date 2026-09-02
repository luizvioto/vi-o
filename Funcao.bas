Attribute VB_Name = "Funcao"
Public Sair As Integer
Public LIBER As String
Public VarDisc As String
Public TPsof As Integer
Global DB As Database
Global TB As Recordset
Global WR As Workspace
Global SR As String

Private Function AbreBanc()
On Error GoTo CriaTab
Inicia:
    Set WR = DBEngine.Workspaces(0)
    Set DB = OpenDatabase(App.Path + "\Alvaro.mdb", True, False, ";PWD=AF33216028")
    Exit Function
    Set TB = DB.OpenRecordset("Select * from Protege")
    If TB.RecordCount > 0 Then
        If Dir("C:\D3dbf.dll") = "" Then
            MsgBox "Contate o seu fornecedor!", 16, "Erro de sistema!": Form32.Show 1
            End
            Exit Function
        Else
            Open "C:\D3dbf.dll" For Input As #1
                VarDisc = Trim(Input$(LOF(1), #1))
            Close #1
        End If
        If Trim(Mid(VarDisc, 1, 9)) <> Trim(TB("hd")) Then MsgBox "Cópia não autorizada!", 16, "Erro de acesso!": End: Exit Function
    Else
        MsgBox "Contate o seu fornecedor!", 16, "Erro de sistema!": End: Exit Function
    End If
    Exit Function
CriaTab:
    Select Case Err
        Case 3078
            DB.Execute "Create table Protege" _
                & "(hd text(9))"
            DB.Execute "Insert into protege" _
            & "(hd) values" _
            & "('" & Trim(VarDisc) & "');"
            
            'Abrindo arquivo Dll
            Open "C:\D3dbf.dll" For Output As #1
                Print #1, Trim(VarDisc)
            Close #1
    End Select
    GoTo Inicia
End Function

Public Function Acesso()
GoTo Continuar
    If Dir("C:\D3dbf.dll") = "" Then
        GoTo Abertura
    Else
        Open "C:\D3dbf.dll" For Input As #1
            VarDisc = Trim(Input$(LOF(1), #1))
        Close #1
        GoTo Final
   End If
Abertura:
   Open "C:\Hab.bat" For Output As #1
        Print #1, "Echo off"
        Print #1, "Cls"
        Print #1, "C:"
        Print #1, "Vol > C:\D"
        Print #1, "Cls"
   Close #1
   Carlos = Shell("C:\Hab.bat", vbHide)
   Tempo.PB1.Max = 3000
   Tempo.Show
   For i = 0 To 2999
    Tempo.PB1.Value = i
   Next i
Continuar:
   On Error Resume Next
   Dim LocalSave As String
   Dim CopiArq As String
   
   LocalSave = App.Path + "\ME-01"
   CopiArq = LocalSave + "\ME-01.bmp"
    If Dir(LocalSave) = "" Then
        MkDir LocalSave
        FileCopy App.Path + "\ME-01.bmp", CopiArq
    End If
    Unload Tempo
    Open "C:\D" For Input As #1
         Numdisco = Input$(LOF(1), #1)
    Close #1
    For i = Len(Numdisco) To 0 Step -1
        If Mid(Numdisco, i, 1) = "A" Then Exit For
        If Mid(Numdisco, i, 1) = "B" Then Exit For
        If Mid(Numdisco, i, 1) = "C" Then Exit For
        If Mid(Numdisco, i, 1) = "D" Then Exit For
        If Mid(Numdisco, i, 1) = "E" Then Exit For
        If Mid(Numdisco, i, 1) = "F" Then Exit For
        If Mid(Numdisco, i, 1) = "0" Then Exit For
        If Mid(Numdisco, i, 1) = "1" Then Exit For
        If Mid(Numdisco, i, 1) = "2" Then Exit For
        If Mid(Numdisco, i, 1) = "3" Then Exit For
        If Mid(Numdisco, i, 1) = "4" Then Exit For
        If Mid(Numdisco, i, 1) = "5" Then Exit For
        If Mid(Numdisco, i, 1) = "6" Then Exit For
        If Mid(Numdisco, i, 1) = "7" Then Exit For
        If Mid(Numdisco, i, 1) = "8" Then Exit For
        If Mid(Numdisco, i, 1) = "9" Then Exit For
   Next i
   Inidisc = i - 9
   VarDisc = Trim(Mid(Numdisco, Inidisc, 10))

   Kill ("C:\D")
   Kill ("C:\Hab.bat")
Final:
   Call AbreBanc
End Function

'PARA RODAR FORMs DIRETO SEM CLICAR BOTÕES

'Private Sub Form_Activate()
'Call fileaerodin_Click
'End Sub
