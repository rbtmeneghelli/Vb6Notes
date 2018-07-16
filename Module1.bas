Attribute VB_Name = "Module1"
'References 'Microsoft ActiveX
Global CRUD, Pk_User As Integer
Global rs As New ADODB.Recordset
Global User_Login  As String
Global BD As New ADODB.Connection

Function AbrebancoACCESS()
'Faz a conexão do VB com o Access
    If Dir(App.Path & "\Agenda.accdb", vbArchive) = "" Then
        MsgBox "Houve um erro ao tentar abrir o banco de dados. O programa será encerrado.", vbCritical
        End
    Else
        'BD.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & App.Path & "\Agenda.mdb;Exclusive=1;Uid=admin;Pwd=;"
        BD.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\Agenda.accdb;Persist Security Info=False;"
        BD.Open
    End If
End Function

Function AbrebancoSQL()
'Faz a conexão do VB com o SQL Server
    On Error GoTo Trata_Erro
        BD.ConnectionString = "Provider=SQLNCLI10;Server=MENEGHELLI\SQLEXPRESS;Database=agenda;Trusted_Connection=yes;"
        BD.Open
        Exit Function
Trata_Erro:
        MsgBox "Houve um erro ao tentar abrir o banco de dados. O programa será encerrado.", vbCritical, "ProgFake"
        End
End Function

Function fillcombo_Social()
'Função para preencher automaticamente o combo de Redesocial
Form2.cboreso.Clear
Set rs = BD.Execute("SELECT * FROM social order by descricao")
    While Not rs.EOF
        Form2.cboreso.AddItem (rs!Descricao)
        rs.MoveNext
    Wend
End Function

Function fillcombo_Musculo()
'Função para preencher automaticamente o combo de Musculo
FrmGym.CboMusc.Clear
Set rs = BD.Execute("SELECT * FROM Musculo order by descricao")
    While Not rs.EOF
        FrmGym.CboMusc.AddItem (rs!Descricao)
        rs.MoveNext
    Wend
End Function

Public Sub Limpa_Campos(Formulario As Form)
    Dim cont As Integer
    For cont = 0 To Formulario.Count - 1
        'Limpa as caixas de texto
        If TypeOf Formulario.Controls(cont) Is TextBox Then
            If Trim(Formulario.Controls(cont).Text) <> Empty Then
                Formulario.Controls(cont).Text = Empty
            End If
        End If
        'Limpa as caixas de combo
        If TypeOf Formulario.Controls(cont) Is ComboBox Then
            If Trim(Formulario.Controls(cont).Text) <> 0 Then
                Formulario.Controls(cont).Clear
            End If
        End If
    Next
End Sub

Public Sub formatTelefone(vObjControl As Control)
    If TypeOf vObjControl Is TextBox Then
        If Len(vObjControl.Text) = 1 Then vObjControl.Text = "(" & vObjControl.Text
        If Len(vObjControl.Text) = 3 Then vObjControl.Text = vObjControl.Text & ") "
        If Len(vObjControl.Text) = 9 Then vObjControl.Text = vObjControl.Text & "-"
    End If
End Sub

Public Sub formatCelular(vObjControl As Control)
    If TypeOf vObjControl Is TextBox Then
        If Len(vObjControl.Text) = 1 Then vObjControl.Text = "(" & vObjControl.Text
        If Len(vObjControl.Text) = 3 Then vObjControl.Text = vObjControl.Text & ") "
        If Len(vObjControl.Text) = 10 Then vObjControl.Text = vObjControl.Text & "-" 'Len passa a ser 10 ao inves de 9
    End If
End Sub

Function NewID_Contato()
'Cria um novo codigo para inserir na tabela Contatos
    Dim codigo As Integer
    Set rs = BD.Execute("SELECT max(Pk_Cont) as Pk_Cont FROM Contatos")
    codigo = CInt(rs!Pk_Cont)
    If codigo > 0 And Trim(codigo) <> "" Then
        Form2.Txtid.Text = codigo + 1
    Else
        Form2.Txtid.Text = CInt(1)
    End If
End Function

Function NewID_Academia()
'Cria um novo codigo para inserir na tabela Academia
    Dim codigo As Integer
    Set rs = BD.Execute("SELECT max(Pk_Acm) as Pk_Acm FROM Academia")
    codigo = CInt(rs!Pk_Acm)
    If codigo > 0 And Trim(codigo) <> "" Then
        FrmGym.Txtid.Text = codigo + 1
    Else
        FrmGym.Txtid.Text = CInt(1)
    End If
End Function

Function NewID_Empresa()
'Cria um novo codigo para inserir na tabela
    Dim codigo As Integer
    Set rs = BD.Execute("SELECT max(Pk_Emp) as Pk_Emp FROM Empresas")
    codigo = CInt(rs!Pk_Emp)
    If codigo > 0 And Trim(codigo) <> "" Then
       FrmEmp.Txtid.Text = codigo + 1
    Else
        FrmEmp.Txtid.Text = CInt(1)
    End If
End Function
