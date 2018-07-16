VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Agenda_Personalizada - Tela de Contatos"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   12015
   Begin VB.CommandButton BtnExit 
      Caption         =   "Sair"
      Height          =   495
      Left            =   6600
      TabIndex        =   22
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   21
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton BtnSave 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   20
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   3840
      TabIndex        =   19
      Top             =   6000
      Width           =   4125
   End
   Begin VB.TextBox txtcel 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7800
      MaxLength       =   15
      TabIndex        =   8
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txttel 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      MaxLength       =   14
      TabIndex        =   7
      Top             =   5400
      Width           =   1335
   End
   Begin VB.ComboBox cboreso 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10200
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txtname 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      MaxLength       =   60
      TabIndex        =   6
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox txtid 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   5400
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   5160
      Width           =   11535
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Rede social:"
         Height          =   195
         Left            =   9000
         TabIndex        =   18
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celular:"
         Height          =   195
         Left            =   6960
         TabIndex        =   17
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.CommandButton BtnNovo 
      Caption         =   "Novo"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnDelete 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Btnedit 
      Caption         =   "Editar"
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtnome 
      Height          =   285
      Left            =   4200
      MaxLength       =   60
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox txtcod 
      Height          =   285
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid dbugrid 
      Height          =   4215
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7435
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisar por nome:"
      Height          =   195
      Left            =   2760
      TabIndex        =   12
      Top             =   240
      Width           =   1440
   End
   Begin VB.Label lblcod 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisar por codigo:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1530
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String
Dim iCont As Integer

Private Sub BtnCancel_Click()
'Botao Cancelar
    Limpa_Campos Me
    Txtid.Enabled = False
    Txtname.Enabled = False
    txtcel.Enabled = False
    txttel.Enabled = False
    cboreso.Enabled = False
    BtnNovo.Enabled = True
    Btnedit.Enabled = True
    BtnDelete.Enabled = True
    BtnSave.Enabled = False
    BtnCancel.Enabled = False
    BtnExit.Enabled = True
    Call fillcombo_Social
End Sub

Private Sub BtnDelete_Click()
'Botão Deletar

    If Not dbugrid.Col = 1 Then dbugrid.Col = 1
        If Trim(Val(dbugrid.Text)) <> 0 Then
            Set rs = BD.Execute("SELECT * FROM CONTATOS WHERE PK_Cont = " & CInt(dbugrid.Text) & "")
            While Not rs.EOF
                Txtid.Text = rs!PK_Cont
                Txtname.Text = rs!Nome
                txttel.Text = rs!Telefone
                txtcel.Text = rs!Celular
                cboreso.Text = rs!RedeSocial
                rs.MoveNext
            Wend
                Txtid.Enabled = False
                Txtname.Enabled = False
                txtcel.Enabled = False
                txttel.Enabled = False
                cboreso.Enabled = False
    
                BtnNovo.Enabled = False
                Btnedit.Enabled = False
                BtnDelete.Enabled = False
                BtnSave.Enabled = True
                BtnCancel.Enabled = True
                BtnExit.Enabled = False
    
                CRUD = 3
        Else
             MsgBox "O registro solicitado não foi localizado", vbCritical, "Agenda_Personalizada"
        End If
End Sub

Private Sub Btnedit_Click()
'Botão Editar
    
    If Not dbugrid.Col = 1 Then dbugrid.Col = 1
        If Trim(Val(dbugrid.Text)) <> 0 Then
            Set rs = BD.Execute("SELECT * FROM CONTATOS WHERE PK_Cont = " & CInt(dbugrid.Text) & "")
            While Not rs.EOF
                Txtid.Text = rs!PK_Cont
                Txtname.Text = rs!Nome
                txttel.Text = rs!Telefone
                txtcel.Text = rs!Celular
                cboreso.Text = rs!RedeSocial
                rs.MoveNext
            Wend
                Txtid.Enabled = False
                Txtname.Enabled = True
                txtcel.Enabled = True
                txttel.Enabled = True
                cboreso.Enabled = True
    
                BtnNovo.Enabled = False
                Btnedit.Enabled = False
                BtnDelete.Enabled = False
                BtnSave.Enabled = True
                BtnCancel.Enabled = True
                BtnExit.Enabled = False
    
                CRUD = 2
                Txtname.SetFocus
        Else
             MsgBox "O registro solicitado não foi localizado", vbCritical, "Agenda_Personalizada"
        End If
End Sub

Private Sub BtnExit_Click()
'Botao Sair
    'Desbloqueia as opções de menu e sai da tela
        Limpa_Campos Me
        Frmmenu.Usuario.Enabled = True
        Frmmenu.Contatos.Enabled = True
        Frmmenu.Empresa.Enabled = True
        Frmmenu.Academia.Enabled = True
        Frmmenu.Sair.Enabled = True
        Unload Me
End Sub

Private Sub BtnNovo_Click()
'Botão Cadastrar
    Txtid.Enabled = False
    Txtname.Enabled = True
    txtcel.Enabled = True
    txttel.Enabled = True
    cboreso.Enabled = True
    
    BtnNovo.Enabled = False
    Btnedit.Enabled = False
    BtnDelete.Enabled = False
    BtnSave.Enabled = True
    BtnCancel.Enabled = True
    BtnExit.Enabled = False
    
    Call NewID_Contato
    CRUD = 1
    txtnome.SetFocus
End Sub

Private Sub BtnSave_Click()
'Funcionalidades CRUD
    Select Case CRUD
        Case 1: 'Cadastrar
            If Trim(Txtname.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                Txtname.SetFocus
                Exit Sub
            ElseIf Trim(txttel.Text) = "" Then
                MsgBox "O Campo telefone está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                txttel.SetFocus
                Exit Sub
            ElseIf Trim(txtcel.Text) = "" Then
                MsgBox "O Campo celular está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                txtcel.SetFocus
                Exit Sub
            ElseIf Trim(cboreso.Text) = "" Then
                MsgBox "O Campo Redesocial está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                cboreso.SetFocus
                Exit Sub
            End If
                BD.Execute "Insert into Contatos values(" & CInt(Txtid.Text) & ",'" & Txtname.Text & "','" & txttel.Text & "','" & txtcel.Text & "','" & cboreso.Text & "'," & CInt(Pk_User) & ")"
                MsgBox "cadastrado efetuado com sucesso", vbInformation, "Agenda_Personalizada"
                Call BtnCancel_Click
                Call CarregaDados
                
        Case 2: 'Alterar
            If Trim(Txtname.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                Txtname.SetFocus
                Exit Sub
            ElseIf Trim(txttel.Text) = "" Then
                MsgBox "O Campo telefone está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                txttel.SetFocus
                Exit Sub
            ElseIf Trim(txtcel.Text) = "" Then
                MsgBox "O Campo celular está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                txtcel.SetFocus
                Exit Sub
            ElseIf Trim(cboreso.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                cboreso.SetFocus
                Exit Sub
            End If
                BD.Execute "Update Contatos set nome = '" & Txtname.Text & "',celular = '" & txtcel.Text & "',telefone = '" & txttel.Text & "',redesocial = '" & cboreso.Text & "' where Pk_cont = " & CInt(Txtid.Text) & ""
                MsgBox "Atualização efetuada com sucesso", vbInformation, "Agenda_Personalizada"
                Call BtnCancel_Click
                Call CarregaDados
                
        Case 3: 'Deletar
                msg = MsgBox("Deseja confirmar a exclusão do registro?", vbQuestion + vbYesNo, "Agenda_Personalizada")
                If msg = 6 Then
                     BD.Execute "Delete from contatos where Pk_cont = " & CInt(Txtid.Text) & ""
                     MsgBox "Exclusão efetuada com sucesso", vbInformation, "Agenda_Personalizada"
                     Call BtnCancel_Click
                     Call CarregaDados
                End If
    End Select
End Sub

Private Sub Form_Load()
'Configura o Formulario de forma correta
        Me.Height = 7695
        Me.Width = 12255
        
        FormataGrid
        CarregaDados
        Call fillcombo_Social
End Sub

Private Sub FormataGrid()
    'Colunas inicial
    dbugrid.Cols = 6
    'Linha inicial
    dbugrid.Rows = 1
    dbugrid.ColWidth(0) = 500
    dbugrid.Row = 0
    
    'Preenchimento das colunas
    dbugrid.ColWidth(1) = 800
    dbugrid.Col = 1
    dbugrid.Text = "Código"
    
    dbugrid.ColWidth(2) = 2000
    dbugrid.Col = 2
    dbugrid.Text = "Nome"
    
    
    dbugrid.ColWidth(3) = 1500
    dbugrid.Col = 3
    dbugrid.Text = "Telefone"
    
    dbugrid.ColWidth(4) = 1500
    dbugrid.Col = 4
    dbugrid.Text = "Celular"
    
    dbugrid.ColWidth(5) = 1500
    dbugrid.Col = 5
    dbugrid.Text = "Rede Social"
    
End Sub

Private Sub CarregaDados()
'Atualizador do Flexgrid
    FormataGrid
    strSQL = "SELECT * FROM Contatos WHERE 0 = 0"
    
    If Trim(Val(txtcod.Text)) <> 0 Then
        strSQL = strSQL & " AND Pk_Cont >= '" & txtcod.Text & "' AND Fk_User = '" & Pk_User & "' order by Pk_Cont"
    ElseIf Trim(txtnome.Text) <> "" Then
        strSQL = strSQL & " AND nome like '%" & txtnome.Text & "%' AND Fk_User = '" & Pk_User & "' order by nome"
    Else
        strSQL = strSQL & " AND Fk_User = " & CInt(Pk_User) & " order by Pk_Cont"
    End If
    
    Set rs = BD.Execute(strSQL)
    iCont = 0
    
    Do While Not rs.EOF
        iCont = iCont + 1
        dbugrid.Col = 0
        dbugrid.Rows = dbugrid.Rows + 1
        dbugrid.Row = iCont
        dbugrid.Text = iCont
        dbugrid.Col = 1
        dbugrid.Text = rs("Pk_Cont")
        dbugrid.Col = 2
        dbugrid.Text = rs("nome")
        dbugrid.Col = 3
        dbugrid.Text = rs("telefone")
        dbugrid.Col = 4
        dbugrid.Text = rs("celular")
        dbugrid.Col = 5
        dbugrid.Text = rs("Redesocial")
        rs.MoveNext
    Loop
End Sub

Private Sub txtcod_Change()
CarregaDados
End Sub

Private Sub txtnome_Change()
CarregaDados
End Sub

Private Sub txtcel_KeyUp(KeyCode As Integer, Shift As Integer)
    formatCelular txtcel
    txtcel.SelStart = Len(txtcel.Text)
End Sub

Private Sub Txtname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
    If KeyAscii = 8 Then Exit Sub
    If IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub Txtcel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub txttel_KeyUp(KeyCode As Integer, Shift As Integer)
    formatTelefone txttel
    txttel.SelStart = Len(txttel.Text)
End Sub

Private Sub Txttel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
    If KeyAscii = 8 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub
