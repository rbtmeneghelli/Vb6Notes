VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmEmp 
   Caption         =   "Agenda_Personalizada - Tela de Empresas"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12015
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   12015
   Begin VB.TextBox txtsenha 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      MaxLength       =   20
      TabIndex        =   8
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox txtlogin 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6360
      MaxLength       =   40
      TabIndex        =   7
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "Sair"
      Height          =   495
      Left            =   6600
      TabIndex        =   11
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   10
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton BtnSave 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   3750
      TabIndex        =   19
      Top             =   6120
      Width           =   4335
   End
   Begin VB.TextBox Txtname 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      MaxLength       =   60
      TabIndex        =   6
      Top             =   5520
      Width           =   2655
   End
   Begin VB.TextBox Txtid 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   5
      Top             =   5520
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   840
      TabIndex        =   12
      Top             =   5280
      Width           =   10215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         Height          =   195
         Left            =   7560
         TabIndex        =   20
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblMusculo 
         AutoSize        =   -1  'True
         Caption         =   "Login:"
         Height          =   195
         Left            =   5040
         TabIndex        =   15
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblExercicio 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblId 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.CommandButton BtnNovo 
      Caption         =   "Novo"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton BtnDelete 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Btnedit 
      Caption         =   "Editar"
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtnome 
      Height          =   285
      Left            =   4200
      MaxLength       =   60
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox txtcod 
      Height          =   285
      Left            =   1920
      MaxLength       =   7
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid dbugrid 
      Height          =   4215
      Left            =   240
      TabIndex        =   16
      Top             =   840
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7435
      _Version        =   393216
   End
   Begin VB.Label lblnome 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisar por nome:"
      Height          =   195
      Left            =   2760
      TabIndex        =   18
      Top             =   360
      Width           =   1440
   End
   Begin VB.Label lblcod 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisar por codigo:"
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   360
      Width           =   1530
   End
End
Attribute VB_Name = "FrmEmp"
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
    txtlogin.Enabled = False
    txtsenha.Enabled = False
    BtnNovo.Enabled = True
    Btnedit.Enabled = True
    BtnDelete.Enabled = True
    BtnSave.Enabled = False
    BtnCancel.Enabled = False
    BtnExit.Enabled = True
End Sub

Private Sub BtnDelete_Click()
'Bot�o Deletar

    If Not dbugrid.Col = 1 Then dbugrid.Col = 1
        If Trim(Val(dbugrid.Text)) <> 0 Then
            Set rs = BD.Execute("SELECT * FROM EMPRESAS WHERE PK_Emp = " & CInt(dbugrid.Text) & "")
            While Not rs.EOF
                Txtid.Text = rs!PK_Emp
                Txtname.Text = rs!Nome
                txtlogin.Text = rs!Login
                txtsenha.Text = rs!Senha
                rs.MoveNext
            Wend
                Txtid.Enabled = False
                Txtname.Enabled = False
                txtlogin.Enabled = False
                txtsenha.Enabled = False

                BtnNovo.Enabled = False
                Btnedit.Enabled = False
                BtnDelete.Enabled = False
                BtnSave.Enabled = True
                BtnCancel.Enabled = True
                BtnExit.Enabled = False
    
                CRUD = 3
        Else
             MsgBox "O registro solicitado n�o foi localizado", vbCritical, "Agenda_Personalizada"
        End If
End Sub

Private Sub Btnedit_Click()
'Bot�o Editar
    
    If Not dbugrid.Col = 1 Then dbugrid.Col = 1
        If Trim(Val(dbugrid.Text)) <> 0 Then
            Set rs = BD.Execute("SELECT * FROM empresas WHERE Pk_emp = " & CInt(dbugrid.Text) & "")
            While Not rs.EOF
                Txtid.Text = rs!PK_Emp
                Txtname.Text = rs!Nome
                txtlogin.Text = rs!Login
                txtsenha.Text = rs!Senha
                rs.MoveNext
            Wend
                Txtid.Enabled = False
                Txtname.Enabled = True
                txtlogin.Enabled = True
                txtsenha.Enabled = True
    
                BtnNovo.Enabled = False
                Btnedit.Enabled = False
                BtnDelete.Enabled = False
                BtnSave.Enabled = True
                BtnCancel.Enabled = True
                BtnExit.Enabled = False
    
                CRUD = 2
                Txtname.SetFocus
        Else
             MsgBox "O registro solicitado n�o foi localizado", vbCritical, "Agenda_Personalizada"
        End If
End Sub

Private Sub BtnExit_Click()
'Botao Sair
    'Desbloqueia as op��es de menu e sai da tela
        Limpa_Campos Me
        Frmmenu.Usuario.Enabled = True
        Frmmenu.Contatos.Enabled = True
        Frmmenu.Empresa.Enabled = True
        Frmmenu.Academia.Enabled = True
        Frmmenu.Sair.Enabled = True
        Unload Me
End Sub

Private Sub BtnNovo_Click()
'Bot�o Cadastrar
    Txtid.Enabled = False
    Txtname.Enabled = True
    txtlogin.Enabled = True
    txtsenha.Enabled = True
    
    BtnNovo.Enabled = False
    Btnedit.Enabled = False
    BtnDelete.Enabled = False
    BtnSave.Enabled = True
    BtnCancel.Enabled = True
    BtnExit.Enabled = False
    
    Call NewID_Empresa
    CRUD = 1
    Txtname.SetFocus
End Sub

Private Sub BtnSave_Click()
'Funcionalidades CRUD
    Select Case CRUD
        Case 1: 'Cadastrar
            If Trim(Txtname.Text) = "" Then
                MsgBox "O Campo nome est� em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                Txtname.SetFocus
                Exit Sub
            ElseIf Trim(txtlogin.Text) = "" Then
                MsgBox "O Campo login est� em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                txtlogin.SetFocus
                Exit Sub
            ElseIf Trim(txtsenha.Text) = "" Then
                MsgBox "O Campo senha est� em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                txtsenha.SetFocus
                Exit Sub
            End If
                BD.Execute "Insert into empresas values(" & CInt(Txtid.Text) & ",'" & Txtname.Text & "','" & txtlogin.Text & "','" & txtsenha.Text & "')"
                MsgBox "cadastrado efetuado com sucesso", vbInformation, "Agenda_Personalizada"
                Call BtnCancel_Click
                Call CarregaDados
                
        Case 2: 'Alterar
            If Trim(Txtname.Text) = "" Then
                MsgBox "O Campo nome est� em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                Txtname.SetFocus
                Exit Sub
            ElseIf Trim(txtlogin.Text) = "" Then
                MsgBox "O Campo login est� em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                txttel.SetFocus
                Exit Sub
            ElseIf Trim(txtsenha.Text) = "" Then
                MsgBox "O Campo senha est� em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                txtcel.SetFocus
                Exit Sub
            End If
                BD.Execute "Update empresas set nome = '" & Txtname.Text & "',login = '" & txtlogin.Text & "',senha = '" & txtsenha.Text & "' where Pk_Emp = " & CInt(Txtid.Text) & ""
                MsgBox "Atualiza��o efetuada com sucesso", vbInformation, "Agenda_Personalizada"
                Call BtnCancel_Click
                Call CarregaDados
                
        Case 3: 'Deletar
                msg = MsgBox("Deseja confirmar a exclus�o do registro?", vbQuestion + vbYesNo, "Agenda_Personalizada")
                If msg = 6 Then
                     BD.Execute "Delete from empresas where Pk_Emp = " & CInt(Txtid.Text) & ""
                     MsgBox "Exclus�o efetuada com sucesso", vbInformation, "Agenda_Personalizada"
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
End Sub

Private Sub FormataGrid()
    'Colunas inicial
    dbugrid.Cols = 5
    'Linha inicial
    dbugrid.Rows = 1
    dbugrid.ColWidth(0) = 500
    dbugrid.Row = 0
    
    'Preenchimento das colunas
    dbugrid.ColWidth(1) = 800
    dbugrid.Col = 1
    dbugrid.Text = "C�digo"
    
    dbugrid.ColWidth(2) = 2000
    dbugrid.Col = 2
    dbugrid.Text = "Nome"
    
    
    dbugrid.ColWidth(3) = 2700
    dbugrid.Col = 3
    dbugrid.Text = "Login"
    
    dbugrid.ColWidth(4) = 1500
    dbugrid.Col = 4
    dbugrid.Text = "Senha"
    
End Sub

Private Sub CarregaDados()
'Atualizador do Flexgrid
    FormataGrid
    strSQL = "SELECT * FROM empresas WHERE 0 = 0"
    
    If Trim(Val(txtcod.Text)) <> 0 Then
        strSQL = strSQL & " AND PK_emp >= '" & txtcod.Text & "' order by Pk_Cont"
    ElseIf Trim(txtnome.Text) <> "" Then
        strSQL = strSQL & " AND nome like '%" & txtnome.Text & "%' order by nome"
    Else
        strSQL = strSQL & " order by Pk_Emp"
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
        dbugrid.Text = rs("Pk_Emp")
        dbugrid.Col = 2
        dbugrid.Text = rs("Nome")
        dbugrid.Col = 3
        dbugrid.Text = rs("Login")
        dbugrid.Col = 4
        dbugrid.Text = rs("Senha")
        rs.MoveNext
    Loop
End Sub

Private Sub txtcod_Change()
CarregaDados
End Sub

Private Sub txtnome_Change()
CarregaDados
End Sub

Private Sub Txtname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
    If KeyAscii = 8 Then Exit Sub
    If IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub
