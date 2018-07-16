VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmGym 
   Caption         =   "Agenda_Personalizada - Tela de Academia"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12015
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   12015
   Begin VB.CommandButton BtnExit 
      Caption         =   "Sair"
      Height          =   495
      Left            =   6600
      TabIndex        =   10
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton BtnSave 
      Caption         =   "Salvar"
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   3750
      TabIndex        =   18
      Top             =   6120
      Width           =   4335
   End
   Begin VB.TextBox Txtname 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      MaxLength       =   60
      TabIndex        =   6
      Top             =   5520
      Width           =   2655
   End
   Begin VB.TextBox Txtid 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      MaxLength       =   7
      TabIndex        =   5
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtcod 
      Height          =   285
      Left            =   1920
      MaxLength       =   7
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtnome 
      Height          =   285
      Left            =   4200
      MaxLength       =   60
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Btnedit 
      Caption         =   "Editar"
      Height          =   495
      Left            =   9240
      TabIndex        =   3
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
   Begin VB.CommandButton BtnNovo 
      Caption         =   "Novo"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   2400
      TabIndex        =   11
      Top             =   5280
      Width           =   7215
      Begin VB.ComboBox CboMusc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5520
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblId 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblExercicio 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblMusculo 
         AutoSize        =   -1  'True
         Caption         =   "Musculo:"
         Height          =   195
         Left            =   4800
         TabIndex        =   12
         Top             =   240
         Width           =   645
      End
   End
   Begin MSFlexGridLib.MSFlexGrid dbugrid 
      Height          =   4215
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7435
      _Version        =   393216
   End
   Begin VB.Label lblcod 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisar por codigo:"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   360
      Width           =   1530
   End
   Begin VB.Label lblnome 
      AutoSize        =   -1  'True
      Caption         =   "Pesquisar por nome:"
      Height          =   195
      Left            =   2760
      TabIndex        =   16
      Top             =   360
      Width           =   1440
   End
End
Attribute VB_Name = "FrmGym"
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
    CboMusc.Enabled = False
    BtnNovo.Enabled = True
    Btnedit.Enabled = True
    BtnDelete.Enabled = True
    BtnSave.Enabled = False
    BtnCancel.Enabled = False
    BtnExit.Enabled = True
    Call fillcombo_Musculo
End Sub

Private Sub BtnDelete_Click()
'Botão Deletar

    If Not dbugrid.Col = 1 Then dbugrid.Col = 1
        If Trim(Val(dbugrid.Text)) <> 0 Then
            Set rs = BD.Execute("SELECT * FROM academia WHERE Pk_Acm = " & CInt(dbugrid.Text) & "")
            While Not rs.EOF
                Txtid.Text = rs!Pk_Acm
                Txtname.Text = rs!Nome
                CboMusc.Text = rs!Musculo
                rs.MoveNext
            Wend
                Txtid.Enabled = False
                Txtname.Enabled = False
                CboMusc.Enabled = False

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
            Set rs = BD.Execute("SELECT * FROM academia WHERE Pk_Acm = " & CInt(dbugrid.Text) & "")
            While Not rs.EOF
                Txtid.Text = rs!Pk_Acm
                Txtname.Text = rs!Nome
                CboMusc.Text = rs!Musculo
                rs.MoveNext
            Wend
                Txtid.Enabled = False
                Txtname.Enabled = True
                CboMusc.Enabled = True
    
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
    CboMusc.Enabled = True
    
    BtnNovo.Enabled = False
    Btnedit.Enabled = False
    BtnDelete.Enabled = False
    BtnSave.Enabled = True
    BtnCancel.Enabled = True
    BtnExit.Enabled = False
    
    Call NewID_Academia
    CRUD = 1
    Txtname.SetFocus
End Sub

Private Sub BtnSave_Click()
'Funcionalidades CRUD
    Select Case CRUD
        Case 1: 'Cadastrar
            If Trim(Txtname.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                Txtname.SetFocus
                Exit Sub
            ElseIf Trim(CboMusc.Text) = "" Then
                MsgBox "O Campo musculo está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                CboMusc.SetFocus
                Exit Sub
            End If
                BD.Execute "Insert into academia values(" & CInt(Txtid.Text) & "','" & Txtname.Text & "','" & CboMusc.Text & "')"
                MsgBox "cadastrado efetuado com sucesso", vbInformation, "Agenda_Personalizada"
                Call BtnCancel_Click
                Call CarregaDados
                
        Case 2: 'Alterar
            If Trim(Txtname.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                Txtname.SetFocus
                Exit Sub
            ElseIf Trim(CboMusc.Text) = "" Then
                MsgBox "O Campo musculo está em branco,favor preencher corretamente.", vbExclamation, "Agenda_Personalizada"
                CboMusc.SetFocus
                Exit Sub
            End If
                BD.Execute "Update academia set nome = '" & Txtname.Text & "',musculo = '" & CboMusc.Text & "' where Pk_Acm = " & CInt(Txtid.Text) & ""
                MsgBox "Atualização efetuada com sucesso", vbInformation, "Agenda_Personalizada"
                Call BtnCancel_Click
                Call CarregaDados
                
        Case 3: 'Deletar
                msg = MsgBox("Deseja confirmar a exclusão do registro?", vbQuestion + vbYesNo, "Agenda_Personalizada")
                If msg = 6 Then
                     BD.Execute "Delete from academia where Pk_Acm = " & CInt(Txtid.Text) & ""
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
        Call fillcombo_Musculo
End Sub

Private Sub FormataGrid()
    'Colunas inicial
    dbugrid.Cols = 4
    'Linha inicial
    dbugrid.Rows = 1
    dbugrid.ColWidth(0) = 500
    dbugrid.Row = 0
    
    'Preenchimento das colunas
    dbugrid.ColWidth(1) = 800
    dbugrid.Col = 1
    dbugrid.Text = "Código"
    
    dbugrid.ColWidth(2) = 2200
    dbugrid.Col = 2
    dbugrid.Text = "Nome"
    
    
    dbugrid.ColWidth(3) = 1500
    dbugrid.Col = 3
    dbugrid.Text = "Musculo"
    
End Sub

Private Sub CarregaDados()
'Atualizador do Flexgrid
    FormataGrid
    strSQL = "SELECT * FROM Academia WHERE 0 = 0"
    
    If Trim(Val(txtcod.Text)) <> 0 Then
        strSQL = strSQL & " AND Pk_Acm >= '" & txtcod.Text & "' order by Pk_Acm"
    ElseIf Trim(txtnome.Text) <> "" Then
        strSQL = strSQL & " AND nome like '%" & txtnome.Text & "%' order by nome"
    Else
        strSQL = strSQL & " order by Pk_Acm"
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
        dbugrid.Text = rs("Pk_Acm")
        dbugrid.Col = 2
        dbugrid.Text = rs("Nome")
        dbugrid.Col = 3
        dbugrid.Text = rs("Musculo")
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
