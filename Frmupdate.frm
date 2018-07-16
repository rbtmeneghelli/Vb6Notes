VERSION 5.00
Begin VB.Form Frmupdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agenda_Personalizada - Tela de troca de senha"
   ClientHeight    =   2055
   ClientLeft      =   2490
   ClientTop       =   -135
   ClientWidth     =   4815
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4815
   Begin VB.CommandButton BtnClose 
      Caption         =   "Fechar"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton BtnChange 
      Caption         =   "Alterar Senha"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtsenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtlog 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Frame Frmupdate 
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   4335
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nova senha:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Login do usuario:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Frmupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnChange_Click()
'Botão de Alterar Senha
    If Trim(txtlog.Text) = "" Then
        MsgBox "O Campo login do usuario está em branco,favor preenche-lo.", vbExclamation, "Agenda_Personalizada"
        Exit Sub
    ElseIf Trim(txtsenha.Text) = "" Then
        MsgBox "O Campo novo senha está em branco,favor preenche-lo.", vbExclamation, "Agenda_Personalizada"
        txtsenha.SetFocus
        Exit Sub
    End If
        BD.Execute ("UPDATE USERS set senha = '" & txtsenha.Text & "' where login = '" & txtlog.Text & "'")
        txtsenha.Text = ""
        MsgBox "Senha alterada com sucesso", vbInformation + vbOKOnly, "Agenda_Personalizada"
End Sub

Private Sub BtnClose_Click()
'Botão Fechar
    'Desbloqueia as opções de menu e sai da tela
        txtsenha.Text = ""
        Frmmenu.Usuario.Enabled = True
        Frmmenu.Contatos.Enabled = True
        Frmmenu.Empresa.Enabled = True
        Frmmenu.Academia.Enabled = True
        Frmmenu.Sair.Enabled = True
        Unload Me
End Sub

Private Sub Form_Load()
    'Configura o Formulario de forma correta
        Me.Top = (Frmmenu.Height / 3) - (Me.Top / 2)
        Me.Left = (Frmmenu.Width / 3) - (Me.Left / 2)
        
        txtlog.Text = User_Login 'Recebe o Login da tela de login para ser possivel a alteração de senha
End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BtnChange_Click
End Sub

