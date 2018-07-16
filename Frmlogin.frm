VERSION 5.00
Begin VB.Form Frmlogin 
   Caption         =   "Agenda_Personalizada - Tela de Login"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4815
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtsenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtlog 
      Height          =   285
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "Sair"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton BtnLog 
      Caption         =   "Logar"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frmlog 
      Height          =   1815
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   4335
      Begin VB.Label lblse 
         AutoSize        =   -1  'True
         Caption         =   "Senha do usuario:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label lbllog 
         AutoSize        =   -1  'True
         Caption         =   "Login do usuario:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msg As Integer

Private Sub BtnExit_Click()
'Botao Sair
    msg = MsgBox("Deseja sair do sistema?", vbQuestion + vbYesNo, "Progfake")
        If msg = 6 Then
            BD.Close
            Unload Me
        End If
End Sub

Private Sub BtnCancel_Click()
'Botão Cancelar
    Limpa_Campos Me
    txtlog.SetFocus
End Sub

Private Sub BtnLog_Click()
'Botão Logar
    If Trim(txtlog.Text) = "" Then
        MsgBox "O Campo login do usuario está em branco,favor preenche-lo.", vbExclamation, "Agenda_Personalizada"
        txtlog.SetFocus
        Exit Sub
    ElseIf Trim(txtsenha.Text) = "" Then
        MsgBox "O Campo senha do usuario está em branco,favor preenche-lo.", vbExclamation, "Agenda_Personalizada"
        txtsenha.SetFocus
        Exit Sub
    End If

    Set rs = BD.Execute("SELECT * FROM USERS WHERE login = '" & txtlog.Text & "' and senha = '" & txtsenha.Text & "'")
    
    If rs.EOF Then
        MsgBox "Usuario ou senha invalida", vbExclamation, "Agenda_Personalizada"
        BtnCancel_Click
        txtlog.SetFocus
    Else
        Pk_User = rs!Pk_User
        User_Login = rs!Login
        Call BtnCancel_Click
        MsgBox "Usuario " & txtlog.Text & " logado com sucesso", vbInformation + vbOKOnly, "Agenda_Personalizada"
        Unload Me
        Frmmenu.Show
    End If
End Sub

Private Sub Form_Load()
'Chama a função de conexão com o banco de dados
Call AbrebancoSQL
'Call AbrebancoACCESS
End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BtnLog_Click
End Sub
