VERSION 5.00
Begin VB.MDIForm Frmmenu 
   BackColor       =   &H8000000C&
   Caption         =   "Agenda_Personalizada - Tela Principal"
   ClientHeight    =   8790
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14550
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Usuario 
      Caption         =   "&Usuario"
      Begin VB.Menu AlterarSenha 
         Caption         =   "&Alterar Senha"
      End
   End
   Begin VB.Menu Contatos 
      Caption         =   "&Contatos"
      Begin VB.Menu Telefones 
         Caption         =   "&Telefones"
      End
   End
   Begin VB.Menu Empresa 
      Caption         =   "&Empresa"
      Begin VB.Menu Sites 
         Caption         =   "&Sites"
      End
   End
   Begin VB.Menu Academia 
      Caption         =   "&Academia"
      Begin VB.Menu Exercicios 
         Caption         =   "&Exercicios"
      End
   End
   Begin VB.Menu Sair 
      Caption         =   "&Sair"
      Begin VB.Menu Logoff 
         Caption         =   "&Logoff"
      End
   End
End
Attribute VB_Name = "Frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AlterarSenha_Click()
    'Carrega a tela de Troca de senha
        Frmupdate.Show
    'Bloqueia as opções de menu
        Usuario.Enabled = False
        Contatos.Enabled = False
        Empresa.Enabled = False
        Academia.Enabled = False
        Sair.Enabled = False
End Sub

Private Sub Exercicios_Click()
    'Carrega a tela de Troca de senha
        FrmGym.Show
    'Bloqueia as opções de menu
        Usuario.Enabled = False
        Contatos.Enabled = False
        Empresa.Enabled = False
        Academia.Enabled = False
        Sair.Enabled = False
End Sub

Private Sub Logoff_Click()
    'Botao Sair
        msg = MsgBox("Deseja fazer logoff do sistema?", vbQuestion + vbYesNo, "Agenda_Personalizada")
        If msg = 6 Then
            Unload Me
        End If
End Sub

Private Sub Sites_Click()
    'Carrega a tela de Troca de senha
        FrmEmp.Show
    'Bloqueia as opções de menu
        Usuario.Enabled = False
        Contatos.Enabled = False
        Empresa.Enabled = False
        Academia.Enabled = False
        Sair.Enabled = False
End Sub

Private Sub Telefones_Click()
    'Carrega a tela de Troca de senha
        Form2.Show
    'Bloqueia as opções de menu
        Usuario.Enabled = False
        Contatos.Enabled = False
        Empresa.Enabled = False
        Academia.Enabled = False
        Sair.Enabled = False
End Sub
