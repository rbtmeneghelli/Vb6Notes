VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Progfake"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcel 
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1320
      Width           =   4695
   End
   Begin VB.TextBox txtobs 
      Enabled         =   0   'False
      Height          =   855
      Left            =   840
      MaxLength       =   500
      TabIndex        =   4
      Top             =   2280
      Width           =   4695
   End
   Begin VB.ComboBox Cbogen 
      Enabled         =   0   'False
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   4695
   End
   Begin VB.TextBox txtnome 
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      MaxLength       =   80
      TabIndex        =   1
      Top             =   840
      Width           =   4695
   End
   Begin VB.ComboBox Cbocod 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "Sair"
      Height          =   615
      Left            =   3720
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2280
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton BtnSearch 
      Caption         =   "Pesquisar"
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton BtnSave 
      Caption         =   "Confirmar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton BtnDelete 
      Caption         =   "Apagar"
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton BtnEdit 
      Caption         =   "Editar"
      Height          =   615
      Left            =   1560
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton BtnInsert 
      Caption         =   "Novo"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   5535
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celular:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Genero:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label3 
         Caption         =   "Obs:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   540
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BtnSave_Click()
'Funcionalidades CRUD
    Select Case CRUD
        Case 1: 'Cadastrar
            If Trim(txtnome.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Progfake"
                txtnome.SetFocus
                Exit Sub
            ElseIf Trim(txtcel.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Progfake"
                txtcel.SetFocus
                Exit Sub
            ElseIf Trim(Cbogen.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Progfake"
                Cbogen.SetFocus
                Exit Sub
            ElseIf Trim(txtobs.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Progfake"
                txtobs.SetFocus
                Exit Sub
            End If
                BD.Execute "Insert into Users values('" & Cbocod.Text & "','" & txtnome.Text & "','" & txtcel.Text & "','" & Cbogen.Text & "','" & txtobs.Text & "')"
                MsgBox "cadastrado efetuado com sucesso", vbInformation, "Progfake"
                Call BtnCancel_Click
                
        Case 2: 'Alterar
            If Trim(txtnome.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Progfake"
                txtnome.SetFocus
                Exit Sub
            ElseIf Trim(txtcel.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Progfake"
                txtcel.SetFocus
                Exit Sub
            ElseIf Trim(Cbogen.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Progfake"
                Cbogen.SetFocus
                Exit Sub
            ElseIf Trim(txtobs.Text) = "" Then
                MsgBox "O Campo nome está em branco,favor preencher corretamente.", vbExclamation, "Progfake"
                txtobs.SetFocus
                Exit Sub
            End If
                BD.Execute "Update Users set nome = '" & txtnome.Text & "',celular = '" & txtcel.Text & "',genero = '" & Cbogen.Text & "',obs = '" & txtobs.Text & "' where codigo = '" & Cbocod.Text & "'"
                MsgBox "Atualização efetuada com sucesso", vbInformation, "Progfake"
                Call BtnCancel_Click
                
        Case 3: 'Deletar
                msg = MsgBox("Deseja confirmar a exclusão do registro?", vbQuestion + vbYesNo, "Progfake")
                If msg = 6 Then
                     BD.Execute "Delete from Users where codigo = '" & Cbocod.Text & "'"
                     MsgBox "Exclusão efetuada com sucesso", vbInformation, "Progfake"
                     Call BtnCancel_Click
                End If
    End Select
End Sub





