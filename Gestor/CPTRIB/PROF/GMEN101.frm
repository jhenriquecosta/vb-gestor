VERSION 5.00
Begin VB.Form GMEN101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GMEN101"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   360
      Left            =   7185
      TabIndex        =   4
      Top             =   1695
      Width           =   1740
   End
   Begin VB.CommandButton CmdCriarPastas 
      Caption         =   "Criar Configurações"
      Height          =   345
      Left            =   4920
      TabIndex        =   3
      Top             =   1710
      Width           =   2235
   End
   Begin VB.TextBox txtCaminhoPastasUsuarios 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   465
      TabIndex        =   1
      Top             =   1095
      Width           =   8520
   End
   Begin VB.Shape Shape2 
      Height          =   555
      Left            =   465
      Top             =   1605
      Width           =   8535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caminho onde serão criadas as configurações dos usuários"
      Height          =   195
      Left            =   465
      TabIndex        =   2
      Top             =   870
      Width           =   6990
   End
   Begin VB.Shape Shape1 
      Height          =   1620
      Left            =   240
      Top             =   675
      Width           =   8865
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAINEL DE CONTROLE DO GESTOR MENSSEGER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   9165
   End
   Begin VB.Image Image1 
      Height          =   3030
      Left            =   -120
      Picture         =   "GMEN101.frx":0000
      Stretch         =   -1  'True
      Top             =   -75
      Width           =   9585
   End
End
Attribute VB_Name = "GMEN101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApagarPastas_Click()
    
End Sub

Private Sub CmdCriarPastas_Click()
    'On Error Resume Next
    Dim Sql As String
    Dim Rs  As VSRecordset
    Dim Arquivo As Integer
    
    If txtCaminhoPastasUsuarios = "" Then Avisa "Informe o caminho para configurações."
    'PROCEDIMENTO QUE CRIA AS PASTAS DE CONFIGURÇÕES DOS USUÁRIOS
    '1º Crio o diretorio padrão...
    If BDados.AbreTabela("Select * from TAB_CONFIGURACAO_MENSSEGER") Then
        If Not IsNull(BDados.Tabela(0)) Then
            If BDados.Tabela(0) <> "" Then
                BDados.DeletaDados "TAB_CONFIGURACAO_MENSSEGER"
                  Dim Arquivos As New FileSystemObject
                  Dim Pasta           As Object
                  Dim Parametro       As Object
                  Dim Caminho         As String
                  
                    Caminho = txtCaminhoPastasUsuarios
                    Set Pasta = Arquivos.GetFolder(Caminho)
                    Pasta.Delete
                    For Each Parametro In Pasta.Folders
                        Parametro.Delete
                    Next
            End If
        End If
    End If
    If BDados.InsereDados("TAB_CONFIGURACAO_MENSSEGER", BDados.PreparaValor(BDados.Converte(txtCaminhoPastasUsuarios, tctexto)), "TCM_CAMINHO") Then
        MkDir txtCaminhoPastasUsuarios
        'PARA CADA USUÁRIO CADASTRADO NO BANCO,
        'CRIO UMA PASTA COM O NOME DO USUÁRIO PARA GUARDAR SUAS CONFIGURAÇÕES PESSAIS..
        Sql = "Select * from tab_usuario"
        If BDados.AbreTabela(Sql, Rs) Then
            Rs.MoveFirst
            Do Until Rs.EOF
                MkDir txtCaminhoPastasUsuarios & "\" & UCase(Rs.Fields("tus_cod_usuario"))
                'PARA CADA PASTA CRIADA GERO OS ARQUIVOS DE MENSAGEM E CFG...
                Arquivo = FreeFile
                Open txtCaminhoPastasUsuarios & "\" & UCase(Rs.Fields("tus_cod_usuario")) & "\MENSAGEM.TXT" For Output As Arquivo
                Close Arquivo
                Arquivo = FreeFile
                Open txtCaminhoPastasUsuarios & "\" & UCase(Rs.Fields("tus_cod_usuario")) & "\CFG.TXT" For Output As Arquivo
                Close Arquivo
                Rs.MoveNext
            Loop
        End If
    End If
    Avisa "Operação concluída com sucesso."
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If BDados.AbreTabela("Select * from TAB_CONFIGURACAO_MENSSEGER") Then
        txtCaminhoPastasUsuarios = BDados.Tabela(0)
    End If
End Sub

