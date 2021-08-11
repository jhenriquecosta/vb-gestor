VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Codigo As String
Private m_Nome As String
Private m_Senha As String
Private m_Ativo As Integer
Private m_Lotacao As Integer
Public Matricula As String

Public Property Get Lotacao() As Integer
    Lotacao = m_Lotacao
End Property

Public Property Let Lotacao(ByVal Value As Integer)
    m_Lotacao = Value
End Property

Public Property Get Ativo() As Integer
    Ativo = m_Ativo
End Property

Public Property Let Ativo(ByVal Value As Integer)
    m_Ativo = Value
End Property

Public Property Get Senha() As String
    Senha = m_Senha
End Property

Public Property Let Senha(ByVal Value As String)
    m_Senha = Value
End Property

Public Property Get Nome() As String
    Nome = m_Nome
End Property

Public Property Let Nome(ByVal Value As String)
    m_Nome = Value
End Property

Public Property Get Codigo() As String
    Codigo = m_Codigo
End Property

Public Property Let Codigo(ByVal Value As String)
    m_Codigo = Value
End Property

Public Sub ExibirUsuarios(Tree As Object)
'--------------------------------------------------------------------------------
' Projeto      :  VSProfile
' Procedimento :  ExibirUsuarios
' Descrição    :  Monta um treeview com os usuários cadastrados
' Criado por   :  Éderson Robson Bragagnolo
' Maquina      :  FTTEC02
' Data         :  16/06/2003
' Alteração    :  16/06/2003
' Parâmetros   :  Tree (Object)
'--------------------------------------------------------------------------------

    On Error GoTo ERR
    Dim Rs As VSRecordset
    Dim Chave As String
    Dim No As Node
    Dim Sql As String
    Dim NohPai As String
        
    Chave = ""
    NohPai = ""
    Sql = "SELECT * FROM TAB_USUARIO ORDER BY TUS_TLO_CODIGO"
    
    If BDados.AbreTabela(Sql, Rs) Then
        Do While Not Rs.EOF
            Chave = "USUARIO:" & Rs!TUS_COD_USUARIO
            NohPai = "LOTACAO:" & Rs!TUS_TLO_CODIGO

            Set No = Tree.Nodes.Add(NohPai, tvwChild, Chave, Rs!TUS_NOME, "PESSOA")
            'TAG DO NÓ INDICADO: CHAVE + CODIGO DA LOTACAO + PERFIL DE SEGURANÇA
            Tree.Nodes(Chave).Tag = Chave & ":" & Rs!TUS_TLO_CODIGO
            Rs.MoveNext
        Loop
    End If
    BDados.FechaTabela Rs
    Exit Sub
ERR:
    If ERR.Number <> 0 Then
        Erro ERR.Description
        BDados.FechaTabela Rs
        Exit Sub
    End If
End Sub

Public Function Gravar(Codigo As String) As Boolean

    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    
    If Not Buscar(Codigo, False) Then
        m_Senha = Seguranca.Criptografa(Temp.PegaParametro(BDados, "SENHA INICIAL"))
    Else
        m_Senha = Seguranca.Criptografa(m_Senha)
    End If
    m_Ativo = 1
    
    Campos = "TUS_COD_USUARIO, TUS_NOME, TUS_SENHA, TUS_ATIVO, TUS_TLO_CODIGO,TUS_TSE_MATRICULA"
    Valores = BDados.PreparaValor(m_Codigo, BDados.Converte(m_Nome, tctexto), m_Senha, m_Ativo, m_Lotacao, BDados.Converte(Matricula, tctexto))
    Condicao = "TUS_COD_USUARIO = '" & Codigo & "'"
    If BDados.GravaDados("TAB_USUARIO", Valores, Campos, Condicao) Then Gravar = True

End Function

Public Function Buscar(Codigo As String, Optional Detalhes As Boolean = True) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "SELECT * FROM TAB_USUARIO WHERE TUS_COD_USUARIO = '" & Codigo & "'"
    
    If BDados.AbreTabela(Sql, Rs) Then
        m_Codigo = Rs!TUS_COD_USUARIO
        m_Senha = Seguranca.DesCriptografa(Rs!TUS_SENHA)
        If Detalhes Then
            m_Nome = Rs!TUS_NOME
            m_Senha = Seguranca.DesCriptografa(Rs!TUS_SENHA)
            m_Ativo = Rs!TUS_ATIVO
            Matricula = "" & Rs!TUS_TSE_MATRICULA
'            m_Lotacao = RS!TUS_TLO_CODIGO
        End If
        Buscar = True
    End If
End Function

Public Function Excluir(Codigo As String) As Boolean
    If Codigo <> Aplicacoes.Usuario Then
        If Util.Confirma("Excluir o usuário " & Codigo & " do controle de acesso?") Then
            ' EXCLUI TODOS OS ACESSOS DO USUÁRIO
            If BDados.DeletaDados("TAB_ACESSO_USUARIO", "TAU_TUS_COD_USUARIO = '" & Codigo & "'") Then
                'EXCLUI O USUÁRIO
                If BDados.DeletaDados("TAB_USUARIO", "TUS_COD_USUARIO = '" & Codigo & "'") Then Excluir = True
            End If
        End If
    Else
        Util.Mensagem "O Usuário atual somente pode ser excluído por outro usuário do sistema."
    End If
End Function
