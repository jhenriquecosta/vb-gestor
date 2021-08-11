VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pLotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Codigo As Integer
Private m_Hierarquia As Integer
Private m_Sigla As String
Private m_Descricao As String

Public Property Get Descricao() As String
    Descricao = m_Descricao
End Property

Public Property Let Descricao(ByVal Value As String)
    m_Descricao = Value
End Property

Public Property Get Sigla() As String
    Sigla = m_Sigla
End Property

Public Property Let Sigla(ByVal Value As String)
    m_Sigla = Value
End Property

Public Property Get Hierarquia() As Integer
    Hierarquia = m_Hierarquia
End Property

Public Property Let Hierarquia(ByVal Value As Integer)
    m_Hierarquia = Value
End Property

Public Property Get Codigo() As Integer
    Codigo = m_Codigo
End Property

Public Property Let Codigo(ByVal Value As Integer)
    m_Codigo = Value
End Property

Public Function Gravar(Codigo As Integer) As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    
    Campos = "TLO_CODIGO, TLO_HIERARQUIA, TLO_SIGLA, TLO_DESCRICAO"
    Valores = BDados.PreparaValor(m_Codigo, m_Hierarquia, m_Sigla, m_Descricao)
    Condicao = "TLO_CODIGO = " & Codigo
    If BDados.GravaDados("TAB_LOTACAO", Valores, Campos, Condicao) Then Gravar = True

End Function

Public Function Buscar(Codigo As Integer, Optional Detalhes As Boolean = True) As Boolean
    Dim Sql As String
    Dim RS As VSRecordset
    
    Sql = "SELECT * FROM TAB_LOTACAO WHERE TLO_CODIGO = " & Codigo
    If BDados.AbreTabela(Sql) Then
        m_Codigo = RS!TLO_CODIGO
        If Detalhes Then
            m_Hierarquia = RS!TLO_HIERARQUIA
            m_Sigla = RS!TLO_SIGLA
            m_Descricao = RS!TLO_DESCRICAO
        End If
        Buscar = True
    End If
End Function


Public Sub ExibirEstrutura(Tree As Object, Optional Icone As Boolean = False)
'--------------------------------------------------------------------------------
' Projeto      :  VSProfile
' Procedimento :  ExibirEstrutura
' Descrição    :  Monta um treeview com a hierarquia montada
' Criado por   :  Éderson Robson Bragagnolo
' Maquina      :  FTTEC02
' Data         :  13/06/2003
' Alteração    :  13/06/2003
' Parâmetros   :  Tree (Object)
'--------------------------------------------------------------------------------

    On Error GoTo ERR
    Dim RS As VSRecordset
    Dim Chave As String
    Dim No As Node
    Dim Sql As String
    
    Tree.Nodes.Clear
    Chave = ""
    Sql = "SELECT * FROM TAB_LOTACAO ORDER BY TLO_HIERARQUIA, TLO_CODIGO"
    
    If BDados.AbreTabela(Sql, RS) Then
        Do While Not RS.EOF
        Chave = "LOTACAO:" & RS!TLO_CODIGO
            If RS!TLO_HIERARQUIA = 0 And RS!TLO_CODIGO = 0 Then
                If Icone Then
                    Set No = Tree.Nodes.Add(, , Chave, RS!TLO_SIGLA & "  - " & RS!TLO_DESCRICAO, "CLOSE")
                Else
                    Set No = Tree.Nodes.Add(, , Chave, RS!TLO_SIGLA & "  - " & RS!TLO_DESCRICAO)
                End If
                No.Expanded = True
            Else
                Dim NohPai As String
                NohPai = "LOTACAO:" & RS!TLO_HIERARQUIA
                If Icone Then
                    Set No = Tree.Nodes.Add(NohPai, tvwChild, Chave, RS!TLO_SIGLA & "  - " & RS!TLO_DESCRICAO, "CLOSE")
                Else
                    Set No = Tree.Nodes.Add(NohPai, tvwChild, Chave, RS!TLO_SIGLA & "  - " & RS!TLO_DESCRICAO)
                End If
            End If
            Tree.Nodes(Chave).Tag = Chave & ":" & RS!TLO_CODIGO & ":" & RS!TLO_HIERARQUIA
            RS.MoveNext
        Loop
    End If
    BDados.FechaTabela RS
    Exit Sub
ERR:
'    If ERR.Number = 35601 Then
'        If Icone Then
'            Set No = Tree.Nodes.Add(, , Chave, RS!TLO_SIGLA & "  - " & RS!TLO_DESCRICAO, "CLOSE")
'        Else
'            Set No = Tree.Nodes.Add(, , Chave, RS!TLO_SIGLA & "  - " & RS!TLO_DESCRICAO)
'        End If
'        Resume Next
'    Else
'        Erro ERR.Description
'        BDados.FechaTabela RS
'        Exit Sub
'    End If
End Sub

Public Function PreencherCombo(Combo As Object) As Boolean
    Dim Sql As String
    
    Sql = "SELECT TLO_SIGLA, TLO_CODIGO FROM TAB_LOTACAO ORDER BY TLO_SIGLA"
    If Combo.Preencher(BDados, Sql) Then PreencherCombo = True
End Function

Public Function BuscaProximaLotacao() As Integer
    On Error GoTo Trata
    Dim Sql As String
    
    Sql = "SELECT MAX(TLO_CODIGO) FROM TAB_LOTACAO"
    If BDados.AbreTabela(Sql) Then
        BuscaProximaLotacao = IIf(IsNull(BDados.Tabela(0)), 0, Nvl("" & BDados.Tabela(0), 0) + 1)
    Else
        BuscaProximaLotacao = -1
    End If
    BDados.FechaTabela
    Exit Function
Trata:
    Util.Erro ERR.Description
End Function


