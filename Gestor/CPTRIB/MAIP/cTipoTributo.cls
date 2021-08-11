VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTipoTributo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Codigo As Integer
Private m_Tipo As String
Private m_contaContabil As String
Private m_Receita As cGeral

Public Property Get Receita() As cGeral
    Set Receita = m_Receita
End Property

Public Property Get contaContabil() As String
    contaContabil = m_contaContabil
End Property

Public Property Get Tipo() As String
    Tipo = m_Tipo
End Property

Public Property Get Codigo() As Integer
    Codigo = m_Codigo
End Property

Public Function buscarProximo() As Integer
    Dim rs As VSRecordset
    Dim Sql As String
    
    Sql = "SELECT MAX(TTT_COD_TIPO)+1 FROM TAB_TIPO_TRIBUTO"
    If Bdados.AbreTabela(Sql, rs) Then
        m_Codigo = Nvl(rs(0), 1)
    Else
        m_Codigo = 1
    End If
    buscarProximo = m_Codigo
    Bdados.FechaTabela rs
End Function

Public Function preencherGrid(ByRef grd As Object) As Boolean
    preencherGrid = grd.Preencher(Bdados, "select TTT_COD_TIPO as Codigo, TTT_GRUPO AS Tipo," & _
                                            " TTT_NUM_CONTA AS Conta, " & _
                                            " TGE_NOME AS Receita " & _
                                        " FROM TAB_TIPO_TRIBUTO,TAB_GERAL " & _
                                        " WHERE TTT_RECEITA_TRIBUTARIA = TGE_CODIGO " & _
                                            " AND TGE_TIPO = (SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME ='" & Me.Receita.Tabela & "') ORDER BY TTT_COD_TIPO")
End Function

Public Sub preencherObjeto(Codigo As Integer, Tipo As String, Conta As String, Receita As String)
    If Codigo = 0 Then
        m_Codigo = Me.buscarProximo()
    Else
        m_Codigo = Codigo
    End If
    m_Tipo = Tipo
    m_contaContabil = Conta
    m_Receita.Descricao = Receita
End Sub

Private Function validarObjeto() As Boolean
    validarObjeto = True
    
    validarObjeto = validarObjeto And (Me.Tipo <> "")
End Function

Public Function Salvar() As Boolean
    Dim Valores As String, Campos As String
    
    If validarObjeto() Then
        Campos = "TTT_COD_TIPO, TTT_GRUPO, TTT_NUM_CONTA, TTT_RECEITA_TRIBUTARIA"
        Valores = Bdados.PreparaValor(Me.Codigo, Me.Tipo, Me.contaContabil, Me.Receita.Codigo)
        Salvar = Bdados.GravaDados("TAB_TIPO_TRIBUTO", Valores, Campos, "TTT_COD_TIPO=" & Me.Codigo)
    Else
        Erro "Preencha as informações antes de salvar."
    End If
End Function

Public Function Excluir() As Boolean
    If validarObjeto() Then
        If Confirma("Excluir " & Me.Tipo & " ?") Then
            Excluir = Bdados.DeletaDados("TAB_TIPO_TRIBUTO", "TTT_COD_TIPO=" & Me.Codigo)
        End If
    End If
End Function
Private Sub Class_Initialize()
    Set m_Receita = New cGeral
    m_Receita.Tabela = "TIPO RECEITA"
End Sub
