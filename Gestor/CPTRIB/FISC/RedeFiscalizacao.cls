VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RedeFiscalizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public rCodEtapa As Double
Public rDescricao As String
Public rCodEtapaPai As Double
Public rCodEtapaOrigem As Double
Public rOrdem As Integer
Public rCodFuncionario As Integer
Public rCaminhoRpt As String
Public rTipoEtapa As Integer
Public rPrazo As Integer
Public rDataCadastro As String
Public rUsuario As String
Public rCodParametroFundamento As Double
Public Enum TipoEtapaRede
    etrProcesso = 1
    etrFase = 2
    etrProcedimento = 3
End Enum
Public ParametrosTexto As New Parametros
Public Sub CarregaDadosRede(CodEtapa As String)
    Dim Sql As String
    Dim Rs As VSRecordset
    LimpaDadosRede
    Sql = "SELECT * FROM TAB_FISCALIZACAO_REDE WHERE TFR_COD_ETAPA =" & CodEtapa
    If Bdados.AbreTabela(Sql, Rs) Then
        rCodEtapa = Nvl("" & Rs!TFR_COD_ETAPA, 0)
        rCaminhoRpt = "" & Rs!TFR_CAMINHO_RPT
        rCodEtapaPai = Nvl("" & Rs!TFR_COD_ETAPA_PAI, 0)
        rCodEtapaOrigem = Nvl("" & Rs!TFR_COD_ETAPA_ORIGEM, 0)
        rCodFuncionario = Nvl("" & Rs!TFR_TFU_COD_FUNCIONARIO, 0)
        rDataCadastro = "" & Rs!TFR_DATA_CADASTRO
        rDescricao = "" & Rs!TFR_DESCRICAO
        rOrdem = Nvl("" & Rs!TFR_PRE_REQUISITO, 0)
        rPrazo = Nvl("" & Rs!TFR_PRAZO, 0)
        rTipoEtapa = Nvl("" & Rs!TFR_TIPO_ETAPA, 0)
        rUsuario = "" & Rs!TFR_TUS_COD_USUARIO
        rCodParametroFundamento = Nvl("" & Rs!TFR_TFT_CODIGO_FUNDAMENTOS, 0)
    End If
End Sub


Public Function CriaEtapaRede(COD_ETAPA As String, Descricao As String, _
            COD_ETAPA_PAI As Double, COD_ETAPA_ORIGEM As Double, Ordem As Integer, COD_FUNCIONARIO_RESPONSAVEL As Double, _
            CAMINHO_RPT As String, TIPO_ETAPA As Integer, Prazo As Integer, Optional CodParametroFundamanto As Double = 0) As Boolean
    Dim Campos As String
    Dim Valores As String
    Campos = "TFR_COD_ETAPA,TFR_DESCRICAO,TFR_COD_ETAPA_PAI,TFR_COD_ETAPA_ORIGEM,TFR_PRE_REQUISITO," & _
        "TFR_TFU_COD_FUNCIONARIO,TFR_CAMINHO_RPT,TFR_TIPO_ETAPA,TFR_PRAZO," & _
        "TFR_DATA_CADASTRO,TFR_TUS_COD_USUARIO,TFR_TFT_CODIGO_FUNDAMENTOS"
        
    Valores = Bdados.PreparaValor(COD_ETAPA, Descricao, COD_ETAPA_PAI, COD_ETAPA_ORIGEM, Ordem, _
            COD_FUNCIONARIO_RESPONSAVEL, CAMINHO_RPT, TIPO_ETAPA, Prazo, _
            Bdados.Converte(Date, TCDataHora), AplicacoesVTFuncoes.Usuario, CodParametroFundamanto)
    CriaEtapaRede = Bdados.GravaDados("TAB_FISCALIZACAO_REDE", Valores, Campos, "TFR_COD_ETAPA=" & COD_ETAPA)
    LimpaDadosRede
End Function

Public Function ApagaEtapaRede(COD_ETAPA As String) As Boolean
    ApagaEtapaRede = Bdados.DeletaDados("TAB_FISCALIZACAO_REDE", "TFR_COD_ETAPA=" & COD_ETAPA)
    LimpaDadosRede
End Function
Public Sub PreencheEtapasPossiveisRede(Grid As Object, CodFiscalizacao As String)
    Dim Fisc As New Fiscalizacao
    Grid.ListItems.Clear
    If Fisc.CarregaDadosFiscalizacao(CodFiscalizacao) Then
        If Fisc.vStatus = 1 Then
        Grid.Preencher Bdados, "Select TFR_COD_ETAPA, TFR_DESCRICAO as Procedimento" & _
                " FROM TAB_FISCALIZACAO_REDE WHERE " & _
                " (TFR_TIPO_ETAPA > 2 AND TFR_COD_ETAPA_PAI =" & _
                "(SELECT TFR_COD_ETAPA_PAI FROM TAB_FISCALIZACAO_REDE WHERE TFR_COD_ETAPA =" & Fisc.vProcedimentoAtual & ") AND (TFR_PRE_REQUISITO = (SELECT TFR_PRE_REQUISITO FROM TAB_FISCALIZACAO_REDE WHERE TFR_COD_ETAPA =" & Fisc.vProcedimentoAtual & ") OR TFR_PRE_REQUISITO = 0)) OR " & _
                "(TFR_PRE_REQUISITO = 0 AND TFR_TIPO_ETAPA = 2)" & _
                " ORDER BY TFR_COD_ETAPA", 0, 8000
        End If
    End If
End Sub

Public Sub PreencheGridRedeEtapas(Grid As Object, TipoEtapa As TipoEtapaRede, Optional Pai As String)
    Dim Sql As String
    Sql = "SELECT TFR_COD_ETAPA,TFR_DESCRICAO AS DESCRICAO," & _
        "TFR_PRE_REQUISITO AS ORDEM," & _
        "TFR_DATA_CADASTRO AS DATA_CADASTRO FROM TAB_FISCALIZACAO_REDE " & _
        "WHERE TFR_TIPO_ETAPA = " & TipoEtapa
    If Trim(Pai) <> "" Then
        Sql = Sql & " AND (TFR_COD_ETAPA_PAI = " & Pai & " OR TFR_COD_ETAPA_ORIGEM = " & Pai & ")"
    End If
    Sql = Sql & " ORDER BY TFR_PRE_REQUISITO"
    Grid.Preencher Bdados, Sql, 0, 5000, 1000, 1500
End Sub

Public Sub PreencheComboEtapas(Combo As Object, TipoEtapa As TipoEtapaRede, Optional Pai As String)
    Dim Sql As String
    Sql = "SELECT TFR_COD_ETAPA, TFR_DESCRICAO FROM TAB_FISCALIZACAO_REDE " & _
        "WHERE TFR_TIPO_ETAPA = " & TipoEtapa
    If Trim(Pai) <> "" Then Sql = Sql & " AND TFR_COD_ETAPA_PAI =" & Pai
    Combo.Preencher Bdados, Sql, 1
End Sub

Public Sub LimpaDadosRede()
    rCodEtapa = 0
    rCaminhoRpt = ""
    rCodEtapaPai = 0
    rCodFuncionario = 0
    rDataCadastro = ""
    rDescricao = ""
    rOrdem = 0
    rPrazo = 0
    rTipoEtapa = 0
    rUsuario = ""
    rCodParametroFundamento = 0
End Sub
