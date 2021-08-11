VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fiscalizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Rede As New RedeFiscalizacao
Public Andamento As New AndamentoFiscalizacao
Public Funcionario As New Funcionario
Public vIm As String
Public vDocumentos As String
Public vInformacoes  As String
Public vDataInicio  As String
Public vDataFim  As String
Public vDataDevolucao As String
Public vPeriodoFiscalizado As String
Public vFundamentacao  As String
Public vCodFuncionario As Integer
Public vStatus As Integer
Public vEtapaAtual  As String
Public vProcedimentoAtual  As String

Public Function PreencheGridFiscalizacao(Grid As Object, Optional Inscricao As String, Optional CodFiscalizacao As String, _
            Optional PeriodoInicial As String, Optional PeriodoFinal As String) As Boolean
    Dim Sql As String
    Sql = "Select TFI_COD_FISCALIZACAO AS Número,TFI_COD_TIAF AS TIAF,TFI_COD_TEAF AS TEAF, " & _
        "TFI_TCI_IM AS Inscrição,TFI_DATA_INICIO AS Data_Inicial,TFI_DATA_FIM AS Prazo_Final, TFI_DATA_TEAF AS Data_Encerramento," & _
        "TFI_PERIODO_FISCALIZADO AS Periodo,TGE_NOME AS Status,TFI_DATA_CANCELAMENTO Data_Cancelamento,TFI_FASE_ATUAL AS FA,TFI_PROCEDIMENTO_ATUAL as PA from " & _
        " TAB_FISCALIZACAO,VIS_STATUS_FISCALIZACAO WHERE TFI_STATUS = TGE_CODIGO "
    If Trim(CodFiscalizacao) <> "" Then
        Sql = Sql & " and TFI_COD_FISCALIZACAO = " & CodFiscalizacao
    End If
    If Trim(Inscricao) <> "" Then
        Sql = Sql & " and TFI_TCI_IM = '" & Inscricao & "'"
    End If
    If Trim(PeriodoInicial) <> "" And Trim(PeriodoFinal) <> "" Then
        Sql = Sql & " and (TFI_DATA_FISCALIZACAO >= " & _
            Bdados.Converte(PeriodoInicial, TCDataHora) & " AND TFI_DATA_FISCALIZACAO <=" & _
            Bdados.Converte(PeriodoFinal, TCDataHora) & ")"
    End If
    Grid.Preencher Bdados, Sql
End Function

Public Function CarregaDadosFiscalizacao(CodFiscalizacao As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    CarregaDadosFiscalizacao = False
    Sql = "select * from tab_fiscalizacao where tfi_cod_fiscalizacao = " & _
        CodFiscalizacao
        
    If Bdados.AbreTabela(Sql, Rs) Then
        vIm = Rs!TFI_TCI_IM
        vDocumentos = "" & Rs!TFI_DOCUMENTOS_SOLICITADOS
        vInformacoes = "" & Rs!TFI_INFORMACOES_SOLICITADAS
        vDataInicio = "" & Rs!TFI_DATA_INICIO
        vDataFim = "" & Rs!TFI_DATA_FIM
        vDataDevolucao = "" & Rs!TFI_DATA_DEVOLUCAO
        vPeriodoFiscalizado = "" & Rs!TFI_PERIODO_FISCALIZADO
        vFundamentacao = "" & Rs!TFI_FUNDAMENTACAO
        vCodFuncionario = "" & Rs!TFI_TFU_COD_FUNCIONARIO
        vEtapaAtual = "" & Rs!TFI_FASE_ATUAL
        vProcedimentoAtual = "" & Rs!TFI_PROCEDIMENTO_ATUAL
        vStatus = "" & Rs!TFI_STATUS
        CarregaDadosFiscalizacao = True
        If Trim(Andamento.vCodSequenciaProcedimento) <> "" Then
            Andamento.CarregaAndamentoFiscalizacao CodFiscalizacao, Andamento.vCodSequenciaProcedimento
        End If
    End If
End Function

Public Function CancelaFiscalizacao(CodFiscalizacao As String) As Boolean
    CancelaFiscalizacao = Bdados.AtualizaDados("TAB_FISCALIZACAO", Bdados.PreparaValor(2, Bdados.Converte(Date, TCDataHora), AplicacoesVTFuncoes.Usuario), _
            "TFI_STATUS,TFI_DATA_CANCELAMENTO,TIF_USUARIO_CANCELAMENTO", _
            "TFI_COD_FISCALIZACAO = " & CodFiscalizacao)
End Function

Public Function EncerraFiscalizacao(CodFiscalizacao As String, Data As String) As Boolean
    EncerraFiscalizacao = Bdados.AtualizaDados("TAB_FISCALIZACAO", Bdados.PreparaValor(3, Bdados.Converte(Data, TCDataHora), AplicacoesVTFuncoes.Usuario), "TFI_STATUS,TFI_DATA_TEAF,TIF_USUARIO_TEAF", "TFI_COD_FISCALIZACAO = " & CodFiscalizacao)
End Function

Public Function CriaFiscalizacao(Contribuinte As String, _
        DataInicio As String, DataFim As String, PeriodoFiscalizado As String, Funcionario As String, _
        FaseProcedimental As String) As Boolean
    Dim Conta As New ContaCorrente
    Dim Campos As String
    Dim Valores As String
    Dim CodFiscalizacao As String
    
    Dim Fase As String
    Dim Procedimento As String
    Dim Processo As String
    Processo = FaseProcedimental
    Fase = BuscaProcedimentoInicialDaFase(Processo, etrFase)
    Procedimento = BuscaProcedimentoInicialDaFase(Fase, etrProcedimento)
    CodFiscalizacao = Conta.GeraCodPagamento("77")
    Campos = "TFI_COD_FISCALIZACAO,TFI_TCI_IM," & _
        "TFI_DATA_FISCALIZACAO,TFI_DATA_INICIO,TFI_DATA_FIM,TFI_PERIODO_FISCALIZADO,TFI_TFU_COD_FUNCIONARIO," & _
        "TFI_STATUS,TFI_FASE_ATUAL,TFI_PROCEDIMENTO_ATUAL,TFI_TIPO_PROCESSO"
    Valores = Bdados.PreparaValor(CodFiscalizacao, Bdados.Converte(Contribuinte, tctexto), _
            Bdados.Converte(Date, TCDataHora), _
            Bdados.Converte(DataInicio, TCDataHora), Bdados.Converte(DataFim, TCDataHora), _
            PeriodoFiscalizado, Funcionario, 1, Fase, Procedimento, Processo)
    CriaFiscalizacao = Bdados.InsereDados("TAB_FISCALIZACAO", Valores, Campos)
End Function

Private Function BuscaProcedimentoInicialDaFase(Origem As String, Tipo As TipoEtapaRede) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "select min(tfr_cod_etapa) as CodEtapaInicial from tab_fiscalizacao_rede "
    If Tipo = etrProcedimento Then
        Sql = Sql & " where tfr_pre_requisito > 0 "
        Sql = Sql & " and tfr_cod_etapa_pai = " & Origem
    End If
    If Tipo = etrFase Then
        Sql = Sql & " where tfr_pre_requisito > 0 "
        Sql = Sql & " and tfr_cod_etapa_origem = " & Origem
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaProcedimentoInicialDaFase = "" & Rs!CodEtapaInicial
    End If
End Function
