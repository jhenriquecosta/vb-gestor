VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AndamentoFiscalizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum TipoStatusAndamento
    etaPendenteNoPrazo = 1
    etaPendenteAtrasado = 2
    etaConcluidoNoPrazo = 3
    etaConcluidoAtrasado = 4
End Enum

Public vCodSequenciaProcedimento As String
Public vCodFiscalizacao As String
Public vEtapa As New RedeFiscalizacao
Public vDataAbertura As String
Public vDataConclusao As String
Public vRelato As String
Public vDocumentos As String
Public vInformacoes As String
Public vCodFuncionaro As String
Public vUsuario As String
Public vDataDigitacao As String
Public vStatus As String

Public Function BuscaCodigoAndamento(CodFiscalizacao As String, CodEtapa As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "SELECT TFP_COD_SEQ_PROCEDIMENTO FROM TAB_FISCALIZACAO_ANDAMENTO WHERE " & _
            "TFP_TFI_COD_FISCALIZACAO =" & CodFiscalizacao & " AND TFP_TFR_COD_ETAPA =" & CodEtapa
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaCodigoAndamento = Rs!TFP_COD_SEQ_PROCEDIMENTO
    End If
End Function

Public Function CarregaAndamentoFiscalizacao(CodFiscalizacao As String, Optional CodAndamento As String, _
                Optional CodEtapa As String, Optional DataAbertura As String, Optional Status As TipoStatusAndamento) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    CarregaAndamentoFiscalizacao = False
    Sql = "SELECT TAB_FISCALIZACAO_ANDAMENTO.*,TAB_FISCALIZACAO_ANDAMENTO_DETALHE.* FROM TAB_FISCALIZACAO_ANDAMENTO LEFT JOIN TAB_FISCALIZACAO_ANDAMENTO_DETALHE ON TFP_COD_SEQ_PROCEDIMENTO = TFD_TFP_COD_SEQ_PROCEDIMENTO WHERE TFP_TFI_COD_FISCALIZACAO = " & CodFiscalizacao
    If Trim(CodAndamento) <> "" Then Sql = Sql & " AND TFP_COD_SEQ_PROCEDIMENTO =" & CodAndamento
    If Trim(CodEtapa) <> "" Then Sql = Sql & " AND TFP_TFR_COD_ETAPA =" & CodEtapa
    If Trim(DataAbertura) <> "" Then Sql = Sql & " AND TFP_DATA_ABERTURA =" & DataAbertura
    If Trim(Status) > 0 Then Sql = Sql & " AND TFP_STATUS =" & Status
    
    If Bdados.AbreTabela(Sql, Rs) Then
        vCodSequenciaProcedimento = Rs!TFP_COD_SEQ_PROCEDIMENTO
        vCodFiscalizacao = "" & Rs!TFP_TFI_COD_FISCALIZACAO
        vEtapa.rCodEtapa = "" & Rs!TFP_TFR_COD_ETAPA
        vDataAbertura = "" & Rs!TFP_DATA_ABERTURA
        vDataConclusao = "" & Rs!TFP_DATA_CONCLUSAO
        vRelato = "" & Rs!TFD_RELATOS
        vInformacoes = "" & Rs!TFD_INFORMACOES_SOLICITADAS
        vDocumentos = "" & Rs!TFD_DOCUMENTOS_SOLICITADOS
        vCodFuncionaro = "" & Rs!TFI_TFU_COD_FUNCIONARIO
        vUsuario = "" & Rs!TFI_TUS_COD_USUARIO
        vDataDigitacao = "" & Rs!TFP_DATA_DIGITACAO
        vStatus = "" & Rs!TFP_STATUS
        CarregaAndamentoFiscalizacao = True
    End If
End Function

Public Function GravaAndamentoProcesso(CodFiscalizacao As String, CodEtapa As String, _
        DataEtapa As String, DataConclusao As String, CodFuncionario As Integer, _
        Optional ByRef CodSeqProcedimento As String, _
        Optional DocumentosSolicitados As String, Optional InformacoesSolicitadas As String, _
        Optional Relatos As String) As Boolean
    
    Dim Valores As String
    Dim Campos As String
    Dim Conta As New ContaCorrente
    Dim Status As Integer
    GravaAndamentoProcesso = False
    CarregaAndamentoFiscalizacao CodFiscalizacao, CodSeqProcedimento
    vEtapa.CarregaDadosRede CodEtapa
            
    If Trim(CodSeqProcedimento) = "" Then
        CodSeqProcedimento = Conta.GeraCodPagamento("33")
    End If
    If Trim(DataConclusao) = "" Then
        If DateDiff("d", Date, DateAdd("d", vEtapa.rPrazo, DataEtapa)) >= 0 Then
            Status = 1
        Else
            Status = 2
        End If
    Else
        If DateDiff("d", DataConclusao, DateAdd("d", vEtapa.rPrazo, DataEtapa)) >= 0 Then
            Status = 3
        Else
            Status = 4
        End If
    End If
    Campos = "TFP_TFI_COD_FISCALIZACAO,TFP_TFR_COD_ETAPA,TFP_DATA_ABERTURA,TFP_DATA_CONCLUSAO," & _
        "TFI_TFU_COD_FUNCIONARIO,TFI_TUS_COD_USUARIO,TFP_DATA_DIGITACAO," & _
        "TFP_COD_SEQ_PROCEDIMENTO,TFP_STATUS"
    
    Valores = Bdados.PreparaValor(CodFiscalizacao, CodEtapa, _
        Bdados.Converte(DataEtapa, TCDataHora), _
        DataConclusao, CodFuncionario, AplicacoesVTFuncoes.Usuario, _
        Bdados.Converte(Date, TCDataHora), CodSeqProcedimento, Status)
    GravaAndamentoProcesso = Bdados.GravaDados("TAB_FISCALIZACAO_ANDAMENTO", Valores, Campos, "TFP_COD_SEQ_PROCEDIMENTO=" & CodSeqProcedimento)
    GravaDetalheAndamentoProcesso CodFiscalizacao, CodSeqProcedimento, DocumentosSolicitados, InformacoesSolicitadas, Relatos
    EncaminhaNovoProcedimento CodFiscalizacao, CodEtapa, Status
End Function


Public Function GravaDetalheAndamentoProcesso(CodFiscalizacao As String, CodSeqAndamento As String, Documentos As String, _
        Informacoes As String, Relatos As String)
    
    Dim Valores As String
    Dim Campos As String
    If Trim(Documentos) <> "" Or Trim(Informacoes) <> "" Or Trim(Relatos) <> "" Then
        Campos = "TFD_TFI_COD_FISCALIZACAO,TFD_TFP_COD_SEQ_PROCEDIMENTO,TFD_DOCUMENTOS_SOLICITADOS,TFD_INFORMACOES_SOLICITADAS,TFD_RELATOS"
        Valores = Bdados.PreparaValor(CodFiscalizacao, CodSeqAndamento, Documentos, Informacoes, Relatos)
        GravaDetalheAndamentoProcesso = Bdados.GravaDados("TAB_FISCALIZACAO_ANDAMENTO_DETALHE", Valores, Campos, "TFD_TFP_COD_SEQ_PROCEDIMENTO =" & CodSeqAndamento)
    End If
End Function

Public Sub PreencheAndamentoFiscalizacao(Grid As Object, CodProcesso As String)
    Grid.Preencher Bdados, "SELECT TFP_COD_SEQ_PROCEDIMENTO,TFP_TFI_COD_FISCALIZACAO,TFP_TFR_COD_ETAPA,TFR_DESCRICAO AS DESCRICAO,TGE_NOME AS STATUS,TFP_DATA_ABERTURA AS DATA_ABERTURA,TFP_DATA_CONCLUSAO AS DATA_FIM,TFU_NOME AS FUNCIONARIO" & _
        " FROM TAB_FISCALIZACAO_ANDAMENTO,TAB_FISCALIZACAO_REDE,TAB_FUNCIONARIO,VIS_STATUS_PROCEDIMENTO_FISCAL " & _
        " WHERE TFP_TFR_COD_ETAPA = TFR_COD_ETAPA AND TFI_TFU_COD_FUNCIONARIO = " & _
        " TFU_COD_FUNCIONARIO AND TFP_TFI_COD_FISCALIZACAO = " & CodProcesso & " AND TFP_STATUS = TGE_CODIGO ORDER BY TFP_COD_SEQ_PROCEDIMENTO", 0, 0, 0, 3500, 2200, 1500, 1200, 2000
End Sub

Private Sub EncaminhaNovoProcedimento(CodFiscalizacao As String, CodProcedimentoAtual As _
        String, StatusProcedimentoAtual As Integer)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim EtapaPai As String
    Dim NovaEtapa As String
    Dim NovaOrdem As Integer
    If StatusProcedimentoAtual > 2 Then
        If Me.vEtapa.rOrdem = 0 Then Exit Sub
        Sql = "select min(tfr_cod_etapa_pai) as Pai,min(tfr_cod_etapa) NovaEtapa,min(tfr_pre_requisito) as NovaOrdem from TAB_FISCALIZACAO_rede where  "
        Sql = Sql & " (tfr_cod_etapa_pai = (select tfr_cod_etapa_pai  from TAB_FISCALIZACAO_rede where tfr_cod_etapa = " & CodProcedimentoAtual & ")"
        Sql = Sql & " and tfr_pre_requisito > (select tfr_pre_requisito  from TAB_FISCALIZACAO_rede where tfr_cod_etapa = " & CodProcedimentoAtual & "))"
        Call Bdados.AbreTabela(Sql, Rs)
        If Not Rs.EOF And Trim("" & Rs!Pai) <> "" Then  'Encontrou novo procedimento dentro da mesma fase
            EtapaPai = "" & Rs!Pai
            NovaEtapa = "" & Rs!NovaEtapa
            NovaOrdem = "" & Rs!NovaOrdem
            Bdados.AtualizaDados "TAB_FISCALIZACAO", Bdados.PreparaValor(EtapaPai, NovaEtapa), _
                "TFI_FASE_ATUAL,TFI_PROCEDIMENTO_ATUAL", "TFI_COD_FISCALIZACAO =" & CodFiscalizacao
        Else 'Nova Fase a ser iniciada
            Sql = "select min(tfr_cod_etapa_pai) as Pai,min(tfr_cod_etapa) NovaEtapa,min(tfr_pre_requisito) as NovaOrdem "
            Sql = Sql & " from TAB_FISCALIZACAO_rede where tfr_cod_etapa_pai = (select min(tfr_cod_etapa_pai)  from "
            Sql = Sql & " TAB_FISCALIZACAO_rede where tfr_cod_etapa_pai > (select tfr_cod_etapa_pai  from "
            Sql = Sql & " TAB_FISCALIZACAO_rede where tfr_cod_etapa = " & CodProcedimentoAtual & ")) and tfr_pre_requisito > 0"
            Call Bdados.AbreTabela(Sql, Rs)
            If Not Rs.EOF And Trim("" & Rs!Pai) <> "" Then
                EtapaPai = "" & Rs!Pai
                NovaEtapa = "" & Rs!NovaEtapa
                NovaOrdem = "" & Rs!NovaOrdem
                Bdados.AtualizaDados "TAB_FISCALIZACAO", Bdados.PreparaValor(EtapaPai, NovaEtapa), _
                "TFI_FASE_ATUAL,TFI_PROCEDIMENTO_ATUAL", "TFI_COD_FISCALIZACAO =" & CodFiscalizacao
            Else 'Acabou a rede de procedimentos: Fim Do Processo
                Bdados.AtualizaDados "TAB_FISCALIZACAO", Bdados.PreparaValor(3), "TFI_STATUS", "TFI_COD_FISCALIZACAO =" & CodFiscalizacao
                Avisa "Processo Finalizado. Encerramento automático."
            End If
        End If
    End If
End Sub
