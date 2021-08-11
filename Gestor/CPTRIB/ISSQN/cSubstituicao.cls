VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubstituicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================='
'               VsTIss.cSubstituicao: Usado em TRET101                   '
'                   Feito por Djeison  -  14/04/2003                            '
'=========================================='
Option Explicit
Private m_Nota As cNota

Public Property Get Nota() As cNota
    Set Nota = m_Nota
End Property

Public Property Set Nota(ByVal Value As cNota)
    Set m_Nota = Value
End Property

Private Sub Class_Initialize()
    Set m_Nota = New cNota
End Sub

Private Sub Class_Terminate()
    Set m_Nota = Nothing
End Sub

Public Function Salvar(NovoRemetente As Boolean, ImToma As String, DataAtual As Date) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim Conta As New ContaCorrente
    Dim Incidencia As Integer
    
    Salvar = False
    If NovoRemetente Then m_Nota.GravaNovoRemetente
    m_Nota.Cod_Pagamento = Conta.GeraCodPagamento(m_Nota.Cod_Imposto)
    
    'GRAVA PAGAMENTO
    If Not GravaPagamento(ImToma) Then
        Util.Erro "Erro em GravaPagamento (cSubstituicao)"
        Exit Function
    End If
    'GRAVA DETALHES DA NOTA
    If Not GravaDetalheNota Then
        Util.Erro "Erro em GravaDetalheNota (cSubstituicao)"
        Exit Function
    End If
    Call Conta.CriaContaContribuinte(m_Nota.Cod_Pagamento)
    'Call Conta.MovimentaContaContribuinte(m_Nota.Cod_Pagamento)
    'GRAVA DETALHE SUBSTITUTO
    If Not GravaDetalheSubstituto(ImToma, DataAtual) Then
        Util.Erro "Erro em GravaDetalheSubstituto (cSubstituicao)"
        Exit Function
    End If
    Salvar = True
       
End Function


Private Function GravaPagamento(ImToma As String) As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim Incidencia As Integer
    
    Campos = "tgt_im,tgt_tip_cod_imposto,tgt_periodo,tgt_data_vencimento,tgt_valor_tributo,tgt_cod_pagamento,tgt_data_geracao,tgt_incidencia"
    Incidencia = Imposto.BuscaNumeroIncidencia(ImToma, IIf(Len(Trim(m_Nota.Periodo_Ref)) = 4, m_Nota.Periodo_Ref, CLng(Right(m_Nota.Periodo_Ref, 4) & Left(m_Nota.Periodo_Ref, 2))), m_Nota.Cod_Imposto)
    Valores = Bdados.PreparaValor(ImToma, m_Nota.Cod_Imposto, IIf(Len(Trim(m_Nota.Periodo_Ref)) = 4, m_Nota.Periodo_Ref, CLng(Right(m_Nota.Periodo_Ref, 4) & Left(m_Nota.Periodo_Ref, 2))), Bdados.Converte(m_Nota.Data_Venc, TCDataHora), Bdados.Converte(m_Nota.ISS_Devido, TCDuplo), m_Nota.Cod_Pagamento, Bdados.Converte(Imposto.BuscaDataGeracaoDam(m_Nota.Cod_Pagamento), TCDataHora), Incidencia)
    GravaPagamento = Bdados.GravaDados("Tab_Geracao_tributo", Valores, Campos, "tgt_cod_pagamento=" & m_Nota.Cod_Pagamento)
End Function

Private Function GravaDetalheNota() As Boolean
    Dim Valores As String
    Dim Campos As String
    
    Valores = Bdados.PreparaValor(m_Nota.Cod_Pagamento, m_Nota.Nota_fiscal, m_Nota.Nota_fiscal, Bdados.Converte(m_Nota.Total_Nota, TCDuplo), Bdados.Converte(Nvl(m_Nota.Valor_Material_ICMS, 0), TCDuplo), "")
    Campos = "tdd_tgt_cod_pagamento,tdd_num_nota_inicial,tdd_num_nota_final,tdd_total_nota,tdd_total_material_reducao,tdd_obs"
    GravaDetalheNota = Bdados.GravaDados("TAB_DETALHE_DAM", Valores, Campos, "tdd_tgt_cod_pagamento=" & m_Nota.Cod_Pagamento)
End Function

Private Function GravaDetalheSubstituto(ImToma As String, DataAtual As Date) As Boolean
    Dim Valores As String
    Dim Campos As String
    
    Valores = Bdados.PreparaValor(m_Nota.Cod_Pagamento, m_Nota.Nota_fiscal, ImToma, IIf(Len(m_Nota.NumIM) <> 0, m_Nota.NumIM, Const_ImAvulso), Bdados.Converte(m_Nota.Data_emissao, TCDataHora), _
                    Bdados.Converte(Nvl(m_Nota.Valor_Material_ICMS, 0), TCDuplo), Nvl(m_Nota.Total_Nota, 0), Bdados.Converte(m_Nota.ISS_Devido, TCDuplo), Bdados.Converte(DataAtual, TCDataHora), Right(m_Nota.Periodo_Ref, 4) & Left(m_Nota.Periodo_Ref, 2), m_Nota.Cod_Imposto, m_Nota.Usuario)
    Campos = "TNS_TGT_COD_PAGAMENTO,TNS_NUM_NOTA,TNS_TCI_IM_SUBSTITUTO,TNS_TCI_IM_PRESTADOR," & _
                    "TNS_DATA_EMISSAO_NOTA,TNS_VALOR_MATERIAL_NOTA,TNS_VALOR_NOTA," & _
                    "TNS_VALOR_DEVIDO,TNS_DATA_RECEPCAO,TNS_PERIODO,TNS_TIP_COD_IMPOSTO,TNS_TUS_COD_USUARIO"
    If Trim(m_Nota.Data_Recolhimento) <> "" Then
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(m_Nota.Data_Recolhimento, TCDataHora))
        Campos = Campos & ",TNS_DATA_RECOLHIMENTO"
    End If
    GravaDetalheSubstituto = Bdados.GravaDados("TAB_NOTA_SUBSTITUTA", Valores, Campos, "TNS_TGT_COD_PAGAMENTO=" & m_Nota.Cod_Pagamento)
End Function

Public Function GravaDadosBaixa(PeriodoImposto As String, DataAtual As Date) As Boolean
    On Error GoTo trata
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim Rs As VSRecordset
    
    GravaDadosBaixa = True
    Valores = Bdados.PreparaValor(m_Nota.IM_CPF, m_Nota.Cod_Imposto, Bdados.Converte(m_Nota.Data_Venc, TCDataHora), _
    IIf(Len(m_Nota.Periodo_Ref) = 4, m_Nota.Periodo_Ref, Right(m_Nota.Periodo_Ref, 4) & Left(m_Nota.Periodo_Ref, 2)), Bdados.Converte(m_Nota.Data_Recolhimento, TCDataHora), Bdados.Converte(m_Nota.ISS_Devido, TCDuplo), _
    Bdados.Converte(m_Nota.ISS_Devido, TCDuplo), Bdados.Converte(DataAtual, TCDataHora), m_Nota.Usuario, m_Nota.Cod_Pagamento, Bdados.Converte(m_Nota.ISS_Devido, TCDuplo))
    
    Campos = "tdr_im,tdr_tip_cod_imposto,tdr_data_vencimento,tdr_periodo," & _
        "tdr_data_pagamento,tdr_valor_original," & _
        "tdr_valor_total,tdr_data_entrada,tdr_tus_cod_usuario,tdr_tgt_cod_pagamento,tdr_valor_real_pago"
    Call Bdados.GravaDados("Tab_Darm_Recebido", Valores, Campos, "tdr_tgt_cod_pagamento=" & m_Nota.Cod_Pagamento)
    Exit Function
trata:
    GravaDadosBaixa = False
End Function
Public Sub PreencheGrid(GridSubs As Object, GridPres As Object, Razao As String, Index As Integer)
    Dim Rs As VSRecordset
    Dim Sql As String
    
    Sql = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf as CPF_CGC from Tab_Contribuinte where tci_nome like '" & Razao & "%' or tci_nome like '% " & Razao & "%'"
    Sql = Sql & " and tci_tsc_cod_sit_cad =1"
    '-------------------------------------teste----------------------------------------
    If Index = 0 Then Sql = Sql & " AND tci_tipo_recolhimento_iss =2"
    If Not Bdados.AbreTabela(Sql, Rs) Then
        Call Util.Avisa("Nenhum contribuinte encontrado.")
    End If
    Bdados.FechaTabela Rs
     '----------------------------------------------------------------------------------
    If Index = 0 Then
        GridSubs.Preencher Bdados, Sql, 15, 57, 25
    Else
        GridPres.Preencher Bdados, Sql, 15, 57, 25
    End If
End Sub

Public Sub BuscaAliquota(DataAtual As Date, Periodo As String)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Condicao As String
    Dim Dia As String
    
    m_Nota.Cod_Imposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQNSUBST))
    Sql = "Select tpi_aliquota,tip_nome_imposto,tip_cod_imposto from tab_parametro_imposto,tab_imposto where tpi_tip_cod_imposto = '" & m_Nota.Cod_Imposto & "' and tpi_tip_cod_imposto = tip_cod_imposto"
        
    Dia = IIf(Trim(Periodo) = "", DataAtual, Periodo)
    Condicao = " and tpi_ano_imposto= '" & Imposto.BuscaAnoImposto(m_Nota.Cod_Imposto, Right(Year(Dia), 4)) & "'"
    If Bdados.AbreTabela(Sql & Condicao, Rs) Then
        m_Nota.Aliquota = Format(Rs(0) / 100, Const_Monetario)
        m_Nota.Cod_Imposto = Rs(2)
        m_Nota.Nome_Imposto = Rs(1)
    End If
    Bdados.FechaTabela Rs
End Sub


Public Function BuscarTomador(ImToma As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Impost As New VSImposto
    
    Sql = "Select * FROM tab_Contribuinte where tci_im='" & ImToma & "' or tci_cgc_cpf='" & ImToma & "' and tci_tsc_cod_sit_cad =1 " 'and tci_tipo_recolhimento_iss =2
    
    If Not Bdados.AbreTabela(Sql, Rs) Then
        BuscarTomador = False
        Exit Function
    Else
        BuscarTomador = True
        m_Nota.IM_CPF = "" & Rs!tci_im
        m_Nota.NumCGC = "" & Rs!TCI_CGC_CPF
        If Bdados.AbreTabela("SELECT tae_nome FROM Tab_Atividade_Economica WHERE tae_cae = " & Util.Nvl("" & Rs!tci_tae_cae, 0)) Then
            m_Nota.Atividade = Bdados.Tabela!tae_nome
        End If
        Bdados.FechaTabela
        m_Nota.Nome_Empresa = "" & Rs!tci_nome
        m_Nota.Endereco.Endereco = "" & Rs!tci_logradouro & " " & Rs!tci_nome_logradouro & ", " & Rs!tci_NUMERO & IIf(Trim(Rs!tci_COMPLEMENTO) = "", "", ", " & Rs!tci_COMPLEMENTO) & ", " & Rs!tci_BAIRRO
    End If
    Bdados.FechaTabela Rs
End Function

Public Sub PreencherCboAgente(cbo As Object)
    Dim Sql As String
    Sql = "Select tar_nome_agente, tar_cod_agente from tab_agente_arrecadador where tar_ativo =0 order by tar_nome_agente"
    cbo.Preencher Bdados, Sql
End Sub
