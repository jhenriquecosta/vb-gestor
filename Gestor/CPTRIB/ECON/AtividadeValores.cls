VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtividadeValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public v_CAE As String
Public v_Ano As Integer
Public v_ValorAlvara As Double
Public v_ValorAliquotaTPPC As Double
Public v_ValorAliquotaSPL As Double
Public v_ValorAliquotaPJ As Double
Public v_ValorFixoAnualISS As Double
Public v_TipoMoeda As Integer
Public v_EstimativoAlvara As Boolean
Public v_EstimativoISS As Boolean
Public v_AtividaEstimada As New eAtividadeEstimada
Public v_FatorAlvara As Integer
Public v_FatorDescricao As String
Public Function Buscar(CodAtividade As String, Ano As Integer, Optional LimiteInferior As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Buscar = True
    Sql = "Select TAV_VALOR_ALVARA,TAV_ALIQUOTA_TPPC,TAV_ALIQUOTA_SPL,TAV_ALIQUOTA_PJ,TAV_VALOR_ANUAL_ISS_FIXO,TAV_TIPO_MOEDA,TAV_FATOR,TAV_FATOR_DESCRICAO " & _
        " FROM TAB_ATIVIDADE_VALOR_TRIBUTOS where tav_tae_cae =" & CodAtividade & " and tav_ano =" & Ano
    If Bdados.AbreTabela(Sql, Rs) Then
        v_Ano = Ano
        v_CAE = CodAtividade
        v_EstimativoAlvara = False
        v_EstimativoISS = False
        v_TipoMoeda = Nvl("" & Rs!TAV_TIPO_MOEDA, 0)
        v_ValorAlvara = Nvl("" & Rs!TAV_VALOR_ALVARA, 0)
        v_ValorAliquotaTPPC = Nvl("" & Rs!TAV_ALIQUOTA_TPPC, 0)
        v_ValorAliquotaSPL = Nvl("" & Rs!TAV_ALIQUOTA_SPL, 0)
        v_ValorAliquotaPJ = Nvl("" & Rs!TAV_ALIQUOTA_PJ, 0)
        v_ValorFixoAnualISS = Nvl("" & Rs!TAV_VALOR_ANUAL_ISS_FIXO, 0)
        v_FatorAlvara = Nvl("" & Rs!TAV_FATOR, 0)
        v_FatorDescricao = Nvl("" & Rs!TAV_FATOR_DESCRICAO, 0)
        Exit Function
    Else
        v_ValorAlvara = 0
        v_ValorAliquotaTPPC = 0
        v_ValorAliquotaSPL = 0
        v_ValorAliquotaPJ = 0
        v_ValorFixoAnualISS = 0
        v_FatorAlvara = 0
        v_FatorDescricao = ""
        If v_AtividaEstimada.Buscar(CodAtividade, LimiteInferior, Ano) Then
            v_EstimativoAlvara = True
            Exit Function
        Else
            Buscar = False
        End If
    End If
End Function

Public Function Gravar() As Boolean
    Dim Valores As String
    Dim Campos As String
'    If v_EstimativoAlvara Then
'        v_AtividaEstimada.Gravar
'        Bdados.DeletaDados "TAB_ATIVIDADE_VALOR_TRIBUTOS", "TAV_TAE_CAE =" & v_CAE & " and TAV_ANO =" & v_Ano
'    Else
        Campos = "TAV_TAE_CAE,TAV_ANO,TAV_VALOR_ALVARA,TAV_FATOR,TAV_FATOR_DESCRICAO,TAV_TIPO_MOEDA,TAV_VALOR_ANUAL_ISS_FIXO,TAV_ALIQUOTA_PJ,TAV_ALIQUOTA_SPL,TAV_ALIQUOTA_TPPC"
        Valores = Bdados.PreparaValor(v_CAE, v_Ano, v_ValorAlvara, v_FatorAlvara, v_FatorDescricao, v_TipoMoeda, v_ValorFixoAnualISS, v_ValorAliquotaPJ, v_ValorAliquotaSPL, v_ValorAliquotaTPPC)
        Gravar = Bdados.GravaDados("TAB_ATIVIDADE_VALOR_TRIBUTOS", Valores, Campos, _
            "TAV_TAE_CAE =" & v_CAE & " and TAV_ANO =" & v_Ano)
        If Not v_EstimativoAlvara Then Bdados.DeletaDados "TAB_ATIVIDADE_ESTIMADA", "TAT_TAE_CAE = " & v_CAE & " AND TAT_ANO = " & v_Ano
'    End If
End Function


