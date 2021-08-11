VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eProcedimentosFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Atributos As New eProcedimentoFiscal
Private GeraNum As ContaCorrente
 
Public Function CarregaProcedimentos()
    Dim sql As String
    Dim Condicao As String
    
    sql = " Select * from TAB_PROCEDIMENTO_FISCAL "
    
    If Atributos.Contribuinte <> "" Then
        If (Condicao = "") Then
            Condicao = " TPF_TCI_IM = '" & Atributos.Contribuinte & "'"
        Else
            Condicao = " AND TPF_TCI_IM = '" & Atributos.Contribuinte & "'"
        End If
    End If
    
    If Atributos.NumeroProcedimento <> "" Then
        If (Condicao = "") Then
            Condicao = " TPF_NUMERO_PROCEDIMENTO =" & Atributos.NumeroProcedimento
        Else
            Condicao = " AND TPF_NUMERO_PROCEDIMENTO =" & Atributos.NumeroProcedimento
        End If
    End If
    
    If Atributos.NumeroFiscalizacao <> "" Then
        If (Condicao = "") Then
            Condicao = " TPF_TFI_NUMERO_FISCALIZACAO =" & Atributos.NumeroFiscalizacao
        Else
            Condicao = " TPF_TFI_NUMERO_FISCALIZACAO =" & Atributos.NumeroFiscalizacao
        End If
    End If
    
    If Atributos.TipoAcaoFiscal <> "" Then
        If (Condicao = "") Then
            Condicao = " TPF_TTP_CODIGO_TIPO =" & Tipo
        Else
            Condicao = " TPF_TTP_CODIGO_TIPO =" & Tipo
        End If
    End If
    
    If Condicao <> "" Then sql = sql & " where " & Condicao
End Function


'ndicao As String
    
   
Private Function Gravar() As Boolean
    Dim Campos As String
    Dim Valores As String
    Dim Condicao As String
    
    Gravar = False
        
    Campos = " TPF_NUMERO_PROCEDIMENTO,TPF_TTP_CODIGO_TIPO,TPF_TIPO_ACAO,TPF_TCI_IM,"
    Campos = Campos & " TPF_TFU_MATRICULA,TPF_DESCRICAO_MOTIVOS,TPF_FUNDAMENTACAO_LEGAL,"
    Campos = Campos & " TPF_PREPOSTO_NOME,TPF_PREPOSTO_CPF,TPF_LAVRATURA_LOCAL, "
    Campos = Campos & " TPF_LAVRATURA_DATA,TPF_LAVRATURA_HORA,TPF_DOCUMENTOS_ORIGINARIOS,"
    Campos = Campos & " TPF_DOCUMENTOS_DECORRENTES,TPF_TFI_NUMERO_FISCALIZACAO"
    
    Valores = Bdados.PreparaValor(Atributos.NumeroProcedimento, Atributos.TipoProcedimento, _
    Atributos.TipoAcaoFiscal, Atributos.Contribuinte, Atributos.Autoridade.Atributos.Matricula, _
    Atributos.DescricaoMotivos, Atributos.FundamentacaoLegal, Atributos.RepresentantePassivoNome, _
    Atributos.RepresentantePassivoCPF, Atributos.LavraturaLocal, Atributos.LavraturaData, _
    Atributos.LavraturaHora, Atributos.DocumentosOriginarios, Atributos.DocumentosDecorrentes, Atributos.NumeroFiscalizacao)
    
    Condicao = " TPF_NUMERO_PROCEDIMENTO  = " & Atributos.NumeroProcedimento
    
    If Bdados.GravaDados("TAB_PROCEDIMENTO_FISCAL", Valores, Campos, Condicao) Then Gravar = True

End Function

Private Function GeraNumeroProcedimento()
    Select Case Atributos.TipoProcedimento
    
        Case tpProcedimentoTIAF
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(47)
            
        Case tpProcedimentoAITI
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(48)
            
        Case tpProcedimentoAPRE
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(49)
            
        Case tpProcedimentoATA_RIT
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(50)
            
        Case tpProcedimentoTEAF
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(52)
            
        Case tpProcedimentoTEDI
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(53)
            
        Case tpProcedimentoTI
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(54)
            
        Case tpProcedimentoTIFI
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(58)
            
        Case tpProcedimentoTREF
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(59)
            
        Case tpProcedimentoATA_RIF
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(60)
            
        Case tpProcedimentoATA_NIT
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(61)
            
        Case tpProcedimentoREFI
            Atributos.NumeroProcedimento = GeraNum.GeraCodPagamento(62)
            
    End Select
End Function

