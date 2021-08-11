VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeclaracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const cteItemImpostoDevido As Byte = 13

Public Enum eTipoDeclaracao
    decNormal = 1
    decNegativa = 2
    decSubstitutiva = 3
End Enum

Public Enum eStatusDeclaracao
    decAberta = 1
    decFinalizada = 2
    decGerada = 3
    decNaoAberta = 4
End Enum


Public Enum eOrigemDeclaracao
    orgInternet = 1
    orgSistema = 2
End Enum


Public Enum eFormatoPeriodo
    perAAAMM = 1
    perMMAAAA = 2
End Enum

Private varNumero As String
Private varIM As String
Private varPeriodo As String
Private varTipo As eTipoDeclaracao
Private varDeclaracao As Date
Private varRecepcao As Date
Private varOrigem As eOrigemDeclaracao
Private varItens As cItensDeclaracao
Private varNotas As NotasFiscais
Private varCodImposto As String
Private varVersao As String
Private varBase As Double
Private varStatus As Double
Private varTomador As String

Public tciNome As String
Public tciEndereco As String
Public tciBairro As String
Public tciCEP As String
Public tciCidade As String
Public tciUF As String
Public tciAtividade As String

Private Sub ImprimeDAMDeclaracao(CodObrigacao As String, Optional stringTaxas As String, Optional TotalTaxas As Double)
    Dim Imposto As New VSImposto
    Dim Obrigacao As New Obrigacao
    Dim Cobranca As New VSCobranca
    Dim CodImposto As String, Sigla As String
    Dim MSG As String
    If CDbl(Nvl(Temp.PegaParametro(Bdados, "VERSAO DEC"), 0)) = 3 Then Exit Sub
    Obrigacao.BuscaDetalheObrigacao CodObrigacao, etiContribuinte
    If Util.Confirma("Imprimir DAM?") Then
        
        
    
        Set Cobranca = New VSCobranca
        Dim NomeTributo As String
        Dim NovaData As String
        Dim NovaMulta As Double
        Dim NovoJuro As Double
        Dim NotaInicial As String
        Dim NotaFinal As String
        
        NovaData = Imposto.DataVencimentoNova(Format(Obrigacao.obDataVencimento, "dd/mm/yyyy"))
        MSG = Entrada("Observações", "Mensagem")
        
        If Trim(NovaData) = "" Then Exit Sub
        'Obrigacao.BuscaDetalheObrigacao (CodObrigacao)
        NomeTributo = Imposto.BuscaNomeCompletoImposto(Obrigacao.obCodImposto)
        NovaMulta = Conta.CalculaValoresMultaAvulsos(Obrigacao.obCodImposto, Obrigacao.obPeriodo, EtcCreditoTributario, _
                    NovaData, Obrigacao.obDataVencimento, Obrigacao.obValorObrigacao)
        NovoJuro = Conta.CalculaValoresJurosAvulsos(Obrigacao.obCodImposto, Obrigacao.obPeriodo, EtcCreditoTributario, _
                    NovaData, Obrigacao.obDataVencimento, Obrigacao.obValorObrigacao)
        
'        If Obrigacao.obCodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN)) Then
            If Bdados.AbreTabela("SELECT TDD_VALOR_ITEM FROM TAB_DETALHE_DECLARACAO WHERE TDD_TDC_NUM_DECLARACAO=" & _
                            varNumero & " AND TDD_TCD_COD_ITEM IN (1,2) ORDER BY TDD_TCD_COD_ITEM ASC") Then
                Bdados.Tabela.MoveFirst
                NotaInicial = Format(Bdados.Tabela(0) & "", "000")
                Bdados.Tabela.MoveNext
                If Not Bdados.Tabela.EOF Then NotaFinal = Format(Bdados.Tabela(0) & "", "000")
                Bdados.FechaTabela
            End If
'        End If
'        NotaInicial = ""
'        NotaFinal = ""
        Bdados.FechaTabela
        Dim Rs As VSRecordset
        Dim Sql As String
        Dim obs As String
        Sql = "SELECT TIP_SIGLA_IMPOSTO FROM TAB_IMPOSTO WHERE TIP_COD_IMPOSTO ='" & Obrigacao.obCodImposto & "'"
        
        'bcp
        obs = "ISSQN REF: NF(s) " & NotaInicial & " a " & NotaFinal & " no periodo de " & Right(Me.Periodo, 2) & " / " & Left(Me.Periodo, 4)
        If Bdados.Executa("UPDATE TAB_OBRIGACAO_CONTRIBUINTE SET TOC_OBS='" & obs & "' where toc_cod_obrigacao=" & Obrigacao.obCodigoObrigacao) Then
            
        End If
        If Bdados.AbreTabela(Sql, Rs) Then Sigla = Trim("" & Rs!TIP_SIGLA_IMPOSTO)
        Cobranca.ImprimeDam Rpt, Obrigacao.obCodigoObrigacao, Obrigacao.obContribuinte, tciNome, "", tciEndereco, _
             "", "", Obrigacao.obCodImposto, _
            Sigla, NomeTributo, Right(Me.Periodo, 2) & Left(Me.Periodo, 4), "0", 1, NovaData, Me.BaseGeral, _
            Obrigacao.obValorObrigacao, CStr(NovaMulta), CStr(NovoJuro), TotalTaxas, 0, 0, MSG & " " & obs, _
             , NotaInicial, NotaFinal, , , , , , , , , tdiTela, etdNormal
    End If
End Sub

Public Property Let Numero(Value As String)
    varNumero = Value
End Property

Public Property Get Numero() As String
    Numero = varNumero
End Property


Public Property Let Im(Value As String)
    varIM = Value
End Property
Public Property Get Im() As String
    Im = varIM
End Property

Public Property Let Tomador(Value As String)
    varTomador = Value
End Property
Public Property Get Tomador() As String
    Tomador = varTomador
End Property

Public Property Let CodTributo(Value As String)
    varCodImposto = Value
End Property
Public Property Get CodTributo() As String
    CodTributo = varCodImposto
End Property

Public Sub PreencheNotas()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Item As NotaFiscal
    
    Sql = "SELECT * FROM TAB_NOTA_FISCAL WHERE TNF_TDC_NUM_DECLARACAO=" & varNumero & " ORDER BY TNF_COD_OPERACAO ASC"
    If Bdados.AbreTabela(Sql, Rs) Then
        Do While Not Rs.EOF
            Set Item = New NotaFiscal
            Item.Aliquota = "" & Rs!TNS_ALIQUOTA
            Item.BaseCalculo = "" & Rs!TNS_BASE_CALCULO
            Item.Data = "" & Rs!TNS_DATA_NOTA
            Item.Destinatario = "" & Rs!TNF_INSCRICAO_OPERACAO
            Item.ImpostoDevido = "" & Rs!TNS_VALOR_IMPOSTO_DEVIDO
            Item.ImpostoRetido = "" & Rs!TNS_VALOR_IMPOSTO_RETIDO
            Item.Numero = Rs!TNF_NUM_NOTA
            Item.Status = "" & Rs!TNS_NOTA_CANCELADA
            Item.TipoOperacao = "" & Rs!TNF_COD_OPERACAO
            Item.ValorMaterialICMS = "" & Rs!TNS_VALOR_MATERIAL_ICMS
            Item.ValorTotal = "" & Rs!TNS_VALOR_NOTA
            Item.Tomador = "" & Rs!TNS_TCI_IM_TOMADOR
            Me.Notas.Adicionar Item
            Rs.MoveNext
        Loop
    End If
    Bdados.FechaTabela Rs
End Sub

Public Property Let Status(Value As eStatusDeclaracao)
    varStatus = Value
End Property

Public Property Get Status() As eStatusDeclaracao
    Status = varStatus
End Property

Public Property Let Periodo(Value As String)
    varPeriodo = Value
End Property
Public Property Get Periodo() As String
    Periodo = varPeriodo
End Property

Public Sub CarregaGrid(Grid As Object, Optional Im As String, Optional Periodo As String, Optional TipoDec As eTipoDeclaracao = -1, Optional StatusDec As eStatusDeclaracao = -1, Optional CodTributo As String)
    Dim Aux As Byte
    Dim Sql As String
    Sql = "SELECT TDC_TCI_IM AS Contribuinte,TDC_PERIODO as Periodo,TDC_TIPO_DEC,TDC_STATUS,VIS_TIPO_DECLARACAO.TGE_NOME as Tipo,VIS_STATUS_DECLARACAO.TGE_NOME AS Status,TDC_NUM_DECLARACAO" & _
            ",TDC_TIP_COD_IMPOSTO FROM TAB_DECLARACAO_CONTRIBUINTE,VIS_STATUS_DECLARACAO,VIS_TIPO_DECLARACAO  WHERE TDC_TIPO_DEC = VIS_TIPO_DECLARACAO.TGE_CODIGO AND TDC_STATUS = VIS_STATUS_DECLARACAO.TGE_CODIGO "
    
    If Trim(Im) <> "" Then Sql = Sql & " AND TDC_TCI_IM ='" & Im & "'"
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        If Trim(Periodo) <> "" Then Sql = Sql & " AND TDC_PERIODO=" & Right(Periodo, 4) & Left(Periodo, 2)
    Else
        If Trim(Periodo) <> "" Then Sql = Sql & " AND TDC_PERIODO=" & Left(Periodo, 2) & Right(Periodo, 4)
    End If
    If Trim(CodTributo) <> "" Then Sql = Sql & " AND TDC_TIP_COD_IMPOSTO ='" & CodTributo & "'"
    If TipoDec > 0 Then Sql = Sql & " AND TDC_TIPO_DEC=" & TipoDec
    If StatusDec > 0 Then
        If StatusDec = decNaoAberta Then
            Sql = Sql & " AND TDC_STATUS in (" & decFinalizada & "," & decGerada & ")"
        Else
            Sql = Sql & " AND TDC_STATUS=" & StatusDec
        End If
    End If
    If Not Grid Is Nothing Then Grid.Preencher Bdados, Sql, 2000, 1200, 0, 0, 1500, 1500, 0, 0
End Sub

Public Sub PreencheCamposDeclaracao(ObjEntradas As Object, ObjSaidas As Object, NotaInicial As Object, NotaFinal As Object, TotalNotas As Object, TotalIcms As Object)
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Grid As Object
    
    
    Sql = "SELECT TDD_VALOR_ITEM,TDD_TCD_COD_ITEM FROM TAB_DETALHE_DECLARACAO WHERE TDD_TDC_NUM_DECLARACAO =" & varNumero & _
    " AND TDD_TCD_COD_ITEM IN (1,2,3,4) ORDER BY TDD_TCD_COD_ITEM  ASC"
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        If Not Rs.EOF Then NotaInicial = Rs!TDD_VALOR_ITEM
        If Not Rs.EOF Then Rs.MoveNext
        If Not Rs.EOF Then NotaFinal = Rs!TDD_VALOR_ITEM
        If Not Rs.EOF Then Rs.MoveNext
        If Not Rs.EOF Then TotalNotas = "" & Rs!TDD_VALOR_ITEM
        If Not Rs.EOF Then Rs.MoveNext
        If Not Rs.EOF Then TotalIcms = "" & Rs!TDD_VALOR_ITEM
    End If
    
    Sql = "SELECT * FROM TAB_NOTA_FISCAL WHERE TNF_TDC_NUM_DECLARACAO=" & varNumero & " ORDER BY TNF_COD_OPERACAO ASC"
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do
            If "" & Rs!TNF_COD_OPERACAO = 1 Then
                Set Grid = ObjEntradas.ListItems.Add(, , Rs!TNF_INSCRICAO_OPERACAO)
                
            Else
                Set Grid = ObjSaidas.ListItems.Add(, , Rs!TNF_INSCRICAO_OPERACAO)
            End If
            Grid.SubItems(1) = "" & Rs!TNF_NUM_NOTA
            Grid.SubItems(2) = "" & Rs!TNS_DATA_NOTA
            Grid.SubItems(3) = "" & Rs!TNS_VALOR_NOTA
            Grid.SubItems(4) = "" & Rs!TNS_VALOR_MATERIAL_ICMS
            Grid.SubItems(5) = "" & Rs!TNS_BASE_CALCULO
            Grid.SubItems(6) = "" & Rs!TNS_VALOR_IMPOSTO_DEVIDO
            Grid.SubItems(7) = "" & Rs!TNS_VALOR_IMPOSTO_RETIDO
            If "" & Rs!TNF_COD_OPERACAO = 1 Then
                Grid.SubItems(8) = "" & Rs!TNS_VALOR_IMPOSTO_RETIDO
            Else
                Grid.SubItems(8) = Nvl(Rs!TNS_VALOR_IMPOSTO_DEVIDO, 0) - Nvl(Rs!TNS_VALOR_IMPOSTO_RETIDO, 0)
            End If
            Grid.SubItems(10) = "" & Rs!TNS_ALIQUOTA
            Grid.SubItems(9) = "" & Rs!TNS_NOTA_CANCELADA
            Grid.SubItems(11) = "" & Rs!TNS_AIDF
            Grid.SubItems(12) = "" & Rs!TNS_TCI_IM_TOMADOR
            Rs.MoveNext
        Loop While Not Rs.EOF
    End If
End Sub

Public Property Let Tipo(Value As eTipoDeclaracao)
    varTipo = Value
End Property

Public Property Get Tipo() As eTipoDeclaracao
    Tipo = varTipo
End Property

Public Property Let Data(Value As Date)
    varDeclaracao = Value
End Property
Public Property Get Data() As Date
    Data = varDeclaracao
End Property


Public Property Let Recepcao(Value As Date)
    varRecepcao = Value
End Property

Public Property Get Recepcao() As Date
    Recepcao = varRecepcao
End Property


Public Property Let Origem(Value As eOrigemDeclaracao)
    varOrigem = Value
End Property
Public Property Get Origem() As eOrigemDeclaracao
    Origem = varOrigem
End Property

Public Property Get Itens() As cItensDeclaracao
    Set Itens = varItens
End Property

Public Property Get Notas() As NotasFiscais
    Set Notas = varNotas
End Property

Public Property Let Versao(Value As String)
    varVersao = Value
End Property

Public Property Get Versao() As String
    Versao = varVersao
End Property

Public Property Get BaseGeral() As Double
    BaseGeral = varBase
End Property

Public Property Let BaseGeral(Value As Double)
    varBase = Value
End Property

Private Sub Class_Initialize()
    Set varItens = New cItensDeclaracao
    Set varNotas = New NotasFiscais
End Sub
Private Sub Class_Terminate()
    Set varItens = Nothing
End Sub

Public Function Buscar(Im As String, Periodo As String, Optional TipoDeclaracao As eTipoDeclaracao = -1, Optional CodTributo As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    
    If Im = "" Or Periodo = "" Then Exit Function
    If Left(Periodo, 2) < 13 Then Periodo = Right(Periodo, 4) & Left(Periodo, 2)
    If TipoDeclaracao = decSubstitutiva Then
        Sql = "SELECT  TDC_STATUS from TAB_DECLARACAO_CONTRIBUINTE WHERE TDC_TCI_IM ='" & Im & "'" & _
                " AND TDC_PERIODO =" & Periodo & " and TDC_TIPO_DEC = " & decNormal & "  and TDC_TIP_COD_IMPOSTO ='" & CodTributo & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            If Rs!TDC_STATUS = 1 Then
                Avisa "A Declaracão original não foi finalizada. Não é possível gerar uma Retificadaora."
                Buscar = False
                Exit Function
            End If
        End If
    End If
    Sql = "SELECT TDC_TCI_IM,TDC_PERIODO,TDC_NUM_DECLARACAO,TDC_ORIGEM_DECLARACAO,TDC_VERSAO,TDC_DATA_RECEPCAO,TDC_DATA_DECLARACAO,TDC_TIPO_DEC,TDC_STATUS " & _
            " FROM TAB_DECLARACAO_CONTRIBUINTE" & _
            " WHERE TDC_TCI_IM ='" & Im & "'" & _
                " AND TDC_PERIODO =" & Periodo & "  and TDC_TIP_COD_IMPOSTO ='" & CodTributo & _
                "' AND TDC_TIPO_DEC=" & IIf(TipoDeclaracao <> -1, TipoDeclaracao, _
                " (SELECT MAX(TDC_TIPO_DEC)" & _
                " FROM TAB_DECLARACAO_CONTRIBUINTE" & _
                " WHERE TDC_TCI_IM ='" & Im & "'" & _
                " AND TDC_PERIODO =" & Periodo & ")")
    
    If Bdados.AbreTabela(Sql, Rs) Then
        If Nvl("" & Rs!TDC_STATUS, 1) = 2 And Nvl("" & Rs!TDC_TIPO_DEC, 0) = TipoDeclaracao Then
            If TipoDeclaracao = decSubstitutiva Then
                Avisa "Declaracão finalizada. Alteracão não permitida."
                Buscar = False
                Exit Function
            ElseIf TipoDeclaracao = decNormal Then
                Avisa "Declaracão finalizada. Alteracão permitida somente através de retificadora."
                TipoDeclaracao = decSubstitutiva
            
            End If
        End If
        varNumero = Trim(Nvl("" & Rs!TDC_NUM_DECLARACAO, 0))
        
        varIM = "" & Trim(Rs!TDC_TCI_IM)
        varPeriodo = Trim(Nvl("" & Rs!TDC_PERIODO, 0))
        varTipo = Trim(Nvl("" & Rs!TDC_TIPO_DEC, 1))
        varDeclaracao = Trim("" & Rs!TDC_DATA_DECLARACAO)
        varRecepcao = Trim("" & Rs!TDC_DATA_RECEPCAO)
        varOrigem = Trim(Nvl("" & Rs!TDC_ORIGEM_DECLARACAO, 0))
        varVersao = Trim("" & Rs!TDC_VERSAO)
        varTipo = Trim(TipoDeclaracao)
        varStatus = Trim(Nvl("" & Rs!TDC_STATUS, 1))
        
        PreencherItens
        Buscar = True
    End If
    Bdados.FechaTabela Rs
End Function


Public Sub PreencherItens()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Item As cItemDeclaracao
    Dim Documento As String
    
    Sql = "SELECT * " & _
            " FROM TAB_DETALHE_DECLARACAO" & _
            " WHERE TDD_TDC_NUM_DECLARACAO=" & varNumero & _
            " ORDER BY TDD_TCD_COD_ITEM, TDD_NUM_DOC"
    If Bdados.AbreTabela(Sql, Rs) Then
        Do While Not Rs.EOF
            Set Item = New cItemDeclaracao
            Item.Numero = "" & Rs!TDD_TCD_COD_ITEM
            Item.Documento = "" & Rs!TDD_NUM_DOC
            Item.Valor = "" & Rs!TDD_VALOR_ITEM
            Me.Itens.Adicionar Item
            Rs.MoveNext
        Loop
    End If
    Bdados.FechaTabela Rs
End Sub

Public Property Get ImpostoDevido() As Double
    ImpostoDevido = Me.Itens.Total(cteItemImpostoDevido)
End Property

Public Function Gravar() As Boolean
    Dim Valores As String, Campos As String
    Dim I As Integer
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim ValorItem As String
    Dim NotaInicial As String, NotaFinal As String
    If Len(varPeriodo) <> 4 Then
        varPeriodo = Right(varPeriodo, 4) & Left(varPeriodo, 2)
    End If
    Sql = "Select TDC_NUM_DECLARACAO from tab_declaracao_contribuinte where " & _
            " TDC_PERIODO =" & varPeriodo & " and TDC_TCI_IM ='" & varIM & "' and TDC_TIPO_DEC = " & varTipo & " AND TDC_TIP_COD_IMPOSTO ='" & varCodImposto & "'"
    If Not Bdados.AbreTabela(Sql, Rs) Then
        If Bdados.Conexao.FormatoBanco = SQLServer Then
            varNumero = Conta.GeraCodPagamento("67") 'Conta.Correlativo("TRIB", 67, "DECLARACAO DE ISSQN")
        ElseIf Bdados.Conexao.FormatoBanco = Access Or Bdados.Conexao.FormatoBanco Then
            varNumero = Conta.GeraCodPagamento("67")
        End If
    Else
        varNumero = "" & Rs!TDC_NUM_DECLARACAO
    End If
        
    Gravar = True
    If Len(Trim(varVersao)) = 0 Then varVersao = Nvl(Temp.PegaParametro(Bdados, "VERSAO DEC"), 2)
    Campos = "TDC_NUM_DECLARACAO,TDC_TCI_IM,TDC_PERIODO,TDC_TIPO_DEC,TDC_DATA_DECLARACAO,TDC_DATA_RECEPCAO,TDC_ORIGEM_DECLARACAO,TDC_VERSAO,TDC_STATUS,TDC_TUS_COD_USUARIO,TDC_TIP_COD_IMPOSTO"
    Valores = Bdados.PreparaValor(varNumero, Trim(varIM), varPeriodo, varTipo, varDeclaracao, varRecepcao, varOrigem, Bdados.Converte(varVersao, tctexto), Nvl(Trim(CStr(varStatus)), 1), AplicacoesVTFuncoes.Usuario, varCodImposto)
    Gravar = Gravar And Bdados.GravaDados("tab_declaracao_contribuinte", Valores, Campos, "TDC_NUM_DECLARACAO =" & varNumero)
    'GRAVA ITENS APURACAO
    If Gravar Then
        Bdados.DeletaDados "tab_detalhe_declaracao", "TDD_TDC_NUM_DECLARACAO=" & varNumero
        Campos = "TDD_TDC_NUM_DECLARACAO, TDD_TCD_COD_ITEM, TDD_NUM_DOC, TDD_VALOR_ITEM"
        
        For I = 1 To Me.Itens.Qtd
            Valores = Bdados.PreparaValor(varNumero, Me.Itens(I).Numero, Me.Itens(I).Documento, Bdados.Converte(Me.Itens(I).Valor, TCDuplo))
            Gravar = Gravar And Bdados.InsereDados("tab_detalhe_declaracao", Valores, Campos)
        Next
        
        Gravar = Gravar And Me.Notas.Gravar(varNumero)
'
'        Campos = "tdd_num_nota_inicial,tdd_num_nota_final,tdd_tgt_cod_pagamento"
'        Valores = Bdados.PreparaValor()
'        Bdados.GravaDados "TAB_DETALHE_DAM", Valores, Campos, "tdd_tgt_cod_pagamento =" & CodObrigacao
    End If
End Function

Public Function Finalizar(Optional ImpressaoDAM As Boolean = True, Optional stringTaxas As String, _
                        Optional TotalTaxas As Double, Optional Tipo As eTipoDeclaracao = decNormal, Optional ValorIss As Double)
    Dim Imposto As VSImposto
    Dim Obrigacao As Obrigacao
    Dim CodObrigacao As String, CodImposto As String, Sigla As String
    Dim Campos As String
    Dim Valores As String
    Set Obrigacao = New Obrigacao
    Set Imposto = New VSImposto
    Bdados.AtualizaDados "TAB_DECLARACAO_CONTRIBUINTE", Bdados.PreparaValor(decFinalizada), "TDC_STATUS", "TDC_NUM_DECLARACAO =" & varNumero
    
    CodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN))
    If Tipo = decNegativa Then
        CodObrigacao = Obrigacao.CriaObrigacao(CodImposto, Right(varPeriodo, 2) & Left(varPeriodo, 4), Right(varPeriodo, 2) & _
                    Left(varPeriodo, 4), varIM, ValorIss, etsSemMovimento, , , , , , , , , , , , etiContribuinte)
    Else
        CodObrigacao = Obrigacao.CriaObrigacao(CodImposto, Right(varPeriodo, 2) & Left(varPeriodo, 4), Right(varPeriodo, 2) & _
                    Left(varPeriodo, 4), varIM, ValorIss, , , , , , , , , , , , , etiContribuinte)
    End If
            
    If Val(CodObrigacao) <> 0 Then
        If ImpressaoDAM Then ImprimeDAMDeclaracao CodObrigacao
    End If
End Function
Public Function Salvar_Sem_Finalizar(Optional ImpressaoDAM As Boolean = True, Optional stringTaxas As String, _
        Optional TotalTaxas As Double, Optional Tipo As eTipoDeclaracao = decNormal, Optional ValorIss As Double, Optional Susbtitui As TipoSubstituicaoObrigacao = etsSubstitui)
    Dim Imposto As VSImposto
    Dim Obrigacao As Obrigacao
    Dim CodObrigacao As String, CodImposto As String, Sigla As String
     
    Set Obrigacao = New Obrigacao
    Set Imposto = New VSImposto
    CodImposto = varCodImposto
    If Tipo = decNegativa Then
        CodObrigacao = Obrigacao.CriaObrigacao(CodImposto, Right(varPeriodo, 2) & Left(varPeriodo, 4), _
            Right(varPeriodo, 2) & Left(varPeriodo, 4), varIM, ValorIss, etsSemMovimento, , , , , , , , _
            , varDeclaracao, , , etiContribuinte)
    Else
        CodObrigacao = Obrigacao.CriaObrigacao(CodImposto, Right(varPeriodo, 2) & Left(varPeriodo, 4), _
        Right(varPeriodo, 2) & Left(varPeriodo, 4), varIM, ValorIss, , Susbtitui, , , , , , , , , , , etiContribuinte)
    End If
    
    If Val(CodObrigacao) <> 0 Then
        If ImpressaoDAM Then ImprimeDAMDeclaracao CodObrigacao
    End If
End Function

Public Function CriaObrigacaoISS() As String
    Dim Imposto As VSImposto
    Dim Obrigacao As Obrigacao
    Dim CodObrigacao As String, CodImposto As String, Sigla As String

    Set Obrigacao = New Obrigacao
    Set Imposto = New VSImposto
    
    Sigla = Imposto.NomeTributo(ttr_ISSQN)
    CodImposto = Imposto.BuscaCodImposto(Sigla)
    CodObrigacao = Obrigacao.CriaObrigacao(CodImposto, Right(varPeriodo, 2) & Left(varPeriodo, 4), Right(varPeriodo, 2) & Left(varPeriodo, 4), varIM)

    If Val(CodObrigacao) <> 0 Then
        ImprimeDAMDeclaracao CodObrigacao
    End If
    CriaObrigacaoISS = CodObrigacao
End Function
