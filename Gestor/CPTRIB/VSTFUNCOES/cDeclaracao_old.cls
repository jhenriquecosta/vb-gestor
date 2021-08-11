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

Private Const cteItemImpostoDevido As Byte = 7

Public Enum eTipoDeclaracao
    decNormal = 1
    decSubstitutiva = 2
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
Private varVersao As String

Public tciNome As String
Public tciEndereco As String
Public tciBairro As String
Public tciCEP As String
Public tciCidade As String
Public tciUF As String
Public tciAtividade As String


Public Property Get Numero() As Integer
    Numero = varNumero
End Property


Public Property Let IM(value As String)
    varIM = value
End Property
Public Property Get IM() As String
    IM = varIM
End Property


Public Property Let Periodo(value As String)
    varPeriodo = value
End Property
Public Property Get Periodo() As String
    Periodo = varPeriodo
End Property


Public Property Let Tipo(value As eTipoDeclaracao)
    varTipo = value
End Property
Public Property Get Tipo() As eTipoDeclaracao
    Tipo = varTipo
End Property


Public Property Let Data(value As Date)
    varDeclaracao = value
End Property
Public Property Get Data() As Date
    Data = varDeclaracao
End Property


Public Property Let Recepcao(value As Date)
    varRecepcao = value
End Property
Public Property Get Recepcao() As Date
    Recepcao = varRecepcao
End Property


Public Property Let Origem(value As eOrigemDeclaracao)
    varOrigem = value
End Property
Public Property Get Origem() As eOrigemDeclaracao
    Origem = varOrigem
End Property


Public Property Get Itens() As cItensDeclaracao
    Set Itens = varItens
End Property


Public Property Let Versao(value As String)
    varVersao = value
End Property
Public Property Get Versao() As String
    Versao = varVersao
End Property


Private Sub Class_Initialize()
    Set varItens = New cItensDeclaracao
End Sub
Private Sub Class_Terminate()
    Set varItens = Nothing
End Sub

Public Function Buscar(IM As String, Periodo As String, Optional FormatoPeriodo As eFormatoPeriodo = perAAAMM) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    
    If IM = "" Or Periodo = "" Then Exit Function
    If Len(Periodo) <> 4 Then
        If FormatoPeriodo = perMMAAAA Then Periodo = Right(Periodo, 4) & Left(Periodo, 2)
    End If
    Sql = "SELECT * " & _
            " FROM TAB_DECLARACAO_CONTRIBUINTE" & _
            " WHERE TDC_TCI_IM ='" & IM & "'" & _
                " AND TDC_PERIODO =" & Periodo & _
                " AND TDC_TIPO_DEC=(SELECT MAX(TDC_TIPO_DEC)" & _
                                            " FROM TAB_DECLARACAO_CONTRIBUINTE" & _
                                            " WHERE TDC_TCI_IM ='" & IM & "'" & _
                                                " AND TDC_PERIODO =" & Periodo & ")"
    If Bdados.AbreTabela(Sql, Rs) Then
        varNumero = Nvl("" & Rs!TDC_NUM_DECLARACAO, 0)
        varIM = "" & Rs!TDC_TCI_IM
        varPeriodo = Nvl("" & Rs!TDC_PERIODO, 0)
        varTipo = Nvl("" & Rs!TDC_TIPO_DEC, 0)
        varDeclaracao = "" & Rs!TDC_DATA_DECLARACAO
        varRecepcao = "" & Rs!TDC_DATA_RECEPCAO
        varOrigem = Nvl("" & Rs!TDC_ORIGEM_DECLARACAO, 0)
        varVersao = "" & Rs!TDC_VERSAO
        
        preencherItens
        Buscar = True
    End If
    Bdados.FechaTabela Rs
End Function

Private Sub preencherItens()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Item As cItemDeclaracao
    Dim Documento As String
    
    Sql = "SELECT *" & _
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

Public Function Gravar(Optional FormatoPeriodo As eFormatoPeriodo = perAAAMM) As Boolean
    Dim Valores As String, Campos As String
    Dim I As Integer
    Dim Imposto As VSImposto
    Dim Obrigacao As Obrigacao
    Dim Cobranca As VSCobranca
    Dim CodObrigacao As String, CodImposto As String, Sigla As String
    
    If Len(varPeriodo) <> 4 Then
        If FormatoPeriodo = perMMAAAA Then varPeriodo = Right(varPeriodo, 4) & Left(varPeriodo, 2)
    End If
    
    Dim Sql As String, Rs As VSRecordset
    Sql = "SELECT *" & _
            " FROM TAB_DECLARACAO_CONTRIBUINTE" & _
            " WHERE TDC_TCI_IM ='" & varIM & "'" & _
                " AND TDC_PERIODO =" & varPeriodo & _
                " AND TDC_TIPO_DEC=2"
    If Bdados.AbreTabela(Sql, Rs) Then
        Erro "Contribuinte possui uma substitutiva no período."
        Exit Function
    End If
    Bdados.FechaTabela Rs
    
    varNumero = Bdados.Correlativo("TRIB", 67, "DECLARACAO DE ISSQN")
    
    Gravar = True
    Campos = "TDC_NUM_DECLARACAO,TDC_TCI_IM,TDC_PERIODO,TDC_TIPO_DEC,TDC_DATA_DECLARACAO,TDC_DATA_RECEPCAO,TDC_ORIGEM_DECLARACAO,TDC_VERSAO"
    Valores = Bdados.PreparaValor(varNumero, varIM, varPeriodo, varTipo, varDeclaracao, varRecepcao, varOrigem, Bdados.Converte(varVersao, tctexto))
    Gravar = Gravar And Bdados.InsereDados("tab_declaracao_contribuinte", Valores, Campos)
    Gravar = Gravar And Me.Itens.Gravar(varNumero)
    
    Set Obrigacao = New Obrigacao
    Set Imposto = New VSImposto
    If Len(varPeriodo) <> 4 Then
        If FormatoPeriodo = perMMAAAA Then varPeriodo = Right(varPeriodo, 2) & Left(varPeriodo, 4)
    End If
    Sigla = Imposto.NomeTributo(ttr_ISSQN)
    CodImposto = Imposto.BuscaCodImposto(Sigla)
    CodObrigacao = Obrigacao.CriaObrigacao(CodImposto, varPeriodo, varPeriodo, varIM)
    Obrigacao.BuscaDetalheObrigacao CodObrigacao
    If Util.Confirma("Imprimir DAM?") Then
        Set Cobranca = New VSCobranca
        Dim NomeTributo As String
        Dim NovaData As String
        Dim NovaMulta As Double
        Dim NovoJuro As Double
        
        NovaData = Imposto.DataVencimentoNova(Format(Obrigacao.obDataVencimento, "dd/mm/yyyy"))
        If Trim(NovaData) = "" Then Exit Function
        NomeTributo = Imposto.BuscaNomeCompletoImposto(Obrigacao.obCodImposto)
        NovaMulta = Conta.CalculaValoresMultaAvulsos(Obrigacao.obCodImposto, Obrigacao.obPeriodo, EtcCreditoTributario, _
                    NovaData, Obrigacao.obDataVencimento, Obrigacao.obValorObrigacao)
        NovoJuro = Conta.CalculaValoresJurosAvulsos(Obrigacao.obCodImposto, Obrigacao.obPeriodo, EtcCreditoTributario, _
                    NovaData, Obrigacao.obDataVencimento, Obrigacao.obValorObrigacao)
        
        Dim iNI As String
        Dim fIN As String
        If CodImposto = Imposto.BuscaCodImposto("ISSQN") Then
            If Bdados.AbreTabela("SELECT TDD_VALOR_ITEM FROM TAB_DETALHE_DECLARACAO WHERE TDD_TDC_NUM_DECLARACAO=" & varNumero & " AND TDD_TCD_COD_ITEM IN (1,2) ORDER BY TDD_TCD_COD_ITEM ASC") Then
                Bdados.Tabela.MoveFirst
                iNI = Format(Bdados.Tabela(0) & "", "000")
                Bdados.Tabela.MoveNext
                fIN = Format(Bdados.Tabela(0) & "", "000")
                Bdados.FechaTabela
            End If
        Else
            iNI = ""
            fIN = ""
        End If
        
        Cobranca.ImprimeDam Rpt, Obrigacao.obCodigoObrigacao, Obrigacao.obContribuinte, tciNome, "", tciEndereco, _
             "", "", Obrigacao.obCodImposto, _
            Sigla, NomeTributo, Me.Periodo, "0", 1, NovaData, Me.Itens(7).Valor, _
            Obrigacao.obValorObrigacao, CStr(NovaMulta), CStr(NovoJuro), 0, 0, 0, _
            tciAtividade, , iNI, fIN, , , , , , , , , tdiTela
    End If
End Function
