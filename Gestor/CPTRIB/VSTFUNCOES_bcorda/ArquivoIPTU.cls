VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArquivoIPTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Private FebrArq As New VTFebraban.VTArquivo
'Private FebrHeaderArq As New VTFebraban.VTHeaderArquivo
'Private FebrHeaderLote As New VTFebraban.VTHeaderLote
'Private FebrLote As New VTFebraban.VTLote
'Private FebrDet As New VTFebraban.VTDetalhe
'Private FebrTreilerLote As New VTFebraban.VTTrailerLote
'Private FebrTreilerArq As New VTFebraban.VTTrailerArquivo

Public CodFebraban As String
Public ContaCorrente As String
Public ContaDV As String
Public Agencia As String
Public AgenciaDV As String
Public BancoCod As Integer
Public BancoRazao As String
Public CodRemessa As String

Public Sub InicializaClasse(sFebraban As String, sContaCorrente As String, sContaDV As String, sAgencia As String, sAgenciaDV As String, iBancoCod As Integer, sBancoRazao As String)
    CodFebraban = sFebraban
    ContaCorrente = sContaCorrente
    ContaDV = UCase(sContaDV)
    Agencia = sAgencia
    AgenciaDV = UCase(sAgenciaDV)
    BancoCod = iBancoCod
    BancoRazao = UCase(sBancoRazao)
End Sub

Public Function GerarArquivo() As Boolean
'    Dim i As Integer, j As Integer, intPosicao As Integer
'    Dim Arq As New VTFebraban.VTArquivo
'    Dim Lot As VTFebraban.VTLote
'    Dim Det_A As VTFebraban.VTDetalhe
'    Dim Det_B As VTFebraban.VTDetalheB
'    Dim Lote As VTFebraban.VTLote
'    Dim Valores As String, Campos As String
'    Dim Valor As Double
'    Dim CaminhoArquivo  As String
'    Dim Sql As String
'    Dim Rs As VSRecordset
'    '1 : Preenchendo o header do arquivo...
'    With Arq.Header
'        .NomeEmpresa = Temp.PegaParametro(Bdados, "CLIENTE")
'        .TipoInscricao = inscCNPJ
'        .NumeroInscricao = Temp.PegaParametro(Bdados, "CGC CLIENTE")
'
'        .Convenio = CodFebraban
'        .Banco = BancoCod
'        .NomeBanco = BancoRazao
'        .Agencia = Agencia
'        .DvAgencia = AgenciaDV
'        .Conta = ContaCorrente
'        .DvConta = ContaDV
'
'        .CodigoArquivo = arqRemessa
'        .NSA = 1
'
'        CaminhoArquivo = App.Path & "\REMESSAS\" & Format(.NSA, "000000") & ".rem"
'
'    End With
'    '2 : Inclusao de Detalhe segmento A
'    Set Lot = Arq.Lotes.NovoLote() 'Tirei a criacao de novo lote e botei so 1 lote aqui...
'    Sql = "Select * from VIS_ARQUIVO_IPTU"
'    If Bdados.AbreTabela(Sql, Rs) Then
'        Rs.MoveFirst
'        Do
'            '3 : Preenchimento dos lotes do arquivo
'            'O criar novo Lot estava aqui...
'            With Lot.Header
'                .Banco = BancoCod
'                .Operacao = Arq.CodigoOperacao(operArquivoRemessa)
'                .Servico = servCobranca
'                .FormaLancamento = lanctoCreditoConta
'
'                .TipoInscricao = inscCNPJ
'                .NumeroInscricao = Temp.PegaParametro(Bdados, "CGC CLIENTE")
'
'                .Convenio = CodFebraban
'                .Agencia = Agencia
'                .DvAgencia = AgenciaDV
'                .Conta = ContaCorrente
'                .DvConta = ContaDV
'
'                .Empresa = Temp.PegaParametro(Bdados, "CLIENTE")
'                .Logradouro = Temp.PegaParametro(Bdados, "LOGRADOURO")
'                .NumeroLogradouro = Temp.PegaParametro(Bdados, "NUMERO LOGRADOURO")
'                .ComplementoLogradouro = Temp.PegaParametro(Bdados, "COMPLEMENTO LOGRADOURO")
'                .Cidade = Temp.PegaParametro(Bdados, "CIDADE")
'                .CEP = Left(Trim(Temp.PegaParametro(Bdados, "CEP MUNICIPIO")), 5)
'                .ComplementoCEP = Right(Trim(Temp.PegaParametro(Bdados, "CEP MUNICIPIO")), 3)
'                .Estado = Temp.PegaParametro(Bdados, "ESTADO")
'            End With
'            With Lot.Trailer
'                .Banco = BancoCod
'            End With
'            Set Det_A = Lot.Detalhes.NovoDetalhe
'            With Det_A
'                .Banco = BancoCod
'                .TipoMovimento = tmovInclusao
'                .CodigoMovimento = cmovInclusaoDetalheLiberado
'                .Compensacao = ""
'
'                .BancoFavorecido = BancoCod
'                .AgenciaFavorecido = Agencia
'                .DvAgenciaFavorecido = AgenciaDV
'                .ContaFavorecido = ContaCorrente
'                .DvContaFavorecido = ContaDV
'
'                .NomeFavorecido = Rs!tci_nome
'
'                .SeuNumero = Rs!tgt_cod_pagamento
'                .DataLancamento = Format(Date, "ddmmyyyy")
'                .TipoMoeda = Arq.CodigoMoeda(moedReal)
'
'                .QuantidadeMoeda = 1
'                .ValorLancamento = Format(Rs!TGT_VALOR_TRIBUTO + Rs!TGT_TAXA_EXPEDIENTE, Const_Monetario)
'            End With
''            Set Det_B = Lot.Detalhes.NovoDetalhe
'            With Det_B
''                .Banco = BancoCod
''                .TipoInscricao = inscCPF
''
''                .NumeroInscricao = Rs!TIM_IM
''
''                .Logradouro = Rs!TCI_LOGRADOURO & " " & Rs!TCI_NOME_LOGRADOURO
''                .ComplementoLogradouro = Rs!TCI_COMPLEMENTO
''                .NumeroLogradouro = Rs!TCI_NUMERO
''                .Bairro = Rs!TCI_BAIRRO
''                .CEP = Left(Trim(Rs!TCI_CEP), 5)
''                .ComplementoCEP = Right(Trim(Rs!TCI_CEP), 3)
''                .Cidade = Rs!TCI_CIDADE
''                .Estado = Rs!TCI_UF
''
''                .Valor = Format(Rs!TGT_VALOR_TRIBUTO + Rs!TGT_TAXA_EXPEDIENTE, Const_Monetario)
''                .Vencimento = Rs!TGT_DATA_VENCIMENTO
'            End With
'            Rs.MoveNext
'        Loop While Not Rs.EOF
'    End If
'    '4 : Fechamento do arquivo
'    With Arq.Trailer
'        .Banco = BancoCod
'    End With
'    GerarArquivo = Arq.Gravar(CaminhoArquivo)
'    Set Arq = Nothing
End Function
