VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Arquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ArqHeader As New Header
Private ArqDetalhe() As New Detalhe
Private ArqTrailler As New Trailler
Private TextoGeral As String
Public LoteAberto As Double
Public ValorAcumuladoBaixado As Double
Private ArquivoJaLido As Boolean

Public Enum TipoArquivo
    etaArrecadacao = 1
    etaDeclaracao = 2
End Enum

Public Enum TipoPadrao
    etcRCB = 1
    etcCBR643 = 2
End Enum

Private NovasLinhas() As String

Public Function CarregaArquivos(TipoArquivo As TipoPadrao, Path As String, NumBanco As Object, Agencia As Object, DtGeracao As Object, _
                                TotalArrecadado As Object, DtPagamento As Object, TotalRegistro As Object, _
                                Barra As Object, Percentual As Object, ArqAceitos As Object, ArqRejeitados As Object, _
                                LoteCriado As Object, TotalBaixado As Object, TotalAberto As Object, Optional grdDocumentos As Object, Optional TotalDocumentos As Object, Optional SomenteCabecalho As Boolean = False) As String
    
    Dim Reg As Double
    Dim Arquivo As Integer
    Dim Documento As Object
    Dim NossoNumero As String
    Dim ContaCredito As String
    Dim Cbr724 As Boolean
    Dim CarteiraBancaria As Integer
    Dim Conta As New ContaCorrente
    
    On Error GoTo trata
    Cbr724 = False
    grdDocumentos.ListItems.Clear
    Close
    If Dir(Path) = "" Then
          Util.Avisa "Arquivo não encontrado." & vbCrLf & Path
    Else
        Reg = 0
        Arquivo = FreeFile(0)
        Close
        Open Path For Input As Arquivo
        Barra.Min = 0
        Barra.Visible = True
        Percentual.Visible = True
        TotalRegistro = 0
        TotalAberto = 0
        TotalBaixado = 0
        TotalDocumentos = 0
        ValorAcumuladoBaixado = 0
        TotalArrecadado = 0
        NumBanco = ""
        Do While Not EOF(Arquivo)
            Line Input #Arquivo, TextoGeral
            If TipoArquivo = etcRCB Then
                Select Case UCase(Left(TextoGeral, 1))
                    Case "A" 'Header de arquivo
                        With ArqHeader
                             NumBanco = RetornaTexto(TextoGeral, 43, 3) & " - " & RetornaTexto(TextoGeral, 46, 20)
                             DtGeracao = ConveteData(RetornaTexto(TextoGeral, 66, 8))
                             .Banco = Format(Left(NumBanco, 3), "000")
                             .DataGeracaoArquivo = DtGeracao
                        End With
                        If SomenteCabecalho Then
                            Close
                            Exit Function
                        End If
                    Case "G" 'Detalhes do arquivo(Documentos)
                        DtPagamento = ConveteData(RetornaTexto(TextoGeral, 22, 8))
                        ArqHeader.DataPagamentoDoc = DtPagamento
                        If Reg = 0 Then ReDim ArqDetalhe(1 To 1) As New Detalhe
                        Reg = Reg + 1
                        ReDim Preserve ArqDetalhe(1 To Reg) As New Detalhe
                        ArqDetalhe(Reg).Valor_Tarifa = CDbl(RetornaTexto(TextoGeral, 94, 6)) / 100
                        'Eu Raimundo Troquei o valor inicio de 42 para 82 seguindo o Lay-Out
                        ArqDetalhe(Reg).Valor_Recebido = CDbl(RetornaTexto(TextoGeral, 42, 11)) / 100
                        'ArqDetalhe(Reg).Valor_Recebido = CDbl(RetornaTexto(TextoGeral, 82, 10)) / 100
                        ArqDetalhe(Reg).AgenciaCredito = RetornaTexto(TextoGeral, 2, 5)
                        'Eu Raimundo Troquei  o valor de inicio para o correspondente do Lay-Out
                        ArqDetalhe(Reg).ContaCredito = RetornaTexto(TextoGeral, 7, 15)
                        'ArqDetalhe(Reg).ContaCredito = RetornaTexto(TextoGeral, 2, 20)
                        ArqDetalhe(Reg).AgenciaArrecadacao = RetornaTexto(TextoGeral, 109, 8)
                        'Raimundo Não existe esse campo no arquivo
                        If CDbl(RetornaTexto(TextoGeral, 58, 8)) = 0 Then
                            NossoNumero = CDbl(RetornaTexto(TextoGeral, 58, 24))
                        Else
                            NossoNumero = RetornaTexto(TextoGeral, 57, 25)
                        End If
                        ArqDetalhe(Reg).obNossoNumero = NossoNumero
                        ArqDetalhe(Reg).Versao_Barra = RetornaTexto(TextoGeral, 57, 1)
                        If Temp.PegaParametro(Bdados, "DAM SEM NUMERO") = "SIM" Then
                            If ArqDetalhe(Reg).Versao_Barra >= 3 And CDbl(RetornaTexto(TextoGeral, 58, 8)) = 0 Then 'Versao
                                ArqDetalhe(Reg).CodigoPagamento = NossoNumero
                            Else
                                If CInt(Mid(NossoNumero, 6, 1)) = 1 Then 'Tipo Inscricao
                                    'IM
                                    ArqDetalhe(Reg).obInscricao = Mid(NossoNumero, 7, 13)
                                    ArqDetalhe(Reg).obInscricao = Imposto.FormataInscricao(Left(ArqDetalhe(Reg).obInscricao, 10), InscContrib)
                                    'Periodo
                                    ArqDetalhe(Reg).obPeriodo = Mid(NossoNumero, 20, 6)
                                    If CDbl(Right(ArqDetalhe(Reg).obPeriodo, 2)) = 0 Then ArqDetalhe(Reg).obPeriodo = Left(ArqDetalhe(Reg).obPeriodo, 4)
                                Else
                                    'IC
                                    ArqDetalhe(Reg).obInscricao = Mid(NossoNumero, 7, 15)
                                    'Periodo
                                    ArqDetalhe(Reg).obPeriodo = Mid(NossoNumero, 22, 4)
                                End If
                                'CORRELATIVO
                                If CDbl(Mid(NossoNumero, 2, 2)) <> 11 Then 'BLOCO TEMPORARIO DEVIDO CORR. DO IPTU- TIRAR EM 30/06/2004
                                    ArqDetalhe(Reg).obCodTributo = Imposto.BuscaCodImposto("", Mid(NossoNumero, 2, 2))
                                Else
                                    ArqDetalhe(Reg).obCodTributo = Imposto.BuscaCodImposto("", "34")
                                End If
                                ArqDetalhe(Reg).obParcela = CDbl(Mid(NossoNumero, 4, 2))
                            End If
                        Else
                            ArqDetalhe(Reg).CodigoPagamento = CDbl(RetornaTexto(TextoGeral, 74, 8))
                            If AplicacoesVTFuncoes.municipio = "SANTA MARIA DA BOA VISTA" Then
                                If CDbl(ArqDetalhe(Reg).CodigoPagamento) < 100000 Then
                                    ArqDetalhe(Reg).CodigoPagamento = 81000000 + CDbl(ArqDetalhe(Reg).CodigoPagamento)
                                End If
                            End If
                        End If
                        If Temp.PegaParametro(Bdados, "BARRA COM VALOR") = "NAO" Then
                            ArqDetalhe(Reg).Valor_Recebido = CDbl(RetornaTexto(TextoGeral, 83, 11)) / 100
                        End If
                        Agencia = ArqDetalhe(Reg).AgenciaCredito
                        If Temp.PegaParametro(Bdados, "DAM SEM NUMERO") = "SIM" Then
                            Set Documento = grdDocumentos.ListItems.Add(, , ArqDetalhe(Reg).obNossoNumero)
                        Else
                            Set Documento = grdDocumentos.ListItems.Add(, , ArqDetalhe(Reg).CodigoPagamento)
                        End If
                        Documento.EnsureVisible
                        TotalRegistro = Reg
                        DoEvents
                    Case "Z" 'Trailler de  arquivo
                        TotalArrecadado = Format(CDbl(RetornaTexto(TextoGeral, 8, 18)) / 100, Const_Monetario)
                        Barra.Max = TotalRegistro
                        With ArqTrailler
                           .TotalDeArquivos = TotalRegistro
                           .TotalRemessa = TotalArrecadado
                        End With
                End Select
            ElseIf TipoArquivo = etcCBR643 Then
                If Not Temp.PegaParametro(Bdados, "TIPO TITULO") = "FLOAT" Then 'ANTIFLOAT
                    If Left(TextoGeral, 1) = "0" Then
                        If Mid(TextoGeral, 3, 9) = "RETORNO01" Then 'CARTEIRA 21
                              CarteiraBancaria = 21
                        Else
                            CarteiraBancaria = 1
                        End If
                    End If
                    If CarteiraBancaria = 21 Then
                        Select Case UCase(Left(TextoGeral, 1))
                            Case "0" 'Header de arquivo
                                With ArqHeader
                                     NumBanco = RetornaTexto(TextoGeral, 77, 3) & " - " & RetornaTexto(TextoGeral, 80, 15)
                                     DtGeracao = Left(RetornaTexto(TextoGeral, 95, 6), 2) & "/" & Mid(RetornaTexto(TextoGeral, 95, 6), 3, 2) & "/20" & Right(RetornaTexto(TextoGeral, 95, 6), 2)
                                     DtGeracao = DateAdd("d", 1, DtGeracao)
                                     .Banco = Left(NumBanco, 3)
                                     .DataGeracaoArquivo = DtGeracao
                                     Agencia = CDbl(RetornaTexto(TextoGeral, 27, 5))
                                     ContaCredito = CDbl(RetornaTexto(TextoGeral, 33, 8))
                                    '02RETORNO01COBRANCA       0562200013
                                End With
                            Case "1"
                                ArqHeader.DataPagamentoDoc = DateAdd("d", -1, DtGeracao)
                                DtPagamento = ArqHeader.DataPagamentoDoc
                                If Reg = 0 Then ReDim ArqDetalhe(1 To 1) As New Detalhe
                                Reg = Reg + 1
                                ReDim Preserve ArqDetalhe(1 To Reg) As New Detalhe
                                ArqDetalhe(Reg).AgenciaCredito = Agencia
                                ArqDetalhe(Reg).AgenciaArrecadacao = Agencia
                                ArqDetalhe(Reg).ContaCredito = ContaCredito
                                If Trim(RetornaTexto(TextoGeral, 46, 17)) <> "" Then 'somente para lagoa grande(temporario)
                                    ArqDetalhe(Reg).CodigoPagamento = Right(CDbl(RetornaTexto(TextoGeral, 46, 17)), 8)  'OK
                                Else 'somente para lagoa grande(temporario)
                                    If Not ArquivoJaLido Then
                                        ArqDetalhe(Reg).obCodTributo = Temp.PegaParametro(Bdados, "RECEITAS NAO CLASSIFICADAS")
                                        ArqDetalhe(Reg).CodigoPagamento = Conta.GeraCodPagamento(ArqDetalhe(Reg).obCodTributo)
                                        ArqDetalhe(Reg).obInscricao = Const_ImAvulso
                                        ArqDetalhe(Reg).obPeriodo = 2007
                                        ArqDetalhe(Reg).DataPagamento = DtPagamento
                                    End If
                                End If
                                If ArqDetalhe(Reg).CodigoPagamento <> "" Then
                                    ArqDetalhe(Reg).obNossoNumero = ArqDetalhe(Reg).CodigoPagamento
                                    ArqDetalhe(Reg).Valor_Recebido = CDbl(RetornaTexto(TextoGeral, 254, 13)) / 100  'OK
                                    Set Documento = grdDocumentos.ListItems.Add(, , ArqDetalhe(Reg).obNossoNumero)
                                    Documento.EnsureVisible
                                End If
                                TotalRegistro = Reg
                                TotalArrecadado = Format(CDbl(TotalArrecadado) + ArqDetalhe(Reg).Valor_Recebido, Const_Monetario)
                                
                                ArqTrailler.TotalDeArquivos = TotalRegistro
                                ArqTrailler.TotalRemessa = TotalArrecadado
                            Case "9" 'Trailler de  arquivo
                                If CInt(TotalRegistro) > 0 Then
                                    Barra.Max = TotalRegistro
                                Else
                                    TotalArrecadado = 0
                                End If
                                With ArqTrailler
                                   .TotalDeArquivos = TotalRegistro
                                   .TotalRemessa = TotalArrecadado
                            End With
                        End Select
                    Else 'CARTEIRA 01
                        Select Case UCase(Left(TextoGeral, 1))
                            Case "0" 'Header de arquivo
                                With ArqHeader
                                     NumBanco = RetornaTexto(TextoGeral, 77, 3) & " - " & RetornaTexto(TextoGeral, 80, 15)
                                     DtGeracao = Left(RetornaTexto(TextoGeral, 95, 6), 2) & "/" & Mid(RetornaTexto(TextoGeral, 95, 6), 3, 2) & "/20" & Right(RetornaTexto(TextoGeral, 95, 6), 2)
                                     DtGeracao = DateAdd("d", 1, DtGeracao)
                                     .Banco = Left(NumBanco, 3)
                                     .DataGeracaoArquivo = DtGeracao
                                End With
                            Case "1" 'Detalhes do arquivo(Documentos)
                                ArqHeader.DataPagamentoDoc = DateAdd("d", -1, DtGeracao)
                                DtPagamento = ArqHeader.DataPagamentoDoc
                                If Reg = 0 Then ReDim ArqDetalhe(1 To 1) As New Detalhe
                                Reg = Reg + 1
                                ReDim Preserve ArqDetalhe(1 To Reg) As New Detalhe
                                ArqDetalhe(Reg).CodigoPagamento = Right(CDbl(RetornaTexto(TextoGeral, 46, 17)), 8) 'OK
                                If Trim(ArqDetalhe(Reg).CodigoPagamento) <> "" Then
                                    ArqDetalhe(Reg).obNossoNumero = CDbl(RetornaTexto(TextoGeral, 46, 17))
                                    'RetornaTexto(TextoGeral, 182, 6)
                                    ArqDetalhe(Reg).Valor_Tarifa = CDbl(RetornaTexto(TextoGeral, 182, 7)) / 100 'OK
                                    'ArqDetalhe(Reg).AgenciaArrecadacao = RetornaTexto(TextoGeral, 147, 5) '0K data vencimento
                                    ArqDetalhe(Reg).AgenciaCredito = RetornaTexto(TextoGeral, 18, 5) '0K
                                    ArqDetalhe(Reg).AgenciaArrecadacao = ArqDetalhe(Reg).AgenciaCredito ' RetornaTexto(TextoGeral, 169, 4) '0K
                                    ArqDetalhe(Reg).BancoArrecadacao = CStr(CDbl(RetornaTexto(TextoGeral, 166, 3))) 'OK
                                    
                                    ArqDetalhe(Reg).Valor_Recebido = CDbl(RetornaTexto(TextoGeral, 254, 13)) / 100  'OK
                                    ArqDetalhe(Reg).ContaCredito = RetornaTexto(TextoGeral, 26, 6) 'OK
                                    Agencia = ArqDetalhe(Reg).AgenciaCredito '0K
            '                       ArqDetalhe(Reg).Versao_Barra = RetornaTexto(TextoGeral, 62, 1) 'OK
                                    Set Documento = grdDocumentos.ListItems.Add(, , ArqDetalhe(Reg).obNossoNumero)
                                    Documento.EnsureVisible
                                    TotalRegistro = Reg
                                    TotalArrecadado = Format(CDbl(TotalArrecadado) + ArqDetalhe(Reg).Valor_Recebido, Const_Monetario)
                                    
                                    ArqTrailler.TotalDeArquivos = TotalRegistro
                                    ArqTrailler.TotalRemessa = TotalArrecadado
                                    
                                End If
                                DoEvents
                                
                            Case "9" 'Trailler de  arquivo
                                If CInt(TotalRegistro) > 0 Then
                                    Barra.Max = TotalRegistro
                                Else
                                    TotalArrecadado = 0
                                End If
                                With ArqTrailler
                                   .TotalDeArquivos = TotalRegistro
                                   .TotalRemessa = TotalArrecadado
                                End With
                        End Select
                    End If
                Else 'TIPO ANTIFLOAT
                    If CInt(UCase(Left(TextoGeral, 2))) = 0 Then
                        If Mid(TextoGeral, 7, 6) = "CBR724" Then
                            Cbr724 = True
                        End If
                    End If
                    If Cbr724 Then
                        Select Case CInt(UCase(Left(TextoGeral, 2)))
                            Case 28
                                '28                                        9.172-3   18/019 SIMPLES       568-1 GRAJAU
                            If Trim(NumBanco) = "" Then
                               With ArqHeader
                                     NumBanco = Trim(Right(TextoGeral, 13))
                                     DtGeracao = Mid(TextoGeral, 115, 2) & "/" & Mid(TextoGeral, 117, 2) & "/" & Mid(TextoGeral, 119, 4)
                                     DtGeracao = DateAdd("d", 1, DtGeracao)
                                     .Banco = Right(NumBanco, 3)
                                     .DataGeracaoArquivo = DtGeracao
                                     Agencia = Edita.TiraTudo(Trim(Mid(TextoGeral, 70, 10)))
                                     ContaCredito = Edita.TiraTudo(Trim(Mid(TextoGeral, 40, 10)))
                                    '02RETORNO01COBRANCA       0562200013
                                End With
                            End If
                            Case 37, 7
                                ArqHeader.DataPagamentoDoc = DateAdd("d", -1, DtGeracao)
                                DtPagamento = ArqHeader.DataPagamentoDoc
                                If Reg = 0 Then ReDim ArqDetalhe(1 To 1) As New Detalhe
                                Reg = Reg + 1
                                ReDim Preserve ArqDetalhe(1 To Reg) As New Detalhe
                                ArqDetalhe(Reg).AgenciaCredito = Agencia
                                ArqDetalhe(Reg).AgenciaArrecadacao = Agencia
                                ArqDetalhe(Reg).ContaCredito = ContaCredito
                                ArqDetalhe(Reg).CodigoPagamento = Trim(Mid(TextoGeral, 13, 10))
                                ArqDetalhe(Reg).obNossoNumero = ArqDetalhe(Reg).CodigoPagamento
                                
                                ArqDetalhe(Reg).Valor_Recebido = Trim(Mid(TextoGeral, 87, 35))
                                Set Documento = grdDocumentos.ListItems.Add(, , ArqDetalhe(Reg).obNossoNumero)
                                Documento.EnsureVisible
                                TotalRegistro = Reg
                                TotalArrecadado = Format(CDbl(TotalArrecadado) + ArqDetalhe(Reg).Valor_Recebido, Const_Monetario)
                                
                                ArqTrailler.TotalDeArquivos = TotalRegistro
                                ArqTrailler.TotalRemessa = TotalArrecadado
                          Case 48
                            If CInt(TotalRegistro) > 0 Then
                                    Barra.Max = TotalRegistro
                                Else
                                    TotalArrecadado = 0
                                End If
                                With ArqTrailler
                                   .TotalDeArquivos = TotalRegistro
                                   .TotalRemessa = TotalArrecadado
                                End With
                        End Select
                    Else
                        Select Case UCase(Left(TextoGeral, 1))
                            Case "0" 'Header de arquivo
                                With ArqHeader
                                     NumBanco = RetornaTexto(TextoGeral, 77, 3) & " - " & RetornaTexto(TextoGeral, 80, 15)
                                     DtGeracao = Left(RetornaTexto(TextoGeral, 95, 6), 2) & "/" & Mid(RetornaTexto(TextoGeral, 95, 6), 3, 2) & "/20" & Right(RetornaTexto(TextoGeral, 95, 6), 2)
                                     DtGeracao = DateAdd("d", 1, DtGeracao)
                                     .Banco = Left(NumBanco, 3)
                                     .DataGeracaoArquivo = DtGeracao
                                     Agencia = CDbl(RetornaTexto(TextoGeral, 27, 5))
                                     ContaCredito = CDbl(RetornaTexto(TextoGeral, 33, 8))
                                    '02RETORNO01COBRANCA       0562200013
                                End With
                            Case "7"
                                ArqHeader.DataPagamentoDoc = DateAdd("d", -1, DtGeracao)
                                DtPagamento = ArqHeader.DataPagamentoDoc
                                If Reg = 0 Then ReDim ArqDetalhe(1 To 1) As New Detalhe
                                Reg = Reg + 1
                                ReDim Preserve ArqDetalhe(1 To Reg) As New Detalhe
                                ArqDetalhe(Reg).AgenciaCredito = Agencia
                                ArqDetalhe(Reg).AgenciaArrecadacao = Agencia
                                ArqDetalhe(Reg).ContaCredito = ContaCredito
                                ArqDetalhe(Reg).CodigoPagamento = Right(CDbl(RetornaTexto(TextoGeral, 73, 8)), 8)  'OK
                                ArqDetalhe(Reg).obNossoNumero = ArqDetalhe(Reg).CodigoPagamento
                                
                                ArqDetalhe(Reg).Valor_Recebido = CDbl(RetornaTexto(TextoGeral, 148, 18)) / 100  'OK
                                Set Documento = grdDocumentos.ListItems.Add(, , ArqDetalhe(Reg).obNossoNumero)
                                Documento.EnsureVisible
                                TotalRegistro = Reg
                                TotalArrecadado = Format(CDbl(TotalArrecadado) + ArqDetalhe(Reg).Valor_Recebido, Const_Monetario)
                                
                                ArqTrailler.TotalDeArquivos = TotalRegistro
                                ArqTrailler.TotalRemessa = TotalArrecadado
                            Case "9" 'Trailler de  arquivo
                                If CInt(TotalRegistro) > 0 Then
                                    Barra.Max = TotalRegistro
                                Else
                                    TotalArrecadado = 0
                                End If
                                With ArqTrailler
                                   .TotalDeArquivos = TotalRegistro
                                   .TotalRemessa = TotalArrecadado
                                End With
                            End Select
                        End If
                End If
            End If
        Loop
        If Reg > 0 Then
            GravaDocumentos ArqAceitos, ArqRejeitados, Barra, Percentual, LoteCriado, TotalBaixado, TotalAberto, TotalDocumentos, grdDocumentos, Dir(Path)
            DoEvents
        Else
            Avisa "Arquivo sem registro."
        End If
        Barra.Visible = False
        Percentual.Visible = False
    End If
    Exit Function
trata:
    If Err.Number <> 13 Then Avisa Err.Description
    If Err.Number <> 52 Then
        Resume Next
    End If
    Exit Function
    Resume
End Function

Public Function RetornaTexto(Texto As String, PosInicial As Integer, Tamanho As Integer) As String
    RetornaTexto = Mid$(Texto, PosInicial, Tamanho)
End Function
Public Function ConveteData(Data As String) As Date
        
    Dim Parametro(3)             As String
    Parametro(1) = Left(Data, 4) 'Ano
    Parametro(2) = Right(Data, 2) 'Dia
    Parametro(3) = Mid(Data, Len(Parametro(1)) + 1, 2) 'Mes
    ConveteData = Format(Parametro(2) & "/" & Parametro(3) & "/" & Parametro(1), "dd/mm/yyyy")

End Function
Public Function ChamaArquivos(Dialog As Object, Tipo As TipoArquivo) As String
        '<EhHeader>
        On Error GoTo ChamaArquivos_Err
        '</EhHeader>
        On Error GoTo trata
100     With Dialog
102         .CancelError = True
            If Tipo = etaArrecadacao Then
                .DialogTitle = "Receber arquivos do banco."
                .InitDir = Temp.PegaParametro(Bdados, "CAMINHO ARQUIVO RETORNO")
                .Filter = "Arquivos de remessa do banco|*.RET"
            Else
                .DialogTitle = "Receber arquivos de declaracão."
                .InitDir = "C:\"
                .Filter = "Arquivos de declaracão de contribuintes|*.DEC"
            End If
110         .ShowOpen
112         If .FileName <> "" Then
114             ChamaArquivos = .FileName
            End If
        End With
        Exit Function
trata:
        Exit Function
        '<EhFooter>
        Exit Function

ChamaArquivos_Err:
        
        Resume Next
        '</EhFooter>
End Function

Private Sub GravaDocumentos(ByRef ArquivosAceitos As Object, ByRef ArquivosRejeitados As Object, Barra As Object, _
                            Percentual As Object, LoteCriado As Object, TotalBaixado As Object, TotalAberto As Object, TotalDocumentos As Object, grdDocumentos As Object, Optional NomeArquivo As String)
                
        Dim Arrec As New Arrecadacao
        Dim Doc As Integer
        Dim Conta As String
        Dim Juros As Double
        Dim Correcao As Double
        Dim Multa As Double
        Dim ValorImpostoOriginal As Double
        Dim Seq As Integer, Sql As String, Rs As VSRecordset
        Dim Documento As Object, Situacao As String
        Dim Validacao As ResultadoValidacaoDAM
        Dim Parcela As Integer
        Dim ValorAberto As Double
        Dim Obrig As New Obrigacao
        Dim RsExt As VSRecordset
        Dim CodPagamentoExtrato As Double
        Dim ValorBaixadadoExtrato As Double
        Dim SequenciaDamExtrato As Integer
        Dim TxDam As Double
        Dim ValorExtrato As Double
        On Error GoTo trata
        '1. Gravar Lote
        '2.Valida Pagamento
        '3. Gravar DOC Pago
        ArquivosAceitos = 0
        TotalDocumentos = 0
        ArquivosRejeitados = 0
        Bdados.FechaTabela Rs
        For Doc = 1 To UBound(ArqDetalhe)
            '1.
            If Doc = 1 Then
                If Not AbreNovoLote(ArqDetalhe(Doc).AgenciaCredito, ArqDetalhe(Doc).ContaCredito) Then
                    If Not Confirma("Arquivo já recebido. Deseja prosseguir?") Then
                        Exit Sub
                    Else
                        ArquivoJaLido = True
                        DoEvents
                    End If
                Else
                    ArquivoJaLido = False
                End If
                LoteCriado = LoteAberto
                DoEvents
            End If
            
            '2.
            Arrec.CodImposto = ""
            Validacao = Arrec.ValidaDocumento(IIf(Trim(ArqDetalhe(Doc).CodigoPagamento) <> "", ArqDetalhe(Doc).CodigoPagamento, Nvl(ArqDetalhe(Doc).obNossoNumero, 0)), ArqDetalhe(Doc).Valor_Recebido, ArqHeader.DataPagamentoDoc, ArqHeader.Banco, ArqDetalhe(Doc).AgenciaArrecadacao, ValorImpostoOriginal, NomeArquivo, CStr(LoteAberto))
            If ArqDetalhe(Doc).obCodTributo <> "" And ArqDetalhe(Doc).obInscricao <> "" And ArqDetalhe(Doc).obPeriodo <> 0 Then
                Validacao = ervDocOK
                Arrec.TipoInscricao = 2
                Arrec.Inscricao = ArqDetalhe(Doc).obInscricao
                Arrec.CodImposto = ArqDetalhe(Doc).obCodTributo
                Arrec.Periodo = ArqDetalhe(Doc).obPeriodo
                Arrec.Parcela = 0
                Arrec.DAM = ArqDetalhe(Doc).CodigoPagamento
                Arrec.ValorOriginal = ArqDetalhe(Doc).Valor_Recebido
                Arrec.DtVencimento = ArqDetalhe(Doc).DataPagamento
            End If
            
            If Validacao = ervDocOK Then
                Sql = "Select max(TDR_SEQUENCIA_DAM_LOTE) +1 from tab_darm_recebido where " & _
                        " TDR_TLP_COD_LOTE=" & Me.LoteAberto
                If Bdados.AbreTabela(Sql, Rs) Then
                    Seq = Nvl("" & Rs(0), 1)
                Else
                    Seq = 1
                End If
                '3.
                If Arrec.CodImposto <> Const_Extrato And Arrec.CodImposto <> Const_Notificacao Then
                    If Len(Arrec.Periodo) = 4 Or Len(Arrec.Periodo) = 6 Then
                        CalculaAtualizacao Arrec.DAM, Multa, Juros, Correcao, ArqDetalhe(Doc).Valor_Recebido
                        If ArqDetalhe(Doc).Valor_Recebido > 0 Then
                            If Arrec.ValorOriginal = 0 Then Arrec.ValorOriginal = ArqDetalhe(Doc).Valor_Recebido
                            If Arrec.GravaPagamento(Arrec.Inscricao, Arrec.TipoInscricao, Arrec.CodImposto, IIf(Len(Arrec.Periodo) = 4, Arrec.Periodo, Right(Arrec.Periodo, 2) & Left(Arrec.Periodo, 4)), _
                                Arrec.DtVencimento, Arrec.DtPagamento, Arrec.ValorOriginal, _
                                0, 0, Juros, Multa, Correcao, _
                                Me.LoteAberto, CStr(Seq), Arrec.Parcela, Arrec.DAM, , ArqDetalhe(Doc).Valor_Recebido, , _
                                etsCreditoPago) Then
                                    ValorAcumuladoBaixado = ValorAcumuladoBaixado + ArqDetalhe(Doc).Valor_Recebido
                                    TotalBaixado = Format(ValorAcumuladoBaixado, Const_Monetario)
                                    ArquivosAceitos = ArquivosAceitos + 1
                                    TotalDocumentos = TotalDocumentos + 1
                            Else
                                Validacao = ervDocInexistente
                                ArquivosRejeitados = ArquivosRejeitados + 1
                                ValorAberto = ValorAberto + ArqDetalhe(Doc).Valor_Recebido
                                TotalAberto = Format(ValorAberto, Const_Monetario)
                                Arrec.LogOcorrencia ArqDetalhe(Doc).obNossoNumero, ArqDetalhe(Doc).Valor_Recebido, ArqHeader.DataPagamentoDoc, ervErro, ArqHeader.Banco, ArqDetalhe(Doc).AgenciaArrecadacao
                            End If
                        Else
                            Validacao = ervDocInexistente
                            ArquivosRejeitados = ArquivosRejeitados + 1
                            ValorAberto = ValorAberto + ArqDetalhe(Doc).Valor_Recebido
                            TotalAberto = Format(ValorAberto, Const_Monetario)
                            Arrec.LogOcorrencia Arrec.DAM, ArqDetalhe(Doc).Valor_Recebido, ArqHeader.DataPagamentoDoc, ervErro, ArqHeader.Banco, ArqDetalhe(Doc).AgenciaArrecadacao
                        End If
                    Else 'ERRO
                        Validacao = ervDocInexistente
                        ArquivosRejeitados = ArquivosRejeitados + 1
                        ValorAberto = ValorAberto + ArqDetalhe(Doc).Valor_Recebido
                        TotalAberto = Format(ValorAberto, Const_Monetario)
                        Arrec.LogOcorrencia ArqDetalhe(Doc).CodigoPagamento, ArqDetalhe(Doc).Valor_Recebido, ArqHeader.DataPagamentoDoc, ervErro, ArqHeader.Banco, ArqDetalhe(Doc).AgenciaArrecadacao
                    End If
                Else 'EXTRATO
                    Sql = "SELECT TPE_TGT_COD_PAGAMENTO,TPE_SUB_VALOR,TPE_TIPO_DOCUMENTO FROM TAB_PAGAMENTO_EXTRATO WHERE TPE_COD_PAGAMENTO_EXTRATO = " & Arrec.DAM
                    If Bdados.AbreTabela(Sql, Rs) Then
                        Rs.MoveFirst
                        ValorExtrato = 0
                        Do
                            ValorExtrato = ValorExtrato + Rs!TPE_SUB_VALOR
                            Rs.MoveNext
                        Loop While Not Rs.EOF
                        ValorExtrato = Format(ValorExtrato, Const_Monetario)
'                        ValorAcumuladoBaixado = ValorAcumuladoBaixado + ArqDetalhe(Doc).Valor_Recebido
                        TotalBaixado = Format(ValorAcumuladoBaixado, Const_Monetario)
                        ArquivosAceitos = ArquivosAceitos + 1
                        TxDam = 0
                        SequenciaDamExtrato = 0
                        ValorBaixadadoExtrato = 0
                        Rs.MoveFirst
                        DoEvents
                        Do
                            Bdados.DeletaDados "TAB_DARM_RECEBIDO", "TDR_TGT_COD_PAGAMENTO =" & Rs!TPE_TGT_COD_PAGAMENTO
                            Validacao = Arrec.ValidaDocumento(Rs!TPE_TGT_COD_PAGAMENTO, ArqDetalhe(Doc).Valor_Recebido, ArqHeader.DataPagamentoDoc, ArqHeader.Banco, ArqDetalhe(Doc).AgenciaArrecadacao, ValorImpostoOriginal, NomeArquivo, CStr(LoteAberto))
                            CodPagamentoExtrato = Rs!TPE_TGT_COD_PAGAMENTO
                            If Validacao = ervDocOK Then
                                Sql = "Select max(TDR_SEQUENCIA_DAM_LOTE) +1 from tab_darm_recebido where " & _
                                        " TDR_TLP_COD_LOTE=" & Me.LoteAberto
                                If Bdados.AbreTabela(Sql, RsExt) Then
                                    Seq = Nvl("" & RsExt(0), 1)
                                Else
                                    Seq = 1
                                End If
                                Bdados.FechaTabela RsExt
                                '********************************
                                CalculaAtualizacao CStr(CodPagamentoExtrato), Multa, Juros, Correcao, Rs!TPE_SUB_VALOR
                                If Rs!TPE_SUB_VALOR > 0 Then
                                    If SequenciaDamExtrato = 0 Then
                                        If Rs!TPE_TIPO_DOCUMENTO = 1 And (ValorExtrato < ArqDetalhe(Doc).Valor_Recebido) Then
                                            TxDam = TrocaPic(Nvl(Temp.PegaParametro(Bdados, "TXDAM"), 0), ".", ",")
                                        Else
                                            TxDam = 0
                                        End If
                                    End If
                                    If Arrec.GravaPagamento(Arrec.Inscricao, Arrec.TipoInscricao, Arrec.CodImposto, IIf(Len(Arrec.Periodo) = 4, Arrec.Periodo, Right(Arrec.Periodo, 2) & Left(Arrec.Periodo, 4)), _
                                        Arrec.DtVencimento, Arrec.DtPagamento, Arrec.ValorOriginal, _
                                        0, 0, Juros, Correcao, Multa, _
                                        Me.LoteAberto, CStr(Seq), Arrec.Parcela, Arrec.DAM, , Rs!TPE_SUB_VALOR + IIf(SequenciaDamExtrato = 0, TxDam, 0), , _
                                        etsCreditoPago) Then
                                            ValorBaixadadoExtrato = ValorBaixadadoExtrato + Rs!TPE_SUB_VALOR + IIf(SequenciaDamExtrato = 0, TxDam, 0)
                                            ValorAcumuladoBaixado = ValorAcumuladoBaixado + Rs!TPE_SUB_VALOR + IIf(SequenciaDamExtrato = 0, TxDam, 0)
                                            SequenciaDamExtrato = SequenciaDamExtrato + 1
                                            TotalDocumentos = TotalDocumentos + 1
                                    Else
                                        ValorAberto = ValorAberto + Rs!TPE_SUB_VALOR + IIf(SequenciaDamExtrato = 0, TxDam, 0)
                                        TotalAberto = Format(ValorAberto, Const_Monetario)
                                    End If
                                    DoEvents
                                Else
                                    ValorAberto = ValorAberto + Rs!TPE_SUB_VALOR + IIf(SequenciaDamExtrato = 0, TxDam, 0)
                                    TotalAberto = Format(ValorAberto, Const_Monetario)
                                End If
                                
                            '********************************
                            Else
                                ValorBaixadadoExtrato = 0
                            End If
                            Rs.MoveNext
                        Loop While Not Rs.EOF
                        Arrec.CodImposto = Const_Extrato
                        If Format(ValorBaixadadoExtrato, Const_Monetario) < ArqDetalhe(Doc).Valor_Recebido Then
                            ValorAberto = ValorAberto + Format((ArqDetalhe(Doc).Valor_Recebido - ValorBaixadadoExtrato), Const_Monetario)
                            Validacao = 7
                        ElseIf Format(ValorBaixadadoExtrato, Const_Monetario) > ArqDetalhe(Doc).Valor_Recebido Then
                            Validacao = 7
                        End If
                    Else
                        Arrec.LogOcorrencia ArqDetalhe(Doc).obNossoNumero, ArqDetalhe(Doc).Valor_Recebido, ArqHeader.DataPagamentoDoc, ervErro, ArqHeader.Banco, ArqDetalhe(Doc).AgenciaArrecadacao
                        ArquivosRejeitados = ArquivosRejeitados + 1
                        ValorAberto = ValorAberto + ArqDetalhe(Doc).Valor_Recebido
                        TotalAberto = Format(ValorAberto, Const_Monetario)
                    End If
                End If
            
            Else
                If Validacao = ervDocPago Then
                    Bdados.AtualizaDados "TAB_DARM_RECEBIDO", Bdados.PreparaValor(LoteCriado, Bdados.Converte(ArqHeader.DataPagamentoDoc, TCDataHora)), "TDR_TLP_COD_LOTE,TDR_DATA_PAGAMENTO", "TDR_TGT_COD_PAGAMENTO =" & ArqDetalhe(Doc).CodigoPagamento
                End If
                ArquivosRejeitados = ArquivosRejeitados + 1
                ValorAberto = ValorAberto + ArqDetalhe(Doc).Valor_Recebido
                TotalAberto = Format(ValorAberto, Const_Monetario)
            End If

            Situacao = ""
            If Arrec.CodImposto = Const_Extrato And Validacao <> 7 Then
                Situacao = "OK - EXTRATO"
            ElseIf Validacao = ervDocInexistente Then
                Situacao = "NAO ENCONTRADO"
            ElseIf Validacao = ervDocPago Then
                Situacao = "JA PAGO"
            ElseIf Validacao = ervDocOK Then
                Situacao = "OK"
            ElseIf Validacao = 7 Then
                Situacao = "EXTRATO INCOMPLETO"
            Else
                Situacao = "ERRO DESCONHECIDO"
            End If
            If Trim(ArqDetalhe(Doc).CodigoPagamento) <> "" Then
                Set Documento = grdDocumentos.FindItem(ArqDetalhe(Doc).CodigoPagamento)
            Else
                Set Documento = grdDocumentos.FindItem(ArqDetalhe(Doc).obNossoNumero)
            End If
            If Not Documento Is Nothing Then
            
                Documento.EnsureVisible
                If Trim(ArqDetalhe(Doc).CodigoPagamento) <> "" Then
                    Documento.SubItems(1) = ArqDetalhe(Doc).obInscricao
                    Documento.SubItems(2) = ArqDetalhe(Doc).obCodTributo
                    Documento.SubItems(3) = ArqDetalhe(Doc).obPeriodo
                    Documento.SubItems(4) = Format(ArqDetalhe(Doc).Valor_Recebido, Const_Monetario)
                Else
                    
                    Documento.SubItems(1) = Arrec.Inscricao
                    Documento.SubItems(2) = Arrec.CodImposto
                    Documento.SubItems(3) = Arrec.Periodo
                    Documento.SubItems(4) = Format(ArqDetalhe(Doc).Valor_Recebido, Const_Monetario)
                    Documento.SubItems(6) = Arrec.DAM
                End If
                Documento.SubItems(5) = Situacao
                    
                DoEvents
            End If
            Barra.Value = Doc
            Percentual = Val(100 * (Barra.Value / Barra.Max)) & " % "
            DoEvents
        Next
        If TotalAberto = 0 Then Arrec.FechaLote LoteAberto
        Exit Sub
trata:
    If Err.Number <> 13 Then Avisa Err.Description
    Resume Next
    Exit Sub
    Resume
End Sub

Private Function AbreNovoLote(Agencia As String, Conta As String, Optional ByRef CodLote As Double) As Boolean
        Dim Sql As String
        Dim Rs As VSRecordset
        Dim Arrec As New Arrecadacao
        Sql = "SELECT TLP_COD_LOTE FROM TAB_LOTE_PAGAMENTO WHERE TLP_TAR_COD_AGENTE =" & ArqHeader.Banco & _
                " AND TLP_NUM_SUCURSAL ='" & Agencia & "' AND TLP_NUM_CONTA =" & Conta & _
                " AND TLP_DATA_ARRECADACAO = " & Bdados.FormataValorCampo(ArqHeader.DataPagamentoDoc)
                
        Sql = "SELECT TLP_COD_LOTE FROM TAB_LOTE_PAGAMENTO WHERE TLP_DATA_ARRECADACAO = " & _
            Bdados.FormataValorCampo(ArqHeader.DataPagamentoDoc)
            
        Sql = "SELECT TLP_COD_LOTE FROM TAB_LOTE_PAGAMENTO WHERE TLP_TAR_COD_AGENTE =" & ArqHeader.Banco & _
                " AND TLP_DATA_ARRECADACAO = " & Bdados.FormataValorCampo(ArqHeader.DataPagamentoDoc)
                
        If Bdados.AbreTabela(Sql, Rs) Then
            CodLote = Rs!TLP_COD_LOTE
            LoteAberto = Rs!TLP_COD_LOTE
            AbreNovoLote = False
            Bdados.AtualizaDados "TAB_LOTE_PAGAMENTO", Bdados.PreparaValor(0, Bdados.Converte(ArqTrailler.TotalRemessa, TCDuplo)), "TLP_SITUACAO_LOTE,TLP_VALOR_ARRECADADO", "TLP_COD_LOTE = '" & CodLote & "'"
        Else
            CodLote = Arrec.AbreLote(ArqHeader.Banco, Agencia, Conta, ArqTrailler.TotalRemessa, _
                            ArqHeader.DataPagamentoDoc, Date)
            LoteAberto = CodLote
            AbreNovoLote = True
        End If
        
End Function

Private Sub CalculaAtualizacao(CodPagamento As String, Optional Multas As Double, Optional Juros As Double, _
                            Optional Correcao As Double, Optional ValorPago As Double)
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Conta As ContaCorrente
    Dim Mora As Double
    
    If CodPagamento = 0 Then
        Juros = 0
        Multas = 0
        Mora = 0
        Exit Sub
    End If
    Set Conta = New ContaCorrente
    Set Rs = Conta.BuscaDam(CodPagamento)
    If Rs.EOF Then
        Exit Sub
    End If
    
    Correcao = Format(Conta.CalculaValoresCorrecaoAvulso(Rs!Imposto, Rs!Periodo, Rs!Vencimento, ArqHeader.DataPagamentoDoc, Rs!ValorTributo), Const_Monetario)
    Juros = Format(Conta.CalculaValoresJurosAvulsos(Rs!Imposto, Rs!Periodo, 0, ArqHeader.DataPagamentoDoc, Rs!Vencimento, Rs!ValorTributo + Correcao), Const_Monetario)
    Multas = Format(Conta.CalculaValoresMultaAvulsos(Rs!Imposto, Rs!Periodo, 0, ArqHeader.DataPagamentoDoc, Rs!Vencimento, Rs!ValorTributo + Correcao), Const_Monetario)
        
    Mora = Format(ValorPago - Rs!ValorTributo, Const_Monetario)
'    If Mora <> 0 Then
'        If Mora > Juros + Multas Then
'            'Caso tenha sido pago mais do que devido, calculo a
'            'proporcionalidade de multa/juros
'            Conta.CalculaTributoProporcional Multas, Juros, Multas + Juros, Juros, Mora
'        Else
'            Exit Sub
'        End If
'    Else
'        Juros = 0
'        Multas = 0
'    End If
End Sub
Public Function PesquisaDocumentos(Lista As Object, Banco As String, Agencia As String, Data As String, Ocorrencia As String) As Boolean
    Dim Sql As String
    Sql = "Select TLA_TGT_COD_PAGAMENTO as DOC, to_number(TLA_VALOR,'999999.99') as Valor," & _
        "TLA_DATA_ARRECADACAO as DATA_ARREC,TAR_NOME_AGENTE as BANCO," & _
        "TLA_AGENCIA as AGENCIA,OCORRENCIA.TGE_NOME AS TIPO from TAB_LOG_OCORRENCIA, TAB_AGENTE_ARRECADADOR," & _
            "(SELECT * FROM TAB_GERAL WHERE TGE_TIPO = (SELECT TGE_TIPO FROM " & _
            "TAB_GERAL WHERE TGE_NOME = 'REJEICAO DOCUMENTO' AND TGE_TIPO <> 0)) AS OCORRENCIA where " & _
        " TLA_tar_cod_agente = tar_cod_agente AND TLA_OCORRENCIA = OCORRENCIA.TGE_CODIGO"
        
    If Trim(Banco) <> "" Then Sql = Sql & " and TLA_TAR_COD_AGENTE=" & Banco
    If Trim(Agencia) <> "" Then Sql = Sql & " and TLA_AGENCIA ='" & Agencia & "'"
    
    If Trim(Data) <> "" Then Sql = Sql & " and TLA_DATA_ARRECADACAO= " & Bdados.Converte(Data, TCDataHora)
    If Trim(Ocorrencia) <> "" Then Sql = Sql & " and TLA_OCORRENCIA =" & Ocorrencia
    PesquisaDocumentos = Lista.Preencher(Bdados, Sql, 1000, 1000, 1300, 2300, 1000, 600)
End Function

Public Function CorrigeArquivo(Caminho As String) As Boolean
    Dim Reg As Double
    Dim Arquivo As Integer
    Dim Linha As String
    Dim TotalLinhas As Double
    Dim i As Double
    Dim PosInicioCalc As Double
    Dim TamArquivo As Double
    
    CorrigeArquivo = False
    If Dir(Caminho) = "" Then
          MsgBox "Arquivo não encontrado." & vbCrLf & Caminho
          Exit Function
    Else
        Reg = 0
        TotalLinhas = 1
        ReDim NovasLinhas(1 To 1) As String
        Arquivo = FreeFile(0)
        Caminho = Trim(Caminho)
        Close
        Open Caminho For Input As Arquivo
        Line Input #Arquivo, Linha
        NovasLinhas(1) = Left(Linha, 150)
        TamArquivo = Len(Linha)
        PosInicioCalc = 152
        Do While PosInicioCalc < TamArquivo
            TotalLinhas = TotalLinhas + 1
            ReDim Preserve NovasLinhas(1 To TotalLinhas) As String
            NovasLinhas(TotalLinhas) = Trim$(Mid(Linha, PosInicioCalc, 150))
            PosInicioCalc = PosInicioCalc + 151
        Loop
    End If
    Close
    Open Left(Caminho, Len(Caminho) - 4) & "_Corrigido" & Right(Caminho, 4) For Output As Arquivo
    For i = 1 To TotalLinhas - 1
        Print #Arquivo, NovasLinhas(i)
    Next
    Close
    CorrigeArquivo = True
End Function
