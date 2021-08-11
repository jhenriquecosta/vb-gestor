VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private NumRegPorFolha As Integer
Private NumFolhaPorLivro As Integer
Public Enum TipoDocumentoDAT
    etdMACAL = 1
    etdMALIC = 2
    etdTIDA = 3
    etdLRDA = 4
    etdCDA = 5
End Enum
Public Enum TipoDivida
    edTributaria = 1
    edNaoTributaria = 2
End Enum
Private TabelaObrigacao As String

'Public Function CarregaDividaGerada(Lista As Object, Optional Contribuinte As String, _
        Optional ExercicioInicial As String,
'        Optional ExercicioFinal As String, Optional CodTributo As String) As Boolean
'    Dim Sql As String
'    Sql = "SELECT TDA_INSCRICAO as Inscricao,TDA_ANO_DIVIDA as AnoDivida,TDA_REFERENCIA_DIVIDA as" & _
'        " Referencia,TDA_DIVIDA_ORIGINAL as [Divida Original],TDA_DATA_INSCRICAO as [Data Insc],TDA_MACAL_NUMERO AS MACAL,TDA_MALIC_NUMERO AS MALIC,TDA_TIDA_NUMERO AS TIDA,TDA_REGISTRO AS Registro,TDA_CDA_NUMERO AS CDA from TAB_DIVIDA_ATIVA WHERE 1=1"
Private Function BuscaPropietario(Imovel As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "Select tim_tci_im from tab_imovel where tim_ic  ='" & Imovel & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaPropietario = Rs!tim_tci_im
    End If
    Bdados.FechaTabela Rs
End Function

'    If Trim(Contribuinte) <> "" Then Sql = Sql & " AND TDA_INSCRICAO ='" & Contribuinte & "'"
'    If Trim(ExercicioInicial) <> "" Then Sql = Sql & " AND TDA_ANO_DIVIDA >= " & ExercicioInicial
'    If Trim(ExercicioFinal) <> "" Then Sql = Sql & " AND TDA_ANO_DIVIDA <= " & ExercicioFinal
'    If Trim(CodTributo) <> "" Then Sql = Sql & " AND TDA_TIP_COD_IMPOSTO = '" & CodTributo & "'"
'    CarregaDividaGerada = Lista.Preencher(Bdados, Sql, 1200, 1100, 1000, 1200, 1400, 1500)
'    Screen.MousePointer = 0
'End Function
Public Function CarregaDividaGerada(Lista As Object, Optional Contribuinte As String, Optional Exercicio As String, _
            Optional PeriodoInicial As String, Optional PeriodoFinal As String, Optional CodTributo As String, Optional Registro As String, Optional FolhaInicial As String, Optional FolhaFinal As String, Optional Livro As String, Optional TipoInsc As TipoInsc = InscContrib, Optional Divida As TipoDivida) As Boolean
    
    Dim Sql As String
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "SELECT TDA_INSCRICAO as Inscricao,TDA_ANO_DIVIDA as AnoDivida,TDA_REFERENCIA_DIVIDA as" & _
        " Referencia,TDA_DIVIDA_ORIGINAL as [Divida Original],TDA_DATA_INSCRICAO as [Data Insc],TDA_MACAL_NUMERO AS MACAL,TDA_MALIC_NUMERO AS MALIC,TDA_TIDA_NUMERO AS TIDA,TDA_REGISTRO AS Registro,TDA_CDA_NUMERO AS CDA ,tda_Folha as Folha,Tda_Livro as Livro from TAB_DIVIDA_ATIVA WHERE 1=1"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "SELECT TDA_INSCRICAO as Inscricao,TDA_ANO_DIVIDA as AnoDivida,TDA_REFERENCIA_DIVIDA as" & _
        " Referencia,TDA_DIVIDA_ORIGINAL as Divida_Original,TDA_DATA_INSCRICAO as Data_Insc,TDA_MACAL_NUMERO AS MACAL,TDA_MALIC_NUMERO AS MALIC,TDA_TIDA_NUMERO AS TIDA,TDA_REGISTRO AS Registro,TDA_CDA_NUMERO AS CDA ,tda_Folha as Folha,Tda_Livro as Livro from TAB_DIVIDA_ATIVA WHERE 1=1"
    End If
    
    If Trim(Contribuinte) <> "" Then Sql = Sql & " AND TDA_INSCRICAO ='" & Contribuinte & "' "
    If TipoInsc > 0 Then
        Sql = Sql & " and (tda_tipo_inscricao =" & TipoInsc & ")"
    End If
    If Trim(PeriodoInicial) <> "" Then Sql = Sql & " AND TDA_ANO_DIVIDA >= " & PeriodoInicial
    If Trim(PeriodoFinal) <> "" Then Sql = Sql & " AND TDA_ANO_DIVIDA <= " & PeriodoFinal
    If Trim(CodTributo) <> "" Then Sql = Sql & " AND TDA_TIP_COD_IMPOSTO = '" & CodTributo & "'"
    If Trim(Registro) <> "" Then Sql = Sql & " AND TDA_REGISTRO = '" & Registro & "'"
    
    If Trim(FolhaInicial) <> "" And Trim(FolhaFinal) <> "" Then
    Sql = Sql & " and TDA_FOLHA >= '" & FolhaInicial & "' and TDA_FOLHA <= '" & FolhaFinal & "'"
    ElseIf Trim(FolhaInicial) <> "" And Trim(FolhaFinal) = "" Then
        Sql = Sql & " and TDA_FOLHA >= '" & FolhaInicial & "' and TDA_FOLHA <= '" & FolhaInicial & "'"
    End If
    
    If Trim(Livro) <> "" Then
        Sql = Sql & " and TDA_LIVRO = '" & Livro & "'"
    End If
    If Divida = 0 Then Divida = edTributaria
    Sql = Sql & " AND TDA_TIPO_DIVIDA = '" & Divida & "'"
    
    Sql = Sql & " ORDER BY TDA_REFERENCIA_DIVIDA ASC"
    CarregaDividaGerada = Lista.Preencher(Bdados, Sql, 1200, 1100, 1000, 1200, 1400, 1500)
    Screen.MousePointer = 0
End Function

Public Function BuscaParametro(Parametro As String, Divida As TipoDivida) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    If Divida = edTributaria Then
        Sql = "Select TPD_DESCRICAO "
        Sql = Sql & " FROM TAB_PARAMETRO_DIVIDA_ATIVA "
        Sql = Sql & " where TPD_PARAMETRO ='" & Parametro & "'"
    ElseIf Divida = edNaoTributaria Then
        Sql = "Select TPD_DESCRICAO "
        Sql = Sql & " FROM TAB_PARAM_DIV_ATI_NAO_TRIBU "
        Sql = Sql & " where TPD_PARAMETRO ='" & Parametro & "'"
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaParametro = Rs!TPD_DESCRICAO
    End If
    Bdados.FechaTabela Rs
End Function

'
Public Function GeraDocumentacao(Documento As TipoDocumentoDAT, Optional ExercicioInicial As String, Optional ExercicioFinal As String, _
                Optional Exercicio As String, Optional Inscricao As String, Optional TipoInscricao As TipoInsc = 0, Optional CodTributo As String, Optional NumProcesso As String, _
                Optional Processo As String, Optional Barra As Object, Optional Perc As Object, _
                Optional ByVal Percentual As Object, Optional ByVal RegistroP As Object, Optional Divida As TipoDivida) As Boolean
    Dim Rs As VSRecordset
    Dim Campos As String
    Dim Valores As String
    Dim Sql As String
    Dim RsDAT As VSRecordset
    Dim Obrig As New Obrigacao
    Dim Resgistro As Double
    Dim CondicaoLanc As String
    Dim Registro As Double
    Dim Folha As Double
    Dim Livro As Double
    Dim TotalDivida As Double
    Dim Juros As Double
    Dim Multa As Double
    Dim Correcao As Double
    Inscricao = Trim(Inscricao)
    Dim NomeDocumento As String
    Dim Ano As Integer
    Dim AnoAuxiliar As Integer
    
    Dim Retorno As Boolean
    Dim Sql_Progresso As String
    Dim Max As Variant
    Dim Contador As Variant
    Dim DocumentoAtual As TipoDocumentoDAT
    Dim Proprietario As String
    Retorno = False
    For Ano = Nvl(ExercicioInicial, 0) To Nvl(ExercicioFinal, 0)
        
        GeraDocumentacao = True
        CondicaoLanc = ""
        If Trim(Inscricao) <> "" Then
            CondicaoLanc = CondicaoLanc & " and (TOC_INSCRICAO ='" & Inscricao & "'"
            If TipoInscricao <> 0 Then
                CondicaoLanc = CondicaoLanc & " and TOC_TIPO_INSCRICAO =" & TipoInscricao & ")"
            Else
                CondicaoLanc = CondicaoLanc & ")"
            End If
        End If
        If Trim(CodTributo) <> "" Then CondicaoLanc = CondicaoLanc & " and TOC_TIP_COD_IMPOSTO ='" & CodTributo & "'"
        If Bdados.Conexao.FormatoBanco = SQLServer Then
            If Trim(Exercicio) <> "" Then CondicaoLanc = CondicaoLanc & " and YEAR(TOC_DATA_GERACAO) =" & Exercicio
        ElseIf Bdados.Conexao.FormatoBanco = oracle Then
            If Trim(Exercicio) <> "" Then CondicaoLanc = CondicaoLanc & " and TOC_DATA_GERACAO BETWEEN " & "'01/01/" & Exercicio & "' AND '31/12/" & Exercicio & "'"
        End If
        
        Sql = "SELECT * FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_STATUS_OBRIGACAO in (" & Const_NaoPagos & ")" & _
        " AND TOC_DATA_VENCIMENTO <=" & Bdados.Converte(Date, TCDataHora) & CondicaoLanc & _
        " "
        If Ano > 0 Then
            If Bdados.Conexao.FormatoBanco = SQLServer Then
                Sql = Sql & " AND SUBSTRING(CONVERT(VARCHAR(6),TOC_PERIODO),1,4) = " & Ano
            ElseIf Bdados.Conexao.FormatoBanco = oracle Then
                Sql = Sql & " AND SUBSTR(TO_CHAR(TOC_PERIODO),1,4) = " & Ano
            End If
        End If
        Sql = Sql & "  order by TOC_INSCRICAO, TOC_TIP_COD_IMPOSTO"
        
        Sql_Progresso = "SELECT count(*) as Total FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_STATUS_OBRIGACAO in (" & Const_NaoPagos & ")" & _
        " AND TOC_DATA_VENCIMENTO <=" & Bdados.Converte("31/12/" & Year(Date), TCDataHora) & CondicaoLanc
        If Bdados.Conexao.FormatoBanco = SQLServer Then
            If Ano > 0 Then Sql_Progresso = Sql_Progresso & " AND SUBSTRING(CONVERT(VARCHAR(6),TOC_PERIODO),1,4) = " & Ano
        ElseIf Bdados.Conexao.FormatoBanco = oracle Then
            If Ano > 0 Then Sql_Progresso = Sql_Progresso & " AND SUBSTR(TO_CHAR(TOC_PERIODO),1,4) = " & Ano
        End If
        If Bdados.AbreTabela(Sql_Progresso) Then
            If Bdados.Tabela("Total") >= 0 Then
                Max = Bdados.Tabela("Total")
            End If
        End If
        If Bdados.AbreTabela(Sql, Rs) Then
            Contador = 0
            
            Rs.MoveFirst
            Do While Not Rs.EOF
                Contador = Contador + 1
                RegistroP = Format(Contador, "000000")
                Percentual = Val(100 * (Contador / IIf(Max = 0, Contador, Max))) & " % "
                
                For DocumentoAtual = IIf(Documento = 0, etdMACAL, Documento) To IIf(Documento = 0, etdCDA, Documento)
                    Select Case DocumentoAtual
                        Case etdMACAL
                            Registro = VerificaPreRequisito(Documento, Rs!TOC_COD_OBRIGACAO, Divida)
                            NomeDocumento = "MACAL"
                            If Rs!TOC_TIPO_INSCRICAO = 1 Then
                                Proprietario = BuscaPropietario(Rs!TOC_INSCRICAO)
                            Else
                                Proprietario = Trim(Rs!TOC_INSCRICAO)
                            End If
                            If Registro = 0 Then
                                Registro = GeraCodigoDocumento(Rs!TOC_COD_OBRIGACAO, NomeDocumento, -1, Divida)
                                If Registro = 0 Then
                                    Avisa "Não será possível continuar. Processo interrompido por falta de parâmetros."
                                    GeraDocumentacao = False
                                    Exit Function
                                End If
                            End If
                            
                            Retorno = True
                            Correcao = Format(Conta.CalculaValoresCorrecaoAvulso(Rs!TOC_TIP_COD_IMPOSTO, Rs!TOC_PERIODO, Rs!TOC_DATA_VENCIMENTO, Format(Date, "dd/mm/yyyy"), Rs!TOC_VALOR_OBRIGACAO, 3), Const_Monetario)
                            Juros = Format(Conta.CalculaValoresJurosAvulsos(Rs!TOC_TIP_COD_IMPOSTO, Rs!TOC_PERIODO, EtcCreditoTributario, Format(Date, "dd/mm/yyyy"), Rs!TOC_DATA_VENCIMENTO, Rs!TOC_VALOR_OBRIGACAO + Correcao, 3), Const_Monetario)
                            Multa = Format(Conta.CalculaValoresMultaAvulsos(Rs!TOC_TIP_COD_IMPOSTO, Rs!TOC_PERIODO, EtcCreditoTributario, Format(Date, "dd/mm/yyyy"), Rs!TOC_DATA_VENCIMENTO, Rs!TOC_VALOR_OBRIGACAO + Correcao, 3), Const_Monetario)
                            
                            Campos = "TDA_TOC_COD_OBRIGACAO,TDA_INSCRICAO,TDA_ANO_DIVIDA,TDA_REFERENCIA_DIVIDA," & _
                                    "TDA_DIVIDA_ORIGINAL,TDA_JUROS,TDA_MULTA,TDA_CORRECAO,TDA_DIVIDA_TOTAL,TDA_MACAL_DATA,TDA_MACAL_NUMERO," & _
                                    "TDA_TUS_COD_USUARIO,TDA_NATUREZA_TRIBUTO,TDA_TIP_COD_IMPOSTO,TDA_TCI_IM,TDA_TIPO_INSCRICAO"
                            Valores = Bdados.PreparaValor(Rs!TOC_COD_OBRIGACAO, Bdados.Converte(Trim(Rs!TOC_INSCRICAO), tctexto), IIf(Ano = 0, Left(Rs!TOC_PERIODO, 4), Ano), Rs!TOC_PERIODO, _
                                        Rs!TOC_VALOR_OBRIGACAO, Bdados.Converte(Juros, TCDuplo), Bdados.Converte(Multa, TCDuplo), Rs!TOC_VALOR_OBRIGACAO + Multa + Juros, _
                                        Rs!TOC_VALOR_OBRIGACAO + Multa + Juros + Rs!TOC_VALOR_OBRIGACAO, Format(Date, "dd/mm/yyyy"), Registro, AplicacoesVTFuncoes.Usuario, 1, Rs!TOC_TIP_COD_IMPOSTO, Proprietario, Rs!TOC_TIPO_INSCRICAO)
                        Case etdMALIC
                            Registro = VerificaPreRequisito(Documento, Rs!TOC_COD_OBRIGACAO, Divida)
                            If Registro > 0 Then
                                Retorno = True
                                NomeDocumento = "MALIC"
                                Campos = "TDA_MALIC_DATA,TDA_MALIC_NUMERO"
                                Valores = Bdados.PreparaValor(Format(Date, "dd/mm/yyyy"), Registro)
                            End If
                        Case etdLRDA
                            Registro = VerificaPreRequisito(Documento, Rs!TOC_COD_OBRIGACAO, Divida)
                            If Registro > 0 Then
                                NomeDocumento = "LRDA"
                                Retorno = True
                                Campos = "TDA_LIVRO,TDA_FOLHA"
                                
                                Sql = "SELECT TDA_LIVRO,TDA_FOLHA,TDA_REGISTRO "
                                Sql = Sql & "FROM TAB_DIVIDA_ATIVA "
                                Sql = Sql & "WHERE TDA_TOC_COD_OBRIGACAO = " & Rs!TOC_COD_OBRIGACAO
                                Sql = Sql & " and TDA_TIPO_DIVIDA = '" & Divida & "'"
                                
                                If DocumentoAtual = etdLRDA Then
                                    Retorno = True
                                    Livro = BuscaParametro("LIVRO ATUAL", Divida)
                                    Folha = BuscaParametro("FOLHA ATUAL", Divida)
                                    If (Registro Mod CInt(Nvl(BuscaParametro("REGISTROS POR FOLHA", Divida), 0))) = 0 Then
                                            Folha = Folha + 1
                                            If Divida = edTributaria Then
                                                Bdados.AtualizaDados "TAB_PARAMETRO_DIVIDA_ATIVA", Bdados.PreparaValor(Folha), "TPD_DESCRICAO", "TPD_PARAMETRO ='FOLHA ATUAL'"
                                            ElseIf Divida = edNaoTributaria Then
                                                Bdados.AtualizaDados "TAB_PARAM_DIV_ATI_NAO_TRIBU", Bdados.PreparaValor(Folha), "TPD_DESCRICAO", "TPD_PARAMETRO ='FOLHA ATUAL'"
                                            End If
                                    End If
                                    If Folha >= BuscaParametro("FOLHAS POR LIVRO", Divida) Then
                                        If Divida = edTributaria Then
                                            Bdados.AtualizaDados "TAB_PARAMETRO_DIVIDA_ATIVA", Bdados.PreparaValor(Livro + 1), "TPD_DESCRICAO", "TPD_PARAMETRO ='LIVRO ATUAL'"
                                            Bdados.AtualizaDados "TAB_PARAMETRO_DIVIDA_ATIVA", Bdados.PreparaValor(1), "TPD_DESCRICAO", "TPD_PARAMETRO ='FOLHA ATUAL'"
                                        ElseIf Divida = edNaoTributaria Then
                                            Bdados.AtualizaDados "TAB_PARAM_DIV_ATI_NAO_TRIBU", Bdados.PreparaValor(Livro + 1), "TPD_DESCRICAO", "TPD_PARAMETRO ='LIVRO ATUAL'"
                                            Bdados.AtualizaDados "TAB_PARAM_DIV_ATI_NAO_TRIBU", Bdados.PreparaValor(1), "TPD_DESCRICAO", "TPD_PARAMETRO ='FOLHA ATUAL'"
                                        End If
                                    End If
                                    
                                End If
                                
                                Valores = Bdados.PreparaValor(Livro, Folha)
                            End If
                        Case etdTIDA
                            Registro = VerificaPreRequisito(Documento, Rs!TOC_COD_OBRIGACAO, Divida)
                            If Registro > 0 Then
                                NomeDocumento = "TIDA"
                                Retorno = True
                                Campos = "TDA_TIDA_DATA,TDA_TIDA_NUMERO,TDA_DATA_INSCRICAO,TDA_DIVIDA_ORIGINAL, " & _
                                    "TDA_JUROS,TDA_MULTA,TDA_CORRECAO,TDA_DIVIDA_TOTAL,TDA_REGISTRO"
                                Correcao = Format(Conta.CalculaValoresCorrecaoAvulso(Rs!TOC_TIP_COD_IMPOSTO, Rs!TOC_PERIODO, Rs!TOC_DATA_VENCIMENTO, Format(Date, "dd/mm/yyyy"), Rs!TOC_VALOR_OBRIGACAO, 3), Const_Monetario)
                                Juros = Format(Conta.CalculaValoresJurosAvulsos(Rs!TOC_TIP_COD_IMPOSTO, Rs!TOC_PERIODO, EtcCreditoTributario, Format(Date, "dd/mm/yyyy"), Rs!TOC_DATA_VENCIMENTO, Rs!TOC_VALOR_OBRIGACAO + Correcao, 3), Const_Monetario)
                                Multa = Format(Conta.CalculaValoresMultaAvulsos(Rs!TOC_TIP_COD_IMPOSTO, Rs!TOC_PERIODO, EtcCreditoTributario, Format(Date, "dd/mm/yyyy"), Rs!TOC_DATA_VENCIMENTO, Rs!TOC_VALOR_OBRIGACAO + Correcao, 3), Const_Monetario)
                                
                                Valores = Bdados.PreparaValor(Format(Date, "dd/mm/yyyy"), _
                                            Registro, _
                                            Format(Date, "dd/mm/yyyy"), Bdados.Converte(Rs!TOC_VALOR_OBRIGACAO, TCDuplo), Bdados.Converte(Juros, TCDuplo), Bdados.Converte(Multa, TCDuplo), _
                                            Bdados.Converte(Correcao, TCDuplo), Bdados.Converte(Format(Rs!TOC_VALOR_OBRIGACAO + Juros + Multa, Const_Monetario), TCDuplo), Registro)
                                Obrig.TrocaSitObrigacao Rs!TOC_COD_OBRIGACAO, etsCreditoDividaAtiva
                            End If
                        Case etdCDA
                            Registro = VerificaPreRequisito(Documento, Rs!TOC_COD_OBRIGACAO, Divida)
                            If Registro > 0 Then
                                NomeDocumento = "CDA"
                                Retorno = True
                                Campos = "TDA_CDA_DATA,TDA_CDA_NUMERO"
                                Valores = Bdados.PreparaValor(Format(Date, "dd/mm/yyyy"), Registro)
                            End If
                    End Select
                    
                    If Retorno Then
                        If Divida = edTributaria Then
                            Valores = Valores & Bdados.PreparaValor(NumProcesso, Processo, BuscaParametro("CARGO AUTORIDADE", Divida), BuscaParametro("AUTORIDADE COMPETENTE", Divida), 1, 1)
                            Campos = Campos & ", TDA_NUM_PROCESSO,TDA_PROCESSO,TDA_CARGO,TDA_AUTORIDADE,TDA_TIPO_DIVIDA,TDA_STATUS"
                        ElseIf Divida = edNaoTributaria Then
                            Valores = Valores & Bdados.PreparaValor(NumProcesso, Processo, BuscaParametro("CARGO AUTORIDADE", Divida), BuscaParametro("AUTORIDADE COMPETENTE", Divida), 2, 1)
                            Campos = Campos & ", TDA_NUM_PROCESSO,TDA_PROCESSO,TDA_CARGO,TDA_AUTORIDADE,TDA_TIPO_DIVIDA,TDA_STATUS"
                        End If
                        Call Bdados.GravaDados("TAB_DIVIDA_ATIVA", Valores, Campos, "TDA_TOC_COD_OBRIGACAO = " & Rs!TOC_COD_OBRIGACAO & " AND TDA_TIPO_DIVIDA = '" & Divida & "'")
                    Else
                        GeraDocumentacao = False
                    End If
                Next
                Rs.MoveNext
                DoEvents
            Loop
        Else
            GeraDocumentacao = False
        End If
        Bdados.FechaTabela Rs
    Next
End Function

Private Function VerificaPreRequisito(Documento As TipoDocumentoDAT, CodObrigacao As String, Divida As TipoDivida) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "SELECT TDA_MACAL_NUMERO "
    Sql = Sql & "FROM TAB_DIVIDA_ATIVA "
    Sql = Sql & " WHERE TDA_TOC_COD_OBRIGACAO = " & CodObrigacao
    Sql = Sql & " and TDA_TIPO_DIVIDA = '" & Divida & "'"
    VerificaPreRequisito = True
'    If Documento = etdMACAL Then Exit Function
    Select Case Documento
        Case etdMALIC
            Sql = Sql & " AND TDA_MACAL_DATA IS NOT NULL"
        Case etdTIDA
            Sql = Sql & " AND TDA_MALIC_DATA IS NOT NULL"
        Case etdLRDA
            Sql = Sql & " AND TDA_TIDA_DATA IS NOT NULL"
        Case etdCDA
            Sql = Sql & " AND TDA_LIVRO IS NOT NULL"
    End Select
    If Bdados.AbreTabela(Sql) Then
        VerificaPreRequisito = Bdados.Tabela!TDA_MACAL_NUMERO
    Else
        VerificaPreRequisito = 0
    End If
End Function

Private Function GeraCodigoDocumento(CodObrigacao As String, NomeDocumento As String, Ano As Integer, Divida As TipoDivida) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Campo As String
    Dim registroAtual As String
    
    GeraCodigoDocumento = 0
    Campo = "TDA_TIDA_NUMERO"
    Select Case NomeDocumento
        Case "MACAL"
            Campo = "TDA_TIDA_NUMERO" '" TDA_MACAL_NUMERO "
        Case "MALIC"
            Campo = "TDA_TIDA_NUMERO" '" TDA_MALIC_NUMERO "
        Case "TIDA"
            Campo = "TDA_TIDA_NUMERO"
        Case "CDA"
            Campo = "TDA_TIDA_NUMERO" '" TDA_CDA_NUMERO "
        Case "REGISTRO DAT"
            Campo = "TDA_TIDA_NUMERO" '"TDA_REGISTRO"
        
            
    End Select
    Sql = "SELECT " & Campo & " FROM TAB_DIVIDA_ATIVA"
    Sql = Sql & " WHERE TDA_TOC_COD_OBRIGACAO = " & CodObrigacao
    Sql = Sql & " and TDA_TIPO_DIVIDA = '" & Divida & "'"
    
    NomeDocumento = "REGISTRO DAT"
    If Bdados.AbreTabela(Sql, Rs) Then
        If "" & Rs(0) = "" Then
            GeraCodigoDocumento = BuscaParametro("PROXIMO REGISTRO", Divida) ' Conta.GeraCodPagamento(NomeDocumento, Ano)
            If Divida = edTributaria Then
                Bdados.Executa "UPDATE TAB_PARAMETRO_DIVIDA_ATIVA SET TPD_DESCRICAO  = CAST(TPD_DESCRICAO AS INT) + 1 where TPD_PARAMETRO ='PROXIMO REGISTRO'"
            ElseIf Divida = edNaoTributaria Then
                Bdados.Executa "UPDATE TAB_PARAM_DIV_ATI_NAO_TRIBU SET TPD_DESCRICAO  = CAST(TPD_DESCRICAO AS INT) + 1 where TPD_PARAMETRO ='PROXIMO REGISTRO'"
            End If
        Else
            GeraCodigoDocumento = "" & Rs(0)
        End If
    Else
        registroAtual = BuscaParametro("PROXIMO REGISTRO", Divida)
        If Trim(registroAtual) = "" Then
            If Divida = edTributaria Then
                Avisa "Falta definição do Nº do PROXIMO REGISTRO em [PARAMETROS DA DÍVIDA ATIVA]."
            ElseIf Divida = edNaoTributaria Then
                Avisa "Falta definição do Nº do PROXIMO REGISTRO em [PARAMETROS DA DÍVIDA ATIVA NÃO TRIBUTÁRIA]."
            End If
            GeraCodigoDocumento = 0
        Else
            GeraCodigoDocumento = registroAtual 'Conta.GeraCodPagamento(UCase(NomeDocumento), Ano)
            If Divida = edTributaria Then
                Bdados.Executa "UPDATE TAB_PARAMETRO_DIVIDA_ATIVA SET TPD_DESCRICAO  = CAST(TPD_DESCRICAO AS INT) + 1 where TPD_PARAMETRO ='PROXIMO REGISTRO'"
            ElseIf Divida = edNaoTributaria Then
                Bdados.Executa "UPDATE TAB_PARAM_DIV_ATI_NAO_TRIBU SET TPD_DESCRICAO  = CAST(TPD_DESCRICAO AS INT) + 1 where TPD_PARAMETRO ='PROXIMO REGISTRO'"
            End If
        End If
    End If
End Function
Private Sub Class_Initialize()
    NumRegPorFolha = Nvl(Temp.PegaParametro(Bdados, "REGISTRO POR FOLHA DAT"), 0)
    NumFolhaPorLivro = Nvl(Temp.PegaParametro(Bdados, "FOLHA POR LIVRO"), 0)
    
End Sub

