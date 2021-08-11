VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrdemServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim cv As String, campos As String, valores As String
Dim Rpt As VSRelatorio
    
    
Public Function Salvar(Tipo As Integer, contribuinte As String, fiscal As String, periodoInicial As String, _
                         periodoFinal As String, Data As Date, hora As String, outrosServicos As String, fiscal2 As String) As Boolean
    Dim Codigo As Long, ano As Integer
    Codigo = Imposto.GeraNumNota(1, 100)
    ano = Format(Now, "yyyy")
    cv = ano & Right(contribuinte, 2) & Codigo
    campos = " CODIGO, COD_VERIFICACAO, ANO, IM_CONTRIBUINTE, FISCAL, TIPO," _
        & " PERIODO_INICIAL,PERIODO_FINAL,DIA_PRAZO,DATA_AUTORIZACAO,HORA_AUTORIZACAO,OUTROS_SERVICOS,FISCAL2"
    valores = Bdados.PreparaValor(Codigo, cv, ano, contribuinte, fiscal, Tipo, periodoInicial, periodoFinal, _
        90, Data, hora, outrosServicos, fiscal2)
    If Bdados.InsereDados("Tab_bcp_ordem_servico", valores, campos) Then
        Salvar = True
    Else
        Salvar = False
    End If
End Function
Public Function IncluirDocumentosDiligencia(os As Long, descricao As String) As Boolean
    campos = " COD_ORDEM, COD_ITEM, DESCRICAO"
    Dim ui As Integer ' ULTIMO ITEM
    Dim rs As VSRecordset
    If Bdados.AbreTabela("SELECT MAX(COD_ITEM) AS ITEM FROM TAB_BCP_OS_DILIGENCIA WHERE COD_ORDEM=" & os, rs) Then
        ui = IIf(IsNull(rs("ITEM")), 1, rs("ITEM") + 1)
    Else
        ui = 1
    End If
    valores = Bdados.PreparaValor(os, ui, descricao)
    If Bdados.InsereDados("Tab_bcp_os_diligencia", valores, campos) Then
        IncluirDocumentosDiligencia = True
    Else
        IncluirDocumentosDiligencia = False
    End If
    
End Function
Public Sub Imprime(Codigo As Long, ordemServico As Boolean)
    Dim rs As VSRecordset
    Set Rpt = New VSRelatorio
    Dim S As Integer, d As Integer
    If Bdados.AbreTabela(retornaSQL(Codigo), rs) Then
        With Rpt
            If ordemServico = True Then
                If Not .DefinirArquivo(Bdados, App.Path + "\TOrdemServico.rpt") Then Exit Sub
                For S = 1 To 4
                    If S = rs("COD_TIPO") Then
                        .Formulas "S" & S, "X"
                    Else
                        .Formulas "S" & S, " "
                    End If
                    If rs("COD_TIPO") > 4 Then
                        .Formulas "S" & 4, "X"
                    End If
                Next S
            Else
                If Not .DefinirArquivo(Bdados, App.Path + "\TTermoInicioAF.rpt") Then Exit Sub
                S = Len(rs("DOCUMENTOS"))
                For d = 1 To S
                    Dim m As String 'marcado com X ou espaço em brando cada DOCUMENTO na tela
                    m = Mid(rs("DOCUMENTOS"), d, 1)
                    .Formulas "D" & d, m
                Next d
                .Formulas "TIPO_ACAO_FISCAL", rs("TIPO_ACAO_FISCAL")
                .Formulas "MATRICULA", rs("MATRICULA")
                .Formulas "M2", rs("MAT2")
            
                
            End If
            .Formulas "CPF_CNPJ", IIf(IsNull(rs("CPF_CNPJ")), "", rs("CPF_CNPJ"))
            .Formulas "IM", IIf(IsNull(rs("INSCRICAO")), "", rs("INSCRICAO"))
            .Formulas "RAZAOSOCIAL", IIf(IsNull(rs("RAZAO")), "", rs("RAZAO"))
            .Formulas "NOMEFANTASIA", IIf(IsNull(rs("TCI_FANTASIA")), "", rs("TCI_FANTASIA"))
            .Formulas "ENDERECO", IIf(IsNull(rs("VCI_ENDERECO")), "", rs("VCI_ENDERECO"))
            .Formulas "ATIVIDADE", IIf(IsNull(rs("VCI_NOME_ATIVIDADE")), "", rs("VCI_NOME_ATIVIDADE"))
            .Formulas "PROCESSO", rs("CODIGO")
            .Formulas "ANO", rs("ANO")
            .Formulas "PINICIAL", IIf(Len(rs("INICIAL")) > 4, Format(rs("INICIAL"), "00 0000"), rs("INICIAL"))
            .Formulas "PFINAL", IIf(Len(rs("FINAL")) > 4, Format(rs("FINAL"), "00 0000"), rs("FINAL"))
            .Formulas "DIAS", rs("DIAS")
            .Formulas "DATA", Format(rs("DATA"), "DD/MM/YYYY")
            .Formulas "HORA", rs("HORA")
            .Formulas "FISCAL", rs("TUS_NOME")
            .Formulas "F2", rs("NOME_USER2")
            
            .Formulas "OUTROS", IIf(IsNull(rs("OUTROS_SERVICOS")), "", rs("OUTROS_SERVICOS"))
            If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1179 Then
                .Formulas "CIDADE", "CODÓ (MA)"
                .Formulas "ENDERECO_PREFEITURA", "Praça Ferreira Bayma, n. 538, Centro"
            ElseIf Temp.PegaParametro(Bdados, "MUNICIPIO") = 1204 Then
                .Formulas "CIDADE", "GRAJAÚ (MA)"
                .Formulas "ENDERECO_PREFEITURA", "Rua Frei Benjamin de Borno, n. 5, Centro"
            End If
            .Formulas "CODIGO_VERIFICADOR", rs("COD_VERIFICACAO")
            If ordemServico = True Then
                .Imprimir
            Else
                .visualizar
            End If
        End With
        'visualizar rpt, Codigo
    End If
End Sub
Public Sub ImprimeDiligencia(Codigo As Long)
    Dim rs As VSRecordset
    Set Rpt = New VSRelatorio
    Dim S As Integer
    If Bdados.AbreTabela(retornaSQL(Codigo), rs) Then
        With Rpt
            If Not .DefinirArquivo(Bdados, App.Path + "\TDiligencia.rpt") Then Exit Sub
            .Formulas "CPF_CNPJ", IIf(IsNull(rs("CPF_CNPJ")), "", rs("CPF_CNPJ"))
            .Formulas "IM", IIf(IsNull(rs("INSCRICAO")), "", rs("INSCRICAO"))
            .Formulas "RAZAOSOCIAL", IIf(IsNull(rs("RAZAO")), "", rs("RAZAO"))
            .Formulas "NOMEFANTASIA", IIf(IsNull(rs("TCI_FANTASIA")), "", rs("TCI_FANTASIA"))
            .Formulas "ENDERECO", IIf(IsNull(rs("VCI_ENDERECO")), "", rs("VCI_ENDERECO"))
            .Formulas "ATIVIDADE", IIf(IsNull(rs("VCI_NOME_ATIVIDADE")), "", rs("VCI_NOME_ATIVIDADE"))
            .Formulas "PROCESSO", rs("CODIGO")
            .Formulas "ANO", rs("ANO")
            .Formulas "DATA", Format(rs("DATA"), "DD/MM/YYYY")
            .Formulas "MATRICULA", rs("MATRICULA")
            .Formulas "FISCAL", rs("TUS_NOME")
            .Formulas "M2", rs("MAT2")
            .Formulas "F2", rs("NOME_USER2")
            
            If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1179 Then
                .Formulas "CIDADE", "CODÓ (MA)"
                .Formulas "ENDERECO_PREFEITURA", "Praça Ferreira Bayma, n. 538, Centro"
            ElseIf Temp.PegaParametro(Bdados, "MUNICIPIO") = 1204 Then
                .Formulas "CIDADE", "GRAJAÚ (MA)"
                .Formulas "ENDERECO_PREFEITURA", "Rua Frei Benjamin de Borno, n. 5, Centro"
            End If
            .Formulas "CODIGO_VERIFICADOR", rs("COD_VERIFICACAO")
            .Selecao = "{TAB_BCP_OS_DILIGENCIA.COD_ORDEM} = " & Codigo
            .visualizar
            '.Imprimir
        End With
    End If
End Sub
Public Sub ImprimeTermoIntimacao(Codigo As Long)
    Dim rs As VSRecordset
    Set Rpt = New VSRelatorio
    Dim S As Integer
    If Bdados.AbreTabela(retornaSQL(Codigo), rs) Then
        With Rpt
            If Not .DefinirArquivo(Bdados, App.Path + "\TTermoIntimacao.rpt") Then Exit Sub
            .Formulas "CPF_CNPJ", IIf(IsNull(rs("CPF_CNPJ")), "", rs("CPF_CNPJ"))
            .Formulas "IM", IIf(IsNull(rs("INSCRICAO")), "", rs("INSCRICAO"))
            .Formulas "RAZAOSOCIAL", IIf(IsNull(rs("RAZAO")), "", rs("RAZAO"))
            .Formulas "NOMEFANTASIA", IIf(IsNull(rs("TCI_FANTASIA")), "", rs("TCI_FANTASIA"))
            .Formulas "ENDERECO", IIf(IsNull(rs("TCI_LOGRADOURO")), "", rs("TCI_LOGRADOURO"))
            .Formulas "BAIRRO", IIf(IsNull(rs("TCI_BAIRRO")), "", rs("TCI_BAIRRO"))
            .Formulas "CEP", IIf(IsNull(rs("TCI_CEP")), "", rs("TCI_CEP"))
            .Formulas "CID_CONTRIB", IIf(IsNull(rs("TCI_CIDADE")), "", rs("TCI_CIDADE"))
            .Formulas "N_ENDERECO", IIf(IsNull(rs("TCI_NUMERO")), "", rs("TCI_NUMERO"))
            .Formulas "ATIVIDADE", IIf(IsNull(rs("VCI_NOME_ATIVIDADE")), "", rs("VCI_NOME_ATIVIDADE"))
            .Formulas "PROCESSO", rs("CODIGO")
            .Formulas "EMISSAO", Format(rs("DATA"), "DD/MM/YYYY")
            .Formulas "FISCAL", rs("TUS_NOME")
            .Formulas "F2", rs("NOME_USER2")
            
            If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1179 Then
                .Formulas "CIDADE", "CODÓ (MA)"
            ElseIf Temp.PegaParametro(Bdados, "MUNICIPIO") = 1204 Then
                .Formulas "CIDADE", "GRAJAÚ (MA)"
            End If
            .Formulas "CODIGO_VERIFICADOR", rs("COD_VERIFICACAO")
            .Selecao = "{VIS_BCP_ORDEM_SERVICO.CODIGO} = " & Codigo
            .visualizar
        End With
    End If
End Sub
Public Sub ImprimeAutoInfracao(Codigo As Long, Infracao As String, CapLegal As String, TextoJuros As String, TextoMulta As String, parcelamento As Boolean, np As Integer, ref As String, um As String, dois As String, tres As String, quatro As String, cinco As String)
    Dim rs As VSRecordset, rsUp As VSRecordset
    Set Rpt = New VSRelatorio
    Dim S As Integer
    Dim sqlUp As String
    Dim rsDesc As VSRecordset
    If Bdados.AbreTabela("SELECT * FROM TAB_BCP_AUTO WHERE COD_OS=" & Codigo, rsUp) Then
        Dim Conta As New ContaCorrente
        Dim NovoJuro As Double, NovaMulta As Double, Correcao As Double
        Dim periodo As String
        Do While rsUp.EOF = False
            NovoJuro = 0
            NovaMulta = 0
            Correcao = 0
            periodo = Replace(rsUp("PERIODO"), "-", "")
            NovoJuro = Conta.CalculaValoresJurosAvulsos(rsUp("TRIBUTO"), CLng(periodo), EtcCreditoTributario, Date, CDate(rsUp("DATA_VENCIMENTO")), CCur(rsUp("ISS_DEVIDO")), , 0)
            NovaMulta = Conta.CalculaValoresMultaAvulsos(rsUp("TRIBUTO"), CLng(periodo), EtcCreditoTributario, Date, CDate(rsUp("DATA_VENCIMENTO")), CCur(rsUp("ISS_DEVIDO")), , 0)
            'CALCULO SELIC 12-05-2011 - BCP GLEYSON
            'PRECISO PEGAR O METODO DA CONTA CORRENTE (AtualizAtualizacaoMonetariaSelic) para atualizar de acordo com o vencimento (NovaData)
            Correcao = AtualizAtualizacaoMonetariaSelicLocal(CCur(rsUp("ISS_DEVIDO")), CDate(rsUp("DATA_VENCIMENTO")), CStr(NovoJuro), CStr(NovaMulta), CStr("0"))
            'Bdados.FechaTabela (rsContaContribuinte)
            'FECHA CALCULO SELIC
            NovaMulta = Conta.CalculaValoresMultaAvulsos(rsUp("TRIBUTO"), CLng(periodo), EtcCreditoTributario, Date, CDate(rsUp("DATA_VENCIMENTO")), CCur(rsUp("ISS_DEVIDO")) + Correcao + NovoJuro, , 0)
            NovaMulta = NovaMulta / 2
            sqlUp = "UPDATE TAB_BCP_AUTO SET JUROS='" & converteMonetario(CStr(NovoJuro)) & "', MULTA='" _
            & converteMonetario(CStr(NovaMulta)) & "',ATUALIZACAO='" & converteMonetario(CStr(Correcao)) & "',ISS_ATUALIZADO='" _
            & converteMonetario(CCur(rsUp("ISS_DEVIDO")) + CCur(NovoJuro) + CCur(NovaMulta) + CCur(Correcao)) & "' WHERE COD_GERAL=" _
            & rsUp("COD_GERAL")
            
            If Bdados.Executa(sqlUp) Then
            End If
            rsUp.MoveNext
            'Exit Sub
        Loop
    End If
    Dim relat As String
    relat = "\TAutoInfracao.rpt"
    If parcelamento = True Then
        relat = "\TParcelamento.rpt"
    End If
    If Bdados.AbreTabela(retornaSQL(Codigo), rs) Then
        With Rpt
            If Not .DefinirArquivo(Bdados, App.Path + relat) Then Exit Sub
            .Formulas "CPF_CNPJ", IIf(IsNull(rs("CPF_CNPJ")), "", rs("CPF_CNPJ"))
            .Formulas "IM", IIf(IsNull(rs("INSCRICAO")), "", rs("INSCRICAO"))
            .Formulas "RAZAOSOCIAL", IIf(IsNull(rs("RAZAO")), "", rs("RAZAO"))
            .Formulas "NOMEFANTASIA", IIf(IsNull(rs("TCI_FANTASIA")), "", rs("TCI_FANTASIA"))
            .Formulas "ENDERECO", IIf(IsNull(rs("VCI_ENDERECO")), "", rs("VCI_ENDERECO"))
            .Formulas "ATIVIDADE", IIf(IsNull(rs("VCI_NOME_ATIVIDADE")), "", rs("VCI_NOME_ATIVIDADE"))
            .Formulas "PROCESSO", rs("CODIGO")
            .Formulas "ANO", rs("ANO")
            .Formulas "DATA", Format(rs("DATA"), "DD/MM/YYYY")
            .Formulas "MATRICULA", rs("MATRICULA")
            .Formulas "FISCAL", rs("TUS_NOME")
            .Formulas "TIPO_ACAO_FISCAL", rs("TIPO")
            .Formulas "M2", rs("MAT2")
            .Formulas "F2", rs("NOME_USER2")
            
            .Formulas "I", Infracao
            .Formulas "CL", CapLegal
            .Formulas "J", TextoJuros
            .Formulas "M", TextoMulta
                
            
            If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1179 Then
                .Formulas "CIDADE", "CODÓ (MA)"
            ElseIf Temp.PegaParametro(Bdados, "MUNICIPIO") = 1204 Then
                .Formulas "CIDADE", "GRAJAÚ (MA)"
            End If
            Dim rsDocs As VSRecordset
            Dim docs As String
            docs = ""
            If parcelamento = True Then
                If Bdados.AbreTabela("SELECT NUMERO_NOTA FROM TAB_BCP_AUTO WHERE TIPO_DOC=1 AND COD_OS=" & Codigo, rsDocs) Then
                    Do While rsDocs.EOF = False
                        If Len(docs) = 0 Then
                            docs = "Nf: " & docs & rsDocs(0) & " - "
                        Else
                            docs = docs & rsDocs(0) & " - "
                        End If
                        rsDocs.MoveNext
                    Loop
                    
                End If
                .Formulas "docs", docs
                .Formulas "ref", ref
                .Formulas "np", CStr(np)
                If Len(um) > 0 Then
                    .Formulas "1", "1-" & um
                End If
                If Len(dois) > 0 Then
                    .Formulas "2", "2-" & dois
                End If
                If Len(tres) > 0 Then
                    .Formulas "3", "3-" & tres
                End If
                If Len(quatro) > 0 Then
                    .Formulas "4", "4-" & quatro
                End If
                If Len(cinco) > 0 Then
                    .Formulas "5", "5-" & cinco
                End If
            End If
            .Formulas "CODIGO_VERIFICADOR", rs("COD_VERIFICACAO")
            If Bdados.AbreTabela("SELECT SUM(MULTA) AS DESCONTO_MULTA, SUM(ISS_ATUALIZADO) AS ISS FROM TAB_BCP_AUTO WHERE COD_OS=" & Codigo, rsDesc) Then
                .Formulas "DESCONTO_MULTA", Format(rsDesc("DESCONTO_MULTA") / 2, "#,##0.00")
                .Formulas "VALOR_MENOS_DESCONTO", Format(rsDesc("ISS") - (rsDesc("DESCONTO_MULTA") / 2), "#,##0.00")
                If parcelamento = True Then
                    .Formulas "extenso", Extenso(rsDesc("ISS"), "")
                    .Formulas "vp", Format(rsDesc("ISS") / CInt(np), "#,##0.00")
                    .Formulas "extvp", Extenso(rsDesc("ISS") / CInt(np), "")
                End If
            Else
                .Formulas "DESCONTO_MULTA", Format(0, "#,##0.00")
                .Formulas "VALOR_MENOS_DESCONTO", Format(0, "#,##0.00")
                If parcelamento = True Then
                    .Formulas "extenso", ""
                    .Formulas "vp", Format(0, "#,##0.00")
                    .Formulas "extvp", ""
                End If
            End If
            '.Selecao = "{TAB_BCP_OS_DILIGENCIA.COD_ORDEM} = " & Codigo
            .Selecao = "{VIS_BCP_AUTO.CODIGO} =" & Codigo
            .visualizar
            
        End With
    End If
End Sub
Public Function AtualizAtualizacaoMonetariaSelicLocal(ValorOriginal As String, DataVencimento As Date, jurosAtualizado As String, multaAtualizada As String, Desconto As String) As Double
    Dim AliquotaSelic As Double, Atualizacao As Double
    Dim rs As VSRecordset
    Dim Sql As String, periodoAtual As String, periodoObrigacao As String
    periodoObrigacao = Format(DataVencimento, "yyyyMM")
    
    periodoAtual = Format(Now, "yyyyMM")
    Sql = "select TMO_VALOR as aliquota from  TAB_MONETARIA where TMO_PERIODO = " & periodoObrigacao
    If Bdados.AbreTabela(Sql, rs) Then
        If periodoObrigacao < periodoAtual Then
            AliquotaSelic = IIf(IsNull(rs(0)), 0, rs(0))
        Else
            AliquotaSelic = 0
        End If
        Atualizacao = ValorOriginal * (AliquotaSelic / 100)
    End If
    AtualizAtualizacaoMonetariaSelicLocal = Atualizacao
End Function
Public Sub visualizar(r As VSRelatorio, os As Long)
    Dim f As New TPRT109
    f.carregar os
End Sub
Public Function atualizaProcesso(Processo As Long, Etapa As Integer, DataAtual As Date, Dia As Integer, Prazo As Integer) As Boolean
    If Bdados.Executa("UPDATE TAB_BCP_ORDEM_SERVICO SET TIPO=" & Etapa & " WHERE CODIGO=" & Processo) Then
        If Etapa < 99 Then 'CONCLUIDA 'Etapa = 4 Or Etapa = 2 Then '4=TIAF or 2=DILIGENCIA
            Dim status As String
            Dim rs As VSRecordset
            'If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1204 Then 'GRAJAU
             '   Dia = 10
            'Else
             '   Dia = 10
            'End If
            Dim DataLimite As Date
            DataLimite = DateAdd("D", Dia, DataAtual)
            If Bdados.AbreTabela("select tge_nome from tab_geral where tge_tipo=3259 and tge_codigo=" & Etapa, rs) Then
                status = rs(0)
            End If
            If Prazo > 0 Then
                campos = "COD_OS, DATA_ABERTURA, DIAS, DATA_LIMITE,TIAF_STATUS,OS_STATUS"
                valores = Bdados.PreparaValor(CStr(Processo), DataAtual, Dia, DataLimite, "DENTRO DO PRAZO", status)
                If Bdados.InsereDados("tab_bcp_tiaf", valores, campos) Then
                Else
            End If
                
        End If
    End If
        atualizaProcesso = True
    Else
        atualizaProcesso = False
    End If
End Function
Public Function inserirInscricao(Processo As Long, Im As String) As Boolean
    If Bdados.Executa("UPDATE TAB_BCP_ORDEM_SERVICO SET IM_CONTRIBUINTE='" & Im & "' WHERE CODIGO=" & Processo) Then
        inserirInscricao = True
    Else
        inserirInscricao = False
    End If
End Function

Private Function retornaSQL(Codigo As Long)
    retornaSQL = "SELECT * FROM VIS_BCP_ORDEM_SERVICO_REL WHERE CODIGO=" & Codigo
End Function
Public Function PreencheCombo(Combo As Object) As Boolean
    Dim Sql As String
    Sql = "SELECT TGE_NOME, TGE_CODIGO FROM TAB_GERAL WHERE TGE_TIPO=3259 AND TGE_CODIGO >0 order by TGE_CODIGO"
    If Combo.Preencher(Bdados, Sql) Then PreencheCombo = True
End Function
Public Function PreencheComboDocumento(Combo As Object) As Boolean
    Dim Sql As String
    Sql = "SELECT TGE_NOME, TGE_CODIGO FROM TAB_GERAL WHERE TGE_TIPO=3260 AND TGE_CODIGO >0 order by TGE_CODIGO"
    If Combo.Preencher(Bdados, Sql) Then PreencheComboDocumento = True
End Function
Private Function converteMonetario(Valor As String) As String
    Dim sValor As String
    sValor = Replace(CStr(Valor), ".", "")
    sValor = Replace(CStr(sValor), ",", ".")
    converteMonetario = sValor
End Function
Public Function Extenso(ByVal nValor As Double, simbolo As String)

        If nValor <= 0 Or nValor > 9999999.99 Then
            Extenso = "ZERO"
            Exit Function
        End If

        'Declara as variáveis da função
        Dim nContador, nTamanho As Integer
        Dim cValor, cParte, cFinal As String
        Dim aGrupo(4), aTexto(4) As String

        'Define matrizes com extensos parciais
        Dim aUnid(19) As String
        aUnid(1) = "UM ": aUnid(2) = "DOIS ": aUnid(3) = "TRES "
        aUnid(4) = "QUATRO ": aUnid(5) = "CINCO ": aUnid(6) = "SEIS "
        aUnid(7) = "SETE ": aUnid(8) = "OITO ": aUnid(9) = "NOVE "
        aUnid(10) = "DEZ ": aUnid(11) = "ONZE ": aUnid(12) = "DOZE "
        aUnid(13) = "TREZE ": aUnid(14) = "QUATORZE ": aUnid(15) = "QUINZE "
        aUnid(16) = "DEZESSEIS ": aUnid(17) = "DEZESSETE ": aUnid(18) = "DEZOITO "
        aUnid(19) = "DEZENOVE "

        Dim aDezena(9) As String
        aDezena(1) = "DEZ ": aDezena(2) = "VINTE ": aDezena(3) = "TRINTA "
        aDezena(4) = "QUARENTA ": aDezena(5) = "CINQUENTA "
        aDezena(6) = "SESSENTA ": aDezena(7) = "SETENTA ": aDezena(8) = "OITENTA "
        aDezena(9) = "NOVENTA "

        Dim aCentena(9) As String
        aCentena(1) = "CENTO ": aCentena(2) = "DUZENTOS "
        aCentena(3) = "TREZENTOS ": aCentena(4) = "QUATROCENTOS "
        aCentena(5) = "QUINHENTOS ": aCentena(6) = "SEISCENTOS "
        aCentena(7) = "SETECENTOS ": aCentena(8) = "OITOCENTOS "
        aCentena(9) = "NOVECENTOS "

        'Divide o valor em vários grupos
        cValor = Format(nValor, "0000000000.00")
        aGrupo(1) = Mid$(cValor, 2, 3)
        aGrupo(2) = Mid$(cValor, 5, 3)
        aGrupo(3) = Mid$(cValor, 8, 3)
        aGrupo(4) = "0" + Mid$(cValor, 12, 2)

        'Processa cada grupo
        For nContador = 1 To 4
            cParte = aGrupo(nContador)
            nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
            If nTamanho = 3 Then
                If Right$(cParte, 2) <> "00" Then
                    aTexto(nContador) = aTexto(nContador) + aCentena(Left(cParte, 1)) + "E "
                    nTamanho = 2
                Else
                    aTexto(nContador) = aTexto(nContador) + IIf(Left$(cParte, 1) = "1", "CEM ", aCentena(Left(cParte, 1)))
                End If
            End If
            If nTamanho = 2 Then
                If Val(Right(cParte, 2)) < 20 Then
                    aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 2))
                Else
                    aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
                    If Right$(cParte, 1) <> "0" Then
                        aTexto(nContador) = aTexto(nContador) + "E "
                        nTamanho = 1
                    End If
                End If
            End If
            If nTamanho = 1 Then
                aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 1))
            End If
        Next

        'Gera o formato final do texto
        If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
            cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "CENTAVO", "CENTAVOS")
        Else
            cFinal = ""
            cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, "MILHÕES ", "MILHÃO "), "")
            If Val(aGrupo(2) + aGrupo(3)) = 0 Then
                cFinal = cFinal + "DE "
            Else
                cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "MIL ", "")
            End If
            cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "REAL ", "REAIS ")
            cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "E " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, "CENTAVO", "CENTAVOS"), "")
        End If

        Extenso = cFinal
        
    End Function

