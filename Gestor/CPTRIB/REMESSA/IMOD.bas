Public Sub Imprimir_Boleto(CodPagamento As String, Optional Banco As eBancoBoleto = eBBBanco_Brasil, Optional destino As ediDetinoImpresao = ediTela)
   On Error GoTo TRATA
    Dim rs As VSRecordset
    Dim sql As String
    Dim CodBarra As New CodigoDeBarra
    Dim LinhaDigitavel As String
    Dim Report As New VSRelatorio
    Dim P As Object
    Dim ValJuros As Double
    Dim ValMinMulta As Double
    Dim ValMaxMulta As Double
    Dim Rs1 As VSRecordset
    Dim NovoVencimento As String
     'NovoVencimento = PegaNovaDataVencimento(Vencimento)
    
     'If CDate(NovoVencimento) < Date Then
     '   Avisa "Vencimento inválido."
     '   Exit Sub
     'End If
    
    If Banco = eBBBanco_Bradesco Then
        With Report
            If Not .DefinirArquivo(Bdados, App.Path & "\TDAMBARRA_BRADESCO.rpt") Then Exit Sub
                'Dados do boleto...
                .Formulas "VT_CEDENTE", RetornaDadosEscola(PegaConfiguracaoEscola(Escola), TES_NOME)
                .Formulas "VT_AGENCIA_CODIGO", PegaConfiguracaoEscola(TEC_BRADESCO_AGENCIA) & " / " & PegaConfiguracaoEscola(TEC_BRADESCO_CONTA_CORRENTE) & "-" & PegaConfiguracaoEscola(TEC_BRADESCO_DV_CONTA_CORRENTE)
                .Formulas "VT_CARTEIRA", PegaConfiguracaoEscola(TEC_BRADESCO_CARTEIRA)
                .Formulas "VT_ACEITE", "NÃO"
                .Formulas "VT_DATA_PROCESSAMENTO", Date
                
                If Bdados.AbreTabela("SELECT * FROM VIS_CONTA_RECEBER WHERE Código = " & CodPagamento) Then
                sql = "Select * from VIS_DADOS_ALUNO_FINANCEIRO where TRP_CODIGO = " & Bdados.Tabela("CodRepFin")
                If Bdados.AbreTabela(sql, rs) Then
                    .Formulas "VT_SACADO", rs.Fields("TRP_NOME")
                    .Formulas "VT_SACADO_CPF", rs.Fields("TRP_CODIGO") & " - " & rs.Fields("TRP_NOME") & Space(10) & " CPF:" & rs.Fields("TRP_DOC")
                    .Formulas "VT_ENDERECO_SACADO", rs.Fields("TRP_ENDERECO") & " Nº " & rs.Fields("TRP_NUMERO") & " Bairro : " & rs.Fields("TRP_BAIRRO")
                    .Formulas "VT_CEP_CIDADE_ESTADO_SACADO", rs.Fields("TRP_CEP") & Space(3) & "COLINAS-MA"
                End If
                .Formulas "VT_CARTEIRA_NOSSO_NUMERO", PegaConfiguracaoEscola(TEC_BRADESCO_CARTEIRA) & " / " & CodPagamento & "-" & CodBarra.Calculo_DV11(PegaConfiguracaoEscola(TEC_BRADESCO_CARTEIRA) & CodPagamento)
                .Formulas "VT_NOSSO_NUMERO", CodPagamento & "-" & CodBarra.Calculo_DV11(PegaConfiguracaoEscola(TEC_BRADESCO_CARTEIRA) & CodPagamento)
                .Formulas "VT_NUMERO_DOCUMENTO", CodPagamento & "-" & CodBarra.Calculo_DV11(PegaConfiguracaoEscola(TEC_BRADESCO_CARTEIRA) & CodPagamento)
                .Formulas "VT_DATA_DOCUMENTO", Date
                .Formulas "VT_VENCIMENTO", Bdados.Tabela("Vencimento")
                .Formulas "VT_VALOR_DOCUMENTO", FormatNumber(Bdados.Tabela("VALOR"), 2)
                .Formulas "VT_MORA_MULTA ", " " '"0,00"
                .Formulas "VT_OUTROS_ACRESCIMOS ", " " '"0,00"
                .Formulas "VT_OUTRAS_DEDUCOES ", " " '"0,00"
                .Formulas "VT_VALOR_COBRADO", " " ' Format(Bdados.Tabela("Saldo Devedor"), Const_Monetario)
                .Formulas "VT_DESCONTO", " " ' Format(Bdados.Tabela("Desconto"), Const_Monetario)
                .Formulas "VT_MENSAGEM", PegaConfiguracaoEscola(TEC_MENSAGEM_BRADESCO)
                .Formulas "VT_MENSAGEM_2", PegaConfiguracaoEscola(TEC_MENSAGEM_BRADESCO_2)
                .Formulas "VT_MENSAGEM_3", PegaConfiguracaoEscola(TEC_MENSAGEM_BRADESCO_3)
                .Formulas "VT_MENSAGEM_4", "APÓS O VENCIMENTO COBRAR JUROS DE  " & Format(PegaConfiguracaoEscola(Juros) * Bdados.Tabela("Saldo Devedor") / 30, Const_Monetario) & " AO DIA E MULTA DE 2% AO MÊS"
                .Formulas "VT_MENSAGEM_5", PegaConfiguracaoEscola(TEC_MENSAGEM_BRADESCO_5)
                
                If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
                    LinhaDigitavel = CodBarra.CriaLinhaDigitavelCBR(rs.Fields("TRP_CODIGO"), CDbl(Bdados.Tabela("Saldo Devedor")), Year(Bdados.Tabela("Vencimento")), P, Bdados.Tabela("Vencimento"), Bdados.Tabela("Parcela"), CodPagamento)
                Else
                    LinhaDigitavel = CodBarra.CriaLinhaDigitavel(rs.Fields("TRP_CODIGO"), CDbl(Bdados.Tabela("Saldo Devedor")), Year(Bdados.Tabela("Vencimento")), Bdados.Tabela("Vencimento"), Bdados.Tabela("Parcela"), etcbDebitoNormal)
                End If
                'Dados do código de barras...
                .Formulas "VT_LINHA_DIGITAVEL", LinhaDigitavel
                .Formulas "VT_LinhaBarra", CodBarra.LinhaBarraGerada
                '.Formulas "VT_CodBarra", CodBarra.LinhaBarraGerada
                .Titulo = "Boleto Bancário"
        '        .CopiasDetalhes = 3
                .Arvore = False
                If destino = ediTela Then
                    .Visualizar
                Else
                    .Imprimir
                End If
                DoEvents
            End If
        End With
    End If
    Set Report = Nothing
    Exit Sub

TRATA:
    
    If Err.Number = 20515 Or Err.Number = 3265 Then
        Report.Formulas "Mensagem1 ", ""
        Resume
    End If
    Avisa "O Boleto não será impresso. Informe o erro a seguir ao suporte."
    Avisa Err.Number & " - " & Err.Description
    Exit Sub
    Resume
End Sub
