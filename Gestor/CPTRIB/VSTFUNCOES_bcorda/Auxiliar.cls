VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Auxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Public Function GeraCodBarra(CodigoPagamento As String, CodigoImposto As String, ByVal ValorTotal As Double, ObjetoPicture As Object, Optional ByVal Periodo As String, Optional ByVal DtVenc As String, Optional Parcela As String, Optional Conta As Byte, Optional GeraBitMap As Boolean = True) As String
'    On Error Resume Next
'    Dim Codebar As New CodBarBB
'    Codebar.Imprime_Barra
'    Codebar.Imp_Codbar 7, 0, 0&, 1, 50, 0, "Times New Roman", 12, 0, 0, 30, 0, 0, StrDigitavel, ObjetoPicture
'
'    Dim Pos As Integer
'    Dim Linha1 As String
'    Dim Linha2 As String
'    Dim Linha3 As String
'    Dim Linha4  As String
'    Dim Linha As String
'    Dim StrValor  As String
'    Dim StrOrgao As String
'    Dim StrNossoNumero As String
'    Dim StrDigitavel As String
'    Dim StrDV  As String
'
'    If Len(Periodo) = 4 Then
'        Periodo = Periodo & "00"
'    Else
'        Periodo = Right(Periodo, 4) & Left(Periodo, 2)
'    End If
'    DtVenc = Right(DtVenc, 4) & Mid(DtVenc, 4, 2) & Left(DtVenc, 2)
'    Parcela = Format(Parcela, "00")
'    'Identificacao do Produto = 8(Arrecadacao) / Segmento = 1(Prefeitura) / Identificador=6(Real)
'    StrDigitavel = "816"
'    StrDigitavel = StrDigitavel & Format((CDbl(ValorTotal) * 100), "00000000000")
'    StrOrgao = Edita.TiraPic(Temp.PegaParametro(Bdados, "FEBRABAN"), ".")   'Busca CGC da Prefeitura
'    StrOrgao = Left(StrOrgao, 4) 'Orgao
'    StrDigitavel = StrDigitavel & StrOrgao
'    'Campo Livre
'    If Len(StrOrgao) = 8 Then
'        StrNossoNumero = MontaNumeroPrefeituraCgc(CodigoImposto, CodigoPagamento)
'    Else
'        StrNossoNumero = MontaNumeroPrefeituraFebraBan(Periodo & DtVenc & Parcela & Conta & CodigoPagamento)
'    End If
'    StrDigitavel = StrDigitavel & StrNossoNumero
'
'    'CALCULA DV GERAL
'    StrDV = Codebar.Mod10_Padrao(StrDigitavel)
'    StrDigitavel = Mid(StrDigitavel, 1, 3) & StrDV & Mid(StrDigitavel, 4)
'    'PARTE 1
'    Pos = 1: Linha1 = Mid(StrDigitavel, Pos, 11) & " " & Codebar.Mod10_Padrao(Mid(StrDigitavel, Pos, 11))
'    'PARTE 2
'    Pos = Pos + 11: Linha2 = Mid(StrDigitavel, Pos, 11) & " " & Codebar.Mod10_Padrao(Mid(StrDigitavel, Pos, 11))
'    'PARTE 3
'    Pos = Pos + 11: Linha3 = Mid(StrDigitavel, Pos, 11) & " " & Codebar.Mod10_Padrao(Mid(StrDigitavel, Pos, 11))
'    'PARTE 4
'    Pos = Pos + 11: Linha4 = Mid(StrDigitavel, Pos, 11) & " " & Codebar.Mod10_Padrao(Mid(StrDigitavel, Pos, 11))
'
'    Linha = Linha1 & " " & Linha2 & " " & Linha3 & " " & Linha4
'
'    Codebar.Imp_Codbar 7, 0, 0&, 1, 50, 0, "Times New Roman", 12, 0, 0, 30, 0, 0, StrDigitavel, ObjetoPicture
'    GeraCodBarra = Linha
'End Function
'

Public Sub ImprimeDam(Report As VSRelatorio, CodPagamento As String, InscMunicipal As String, RazaoSocial As String, _
CgcCpf As String, EnderecoContribuinte As String, InscCadastral As String, EnderecoImovel As String, CodTributo As String, _
Sigla As String, NomeImposto As String, Exercicio As String, Parcela As String, Conta As Integer, DataVencimento As String, _
BaseCalculo As String, ValorTributo As String, ValorMulta As String, ValorJuros As String, TaxaServico As Double, _
Desconto As String, AtividadeEconomica As String, Observacao As String, Optional Picture As Object, Optional NotaInicial As String, _
Optional NotaFinal As String, Optional MaterialICMS As String, Optional ValorMetro As Double, Optional TaxaParcela As Double, _
Optional AreaTotal As Double, Optional AreaConstruida As Double, Optional ValorTerreno As Double, Optional ValorEdificacao As Double, Optional Zona As Integer, Optional Destino As TipoDestinoImpressao = tdiTela)
   On Error GoTo Trata
    Dim rs As VSRecordset
    Dim Sql As String
    Dim MaxCotas As Byte
    Dim CodBarra As New CodigoDeBarra
    Dim a As Byte
    Dim DataDesconto As String
    Dim CgcPref As String
    Dim LinhaDigitavel As String
    Dim InscricaoSP As String
    InscMunicipal = Trim(InscMunicipal)
    InscCadastral = Trim(InscCadastral)
    DataDesconto = Temp.PegaParametro(Bdados, "DATA DESCONTO")
    With Report
        ' gerar o arquivo bmp da barra antes de chamar o rpt pois somente assim irá sair com o bmp correto
        
        If Not .DefinirArquivo(Bdados, App.Path & "\TdamBarra.rpt") Then Exit Sub
        InscricaoSP = IIf(Trim(InscCadastral) <> "", Trim(InscCadastral), Trim(InscMunicipal))
        InscricaoSP = TiraPic(TiraPic(TiraPic(InscricaoSP, "."), "-"), "/")
        If Len(InscricaoSP) = 10 Then
            InscricaoSP = Left(InscMunicipal, 2) & "." & Mid(InscMunicipal, 3, 3) & "." & Mid(InscMunicipal, 6, 3) & Right(InscMunicipal, 3)
            .Formulas "endereco", EnderecoContribuinte
        Else
            If Len(InscricaoSP) = 9 Or Len(InscricaoSP) = 15 Then
                InscricaoSP = IIf(Len(InscCadastral) = 15, Left(InscCadastral, 2) & "." & Mid(InscCadastral, 3, 2) & "." & Mid(InscCadastral, 5, 4) & "." & Mid(InscCadastral, 9, 4) & "." & Right(InscCadastral, 3), InscCadastral)
                .Formulas "endereco", EnderecoImovel
            Else
                Dim Objeto As Object
                .Formulas "endereco", EnderecoContribuinte
            End If
        End If
        .Formulas "Inscricao", InscricaoSP
        .Formulas "nome", RazaoSocial
        .Formulas "cpf/cnpj", CgcCpf
        .Formulas "CodigoTributo", CodTributo
        .Formulas "exercicio", IIf(Len(Exercicio) = 4, Exercicio, Left(Exercicio, 2) & "/" & Right(Exercicio, 4))
        .Formulas "BASECALCULO", Format(Nvl(BaseCalculo, 0), Const_Monetario)
        .Formulas "datavencimento", DataVencimento
        .Formulas "ValorTributo", Format((CDbl(Nvl(ValorTributo, 0))), Const_Monetario)
        .Formulas "ValorMulta ", Format(ValorMulta, Const_Monetario)
        .Formulas "ValorJuros ", Format(ValorJuros, Const_Monetario)
        .Formulas "TaxaExpediente ", Format(TaxaServico, Const_Monetario)
        .Formulas "ValorTotal ", Format((CDbl(Nvl(ValorTributo, 0)) + TaxaServico) * (100 - Nvl(Desconto, 0)) / 100 + CDbl(Nvl(ValorMulta, 0)) + CDbl(Nvl(ValorJuros, 0)), Const_Monetario)
        If Sigla <> Imposto.NomeTributo(ttr_IPTU) Then
            .Formulas "OBSERVACAO ", UCase(Trim(Observacao))
        Else
'******************COLETA DE DADOS******************
        Sql = "select tdi_tco_cod_componente,tdi_valor_item from tab_detalhe_imovel where tdi_tim_ic='" & InscCadastral & _
            "' and (tdi_tco_cod_componente IN (110,108))"
        If Bdados.AbreTabela(Sql, rs) Then
            rs.MoveFirst
            Do While Not rs.EOF
                If rs(0) = 110 Then
                    AreaTotal = Nvl("" & rs(1), 0)
                ElseIf rs(0) = 108 Then
                    AreaConstruida = Nvl(rs(1), 0)
                End If
                rs.MoveNext
            Loop
        End If
        Bdados.FechaTabela rs
        If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) <> 1 Then
            Sql = "select tvl_valor from tab_valor_terreno where tvl_tlg_cod_logradouro=(" & _
                " select tim_tlg_cod_logradouro from tab_imovel where tim_ic='" & InscCadastral & "')"
            If Bdados.AbreTabela(Sql, rs) Then
                ValorMetro = Format(TrocaPic(rs(0), ".", ","), Const_Monetario)
            End If
        End If
        If Conta = 4 Then
            Sql = "Select tpp_max_cotas from tab_parametro_parcelamento"
             If Bdados.AbreTabela(Sql, rs) Then
                 MaxCotas = rs!tpp_max_cotas
             End If
             Bdados.FechaTabela rs
        Else
            MaxCotas = 1
        End If
        
        Sql = "select TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC," & _
            " tim_valor ,TIM_SITUACAO_LOTE   from TAB_IMOVEL where tim_ic='" & InscCadastral & "'"
        If Bdados.AbreTabela(Sql, rs) Then
            ValorTerreno = Nvl("" & rs!TIM_VALOR_TERRENO, 0)
            ValorEdificacao = Nvl("" & rs!TIM_VALOR_EDIFIC, 0)
            BaseCalculo = "" & rs!tim_valor
        End If
'******************FIM DE COLETA DE DADOS******************
            If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 Then
                .Formulas "ObsMaterial", Trim(Observacao)
            Else
                .Formulas "ObsMaterial", "Valor m2 Terreno: R$" & Format(ValorMetro, Const_Monetario) & " - Valor Venal: R$" & Format(BaseCalculo, Const_Monetario) '& " - TSU: R$" & Format(TaxaServico, Const_Monetario) & " - Valor IPTU: R$" & Format(MaxCotas * (CDbl(Nvl(ValorTributo, 0)) * (100 + Nvl(Desconto, 0)) / 100) - TaxaParcela, Const_Monetario) & _
                IIf(Not IsNumeric(Parcela) And Nvl(Desconto, 0) > 0, Space(32) & "DESCONTO DE " & Desconto & "% EM COTA ÚNICA", "")
            End If
        End If
         If Trim(NotaInicial) <> "" And Trim(NotaFinal) <> "" Then
            .Formulas "NUM_NOTAS", "Nº da Nota Inicial: " & NotaInicial & " - Nº da Nota Final: " & NotaFinal
         ElseIf Sigla = Imposto.NomeTributo(ttr_IPTU) Then
            If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 And Conta = 1 Then
                .Formulas "NUM_NOTAS", "Valor Terreno: R$" & Format(ValorTerreno, Const_Monetario) & "     -       Valor Edificacao: R$" & Format(ValorEdificacao, Const_Monetario)
            ElseIf Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) <> 1 Then
                .Formulas "NUM_NOTAS ", "Área Total do Imóvel: " & Format(AreaTotal, Const_Monetario) & "m2     -       Área Total Construída: " & Format(AreaConstruida, Const_Monetario) & "m2"
            End If
         End If
        .Formulas "DESCTRIBUTO ", NomeImposto
        .Formulas "Observacao", Trim(Observacao)
        .Formulas "Parcela ", Nvl(Parcela, "ÚNICA")
        .Formulas "DEDUCAO", Format(CDbl(CDbl(Nvl(ValorTributo, 0)) + TaxaServico) * (Nvl(Desconto, 0) / 100), Const_Monetario)
        .Formulas "PREFEITURA", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
        CgcPref = Temp.PegaParametro(Bdados, "CGC CLIENTE")
        .Formulas "CgcPrefeitura", "CNPJ " & Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2)
       .Formulas "EMISSAO", Date 'Imposto.BuscaDataGeracaoDam(CodPagamento)
        If Nvl(MaterialICMS, 0) > 0 Then
            .Formulas "ObsMaterial", "Valor de material sujeito ao ICMS: R$" & Format(MaterialICMS, Const_Monetario)
        End If
        If Sigla = Imposto.NomeTributo(ttr_IPTU) Then
            If Nvl(Temp.PegaParametro(Bdados, "TIPO IPTU"), 0) = 1 And Conta = 1 Then
                .Formulas "ObsMaterial ", " Valor Venal: R$" & Format(BaseCalculo, Const_Monetario) & Space(5) & " TSU: R$" & Format(TaxaServico, Const_Monetario) & " - Valor IPTU: R$" & Format(ValorTributo, Const_Monetario)
            End If
        End If
        If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
            LinhaDigitavel = CodBarra.CriaLinhaDigitavelCBR(InscricaoSP, CodTributo, (CDbl(Nvl(ValorTributo, 0)) * (100 + Nvl(Desconto, 0)) / 100) + CDbl(Nvl(ValorMulta, 0)) + CDbl(Nvl(ValorJuros, 0)) + CDbl(TaxaServico), Exercicio, Picture, DataVencimento, Nvl(Parcela, 0), CodPagamento)
        Else
            LinhaDigitavel = CodBarra.CriaLinhaDigitavel(InscricaoSP, CodTributo, (CDbl(Nvl(ValorTributo, 0)) * (100 + Nvl(Desconto, 0)) / 100) + CDbl(Nvl(ValorMulta, 0)) + CDbl(Nvl(ValorJuros, 0)) + CDbl(TaxaServico), Exercicio, DataVencimento, Nvl(Parcela, 0))
        End If
        .Formulas "LinhaDigitavel", LinhaDigitavel
        .Formulas "VT_CodBarra", CodBarra.LinhaBarraGerada
        .Titulo = "Documento de Arrecadação Municipal - DAM"
        .CopiasDetalhes = 3
        .Arvore = False
        If Destino = tdiTela Then
            .Visualizar
        Else
            .Imprimir
        End If
        DoEvents
    End With
    Set Rpt = Nothing
    Exit Sub

Trata:
    If Err.Number = 20515 Or Err.Number = 3265 Then
        Report.Formulas "OBSERVACAO ", ""
        Resume
    End If
    Avisa "O DAM não será impresso. Informe o erro a seguir ao suporte."
    Avisa Err.Number & " - " & Err.Description
End Sub


