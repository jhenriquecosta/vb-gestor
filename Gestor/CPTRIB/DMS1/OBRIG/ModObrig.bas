Attribute VB_Name = "ModObrig"
Public Sub ImprimeSelecionado(Lista As Object, Razao As String, Endereco As String, Optional ComNovaData As Boolean = False, Optional NovaData As String, Optional DetinoDAM As TipoDestinoImpressao = tdiTela)
    Dim CpfCgc As String
    Dim declaracao As New cDeclaracao
    Dim Desconto As Double
    Dim Exer As String
    Dim BaseCalculo As String
    Dim Obs As String
    Dim Imposto As New VSImposto
    Dim Lote As New BCI
    Dim Cobranca As New VSCobranca
    Dim NovoJuro As String
    Dim NovaMulta As String
    Dim Conta As New ContaCorrente
    Dim Obrig As New Obrigacao
    Dim NotaInicial As String
    Dim NotaFinal As String
    Obs = ""
    Obs = Util.Entrada("Observacao", "Impressão de Parcela.")
    With Lista
  
        Exer = IIf(Len(.SelectedItem.SubItems(4)) = 4, .SelectedItem.SubItems(4), Right(.SelectedItem.SubItems(4), 2) & Left(.SelectedItem.SubItems(4), 4))
        
        If .SelectedItem.SubItems(3) = Imposto.NomeTributo(ttr_IPTU) Or .SelectedItem.SubItems(3) = Imposto.NomeTributo(ttr_ITU) Then
            If CDate(.SelectedItem.SubItems(5)) >= Date Then
                Desconto = .SelectedItem.SubItems(6) * (CInt(Nvl(Bdados.BuscaCodigo("Select TPP_DESCONTO from TAB_PARAMETRO_PARCELA_IPTU WHERE tpp_parcela=0 and TPP_ANO=" & Exer), 0)) / 100)
            End If
        End If
        
        If .SelectedItem.SubItems(3) = Imposto.NomeTributo(ttr_ISSQN) Then
            If declaracao.Buscar(.SelectedItem.SubItems(1), .SelectedItem.SubItems(4)) Then
                BaseCalculo = declaracao.Itens(4).Valor
                NotaInicial = declaracao.Itens(2).Valor
                NotaFinal = declaracao.Itens(3).Valor
'                Obs = "Nota Inicial: " & declaracao.Itens(1).Valor & "     Nota Final: " & declaracao.Itens(2).Valor
            Else
                BaseCalculo = 0
                Obs = ""
            End If
        ElseIf .SelectedItem.SubItems(3) = Imposto.NomeTributo(ttr_IPTU) Then
            BaseCalculo = Format(Lote.ValorVenal(.SelectedItem.SubItems(1)), Const_Monetario)
        Else
            BaseCalculo = .SelectedItem.SubItems(6)
        End If
        Set declaracao = Nothing
        If (Not ComNovaData) Or (NovaData = .SelectedItem.SubItems(5)) Then
            NovaData = .SelectedItem.SubItems(5)
            NovoJuro = .SelectedItem.SubItems(7)
            NovaMulta = .SelectedItem.SubItems(8)
        Else
            NovoJuro = Conta.CalculaValoresJurosAvulsos(.SelectedItem.SubItems(11), .SelectedItem.SubItems(4), EtcCreditoTributario, NovaData, .SelectedItem.SubItems(5), .SelectedItem.SubItems(6))
            NovaMulta = Conta.CalculaValoresMultaAvulsos(.SelectedItem.SubItems(11), .SelectedItem.SubItems(4), EtcCreditoTributario, NovaData, .SelectedItem.SubItems(5), .SelectedItem.SubItems(6))
        End If
        Dim PicBarra As Object
        Cobranca.ImprimeDam Rpt, "", .SelectedItem.SubItems(1), Razao, CpfCgc, Endereco, .SelectedItem.SubItems(1), Endereco, .SelectedItem.SubItems(11), .SelectedItem.SubItems(2), _
            .SelectedItem.SubItems(12), Exer, 0, 4, NovaData, BaseCalculo, .SelectedItem.SubItems(6) - Desconto, NovaMulta, NovoJuro, 0, 0, "", Obs, _
             PicBarra, NotaInicial, NotaFinal, , , , , , , , , tdiTela
    End With
End Sub
