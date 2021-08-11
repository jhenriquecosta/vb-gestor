Attribute VB_Name = "TMODCOMUM"

Option Explicit
Public User As String
Public Bdados As New VSDados
Public Edita As New VSTexto
Public RPT As New VSRelatorio
Public Util As New VSUtil
Public Instala As New VSInstala
Public Seguranca As New VSSeguranca
Public Temp As New VSTemp
Public Aplicacoes As Aplicacoes
Public Imposto As New VSImposto
Public TempContrib As String
Public MUN As String
Public CODMUN As String
Public Ic As String

'541 3312
'***********************Area de Instanciacoes dos Modulos*******************************
Public AplicacoesVTFuncoes As New VsTFuncoes.VsTFuncAplicacoes
'Public AplicacoesDecl As New DecDll.Aplicacoes
'***********************---------------------------------*******************************
Public Const Const_Extrato As String = "EXTRATO"
Public Const Const_Notificacao As String = "NOTIFICA"
Public Const Const_Monetario As String = "#,##0.00"
Public Const Const_ImAvulso As String = "11000000-00"
Public Const Const_Obrig As String = "63"
Public Const Const_NaoParcelaveis As String = "3,4,6,7"
Public Const Const_NaoPagos As String = "2,4,5"
Public Const Const_SenhaEMail As String = "CIAPENVIODEC"

'=======================================================
' Constantes do módulo Arquivos.cls                                                                      =
'=======================================================
Public Const InicioValorTarifa = 94
Public Const TamanhoValorTarifa = 6
Public Const InicioValorRecebido = 82
Public Const TamanhoValorRecebido = 11
Public Const InicioAgenciaContaDigito = 2
Public Const TamanhoAgenciaContaDigito = 21
'NumDocumento
Public Const InicioCodigoPagamento = 74
Public Const TamanhoCodigoPagamento = 7
Public Const InicioVersao = 80
Public Const TamanhoVersao = 2
Public Const InicioCodigoRemessa = 2
Public Const TamanhoCodigoRemessa = 1
Public Const InicioCodigoBanco = 43
Public Const TamanhoCodigoBanco = 3
Public Const InicioDataGeracao = 66
Public Const TamanhoDataGeracao = 8
Public Const InicioDataPagamento = 22
Public Const TamanhoDataPagaemnto = 8
Public Const InicioNumArquivo = 74
Public Const TamanhoNumArquivo = 6 '?
'Public const  InicioTotalArquivos
Public Const InicioValorTotal = 8
Public Const TamanhoValorTotal = 18
Public Const InicioTotalArquivos = 2
Public Const TamanhoTotalArquivos = 6
'===========X===============================================


Public Inscri As String
Private Const cteEspacamentoColunas As Integer = 3
Private Const cteLinhasCabecalho As Integer = 11
Private Const cteLinhasRodape As Integer = 5

Public Sistema As String
Public Desc_Form As String
Public Cod_sis As String

Public Enum enuTipoCampo
    tipTexto = 5
    tipData = 1
    tipInteiro = 2
    tipMoeda = 3
    tipFloat = 4
End Enum
Public Enum enuAlinhamentoCampo
    aliEsquerda = 0
    aliCentro = 1
    aliDireita = 2
End Enum

Public Enum StatusLivroAforamento
    slaLivroAberto = 1
    slaLivroFechado = 0
End Enum

Public Enum SituacaoImovel
    siNaoAforado = 0
    siAforado = 1
    siTransferencia = 2
    siReimpressao = 3
End Enum

Public Enum StatusGradeLote
    sglAberto = 1
    sglFechado = 2
    sglConferido = 3
End Enum

Public Enum CalcUFM
    Converete_UFM = 0
    Converete_Real = 1
End Enum
Public Const ettParcelada = 2
Public Const ettDividaAtiva = 3
Sub Main()
    Set Aplicacoes = New Aplicacoes
'    Set Imposto = New VSImposto
End Sub

Public Sub VisualizarActiveReport(Relatorio As Object, Dados As Object, Sql As String)
    With Relatorio
        .DataControl1.ConnectionString = Dados.Conexao.ConnectionString
        .DataControl1.Source = Sql
        .Show
    End With
End Sub

Public Function DataServidor() As Date

    If Bdados.Conexao.FormatoBanco = SQLServer Then
        If Bdados.AbreTabela("SELECT GETDATE() AS DataAtual") Then
            DataServidor = Bdados.Tabela(0).Value
        Else
            DataServidor = Date
        End If
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        DataServidor = Date
    End If
    Bdados.FechaTabela
End Function


Public Function BuscaCodigo(Tabela As String) As Long
    Dim Rs As Object
    Dim ConsegAbrir As Boolean
    
    BuscaCodigo = 0
    If Bdados.AbreTabela(Tabela, Rs) Then
        BuscaCodigo = "" & Rs(0)
    End If
    Bdados.FechaTabela Rs
    
End Function

Public Function BuscaIndiceCombo(Combo As ComboBox, Tabela As String, CampoCodigo As String, CampoNome As String, Indice As Integer) As Integer
    Dim Sql As String
    Dim Rs As Object
    Dim i As Integer
    
    Sql = "SELECT " & CampoNome & " from " & Tabela & " where " & CampoCodigo & "=" & Indice
    If Bdados.AbreTabela(Sql, Rs) Then
        For i = 0 To Combo.ListCount - 1
            Combo.ListIndex = i
            If Combo.Text = Rs(0) Then
                BuscaIndiceCombo = i
                Bdados.FechaTabela Rs
                Exit Function
            End If
        Next
    End If
    BuscaIndiceCombo = -1
    Bdados.FechaTabela Rs
End Function

Public Function UltimoDiaDoMes(data As Date) As Date
    UltimoDiaDoMes = DateAdd("d", -1, "01/" & Mid(DateAdd("m", 1, data), 4))
End Function

Function PreencheEspaco(Texto, Tamanho As Byte) As String
        PreencheEspaco = Texto & Space(Tamanho - Len(Trim(Texto)))
End Function

Public Function CepCliente() As String
    Dim retorno As String
    'retorno = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
    'If TiraTudo(retorno) = "" Then
        retorno = Temp.PegaParametro(Bdados, "CEP MUNICIPIO")
    'End If
    CepCliente = retorno
End Function

Public Function FuncaoReal(campo As String) As String
    FuncaoReal = "cast(" & campo & " as decimal(14,2))"
End Function

Public Sub AtualizaUF(Combo As ComboBox)
   Combo.Clear
   Combo.AddItem "MA"
   Combo.AddItem "AC"
   Combo.AddItem "AM"
   Combo.AddItem "AP"
   Combo.AddItem "AL"
   Combo.AddItem "BA"
   Combo.AddItem "CE"
   Combo.AddItem "DF"
   Combo.AddItem "ES"
   Combo.AddItem "GO"
   Combo.AddItem "MG"
   Combo.AddItem "MS"
   Combo.AddItem "MT"
   Combo.AddItem "PA"
   Combo.AddItem "PB"
   Combo.AddItem "PE"
   Combo.AddItem "PI"
   Combo.AddItem "PR"
   Combo.AddItem "SC"
   Combo.AddItem "SE"
   Combo.AddItem "SP"
   Combo.AddItem "RJ"
   Combo.AddItem "RN"
   Combo.AddItem "RO"
   Combo.AddItem "RR"
   Combo.AddItem "RS"
   Combo.AddItem "TO"
End Sub

Public Sub AtualizaCabecalho(Lista As Object, Optional Titulo As String)
    With Lista
        .CabecalhoCliente = Temp.PegaParametro(Bdados, "CLIENTE")
        .CabecalhoDepartamento = Temp.PegaParametro(Bdados, "SETOR")
        .CabecalhoEstado = Temp.PegaParametro(Bdados, "ESTADO")
        .CabecalhoSecretaria = Temp.PegaParametro(Bdados, "SEMFAZ")
        .CabecalhoTitulo = Titulo
        .RodapeUsuario = Aplicacoes.Usuario
    End With
End Sub

Public Function ListIndexDe(Combo As Object, Texto As String) As Integer
    Dim i As Integer
    For i = 0 To Combo.ListCount
        If Combo.List(i) = Texto Then
            ListIndexDe = i
            Exit Function
        End If
    Next
    ListIndexDe = -1
End Function

Public Function DescricaoComponente(Grupo As Integer, Item As Integer)
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "Select tco_descricao_componente from tab_componente_avancado where tco_cod_componente = " & Item & " and tco_grupo =" & Grupo
    If Bdados.AbreTabela(Sql, Rs) Then
        DescricaoComponente = Rs!tco_descricao_componente
    End If
End Function


Public Function NomeDe(Oque As Byte, Codigo As String) As String
    Dim Sql As String
    Sql = "SELECT "
    Select Case Oque
        Case 0
            Sql = Sql & "tco_descricao_componente FROM Tab_Componente where tco_cod_componente = "
            
        Case 1
            Sql = Sql & "tba_nome FROM tab_bairro WHERE TBA_COD_BAIRRO = "
        Case 2
            Sql = Sql & "ttl_nome from tab_tipo_logr where ttl_cod_tip_logr="
    End Select
    
    Sql = Sql & Codigo & IIf(Oque = 1, " and tba_tmu_cod_municipio=" & Temp.PegaParametro(Bdados, "MUNICIPIO"), "")
    
    Dim Rs As Object
    If Bdados.AbreTabela(Sql, Rs) Then
        NomeDe = Rs(0)
    End If
    Bdados.FechaTabela Rs
    
End Function

Public Function CarregaEnderecoImovel(Ic As String, Optional ByRef Endereco As Object, Optional ByRef Im As Object) As Boolean
    Dim Sql As String
    Dim Rs As Object
    
    If Trim(Ic) = "" Then Exit Function
    Sql = "select ttl_nome,tlg_nome,tba_nome,tim_numero," & _
    "tim_tci_im ,TIM_SITUACAO_LOTE   from TAB_IMOVEL,TAB_BAIRRO," & _
    " TAB_LOGRADOURO,TAB_TIPO_LOGR " & _
    " where tim_ic='" & Ic & _
    "' AND tim_tlg_cod_logradouro = " & _
    " TAB_LOGRADOURO.tlg_cod_logradouro AND tlg_ttl_cod_tip_logr = ttl_cod_tip_logr AND " & _
    " tlg_tmu_cod_municipio=" & Temp.PegaParametro(Bdados, "MUNICIPIO") & " AND TBA_TMU_COD_MUNICIPIO =" & _
    Temp.PegaParametro(Bdados, "MUNICIPIO") & "  AND TIM_TBA_COD_BAIRRO = TBA_COD_BAIRRO"
    If Bdados.AbreTabela(Sql, Rs) Then
        If "" & Rs!TIM_SITUACAO_LOTE = 1 Then
            Util.Informa "Imóvel desativado."
            Exit Function
        End If
        If Not Endereco Is Nothing Then Endereco = Rs(0) & " " & Rs(1) & " " & Rs(2) & " " & Rs(3)
        If Not Im Is Nothing Then Im = Rs!tim_tci_im
        CarregaEnderecoImovel = True
    Else
        Util.Informa "Imovel não cadastrado."
    End If
    Bdados.FechaTabela Rs

End Function
    
Public Function BuscaNaGeral(Tabela As String, Registro As Integer) As String
    Dim Sql As String
    Dim Rs As Object
    
    Sql = "select TGE_NOME from tab_geral where TGE_TIPO = (SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME ='" & Tabela & "' ) and TGE_CODIGO =" & Registro
    BuscaNaGeral = ""
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaNaGeral = Rs(0)
    End If
    Rs.Fechar
End Function

Public Function BuscaContribuinte(Inscricao As String, Razao As Object, Optional Endereco As Object, Optional Documento As String, Optional TipoInscricao As TipoInscricaoObrigacao) As String
    
    Dim Obrig As New Obrigacao
    'Dim Obrig As Object
   ' Set Obrig = CreateObject("VsTFuncoes.Obrigacao")
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(Trim(Inscricao)) = 10 And IsNumeric(Inscricao) Then Inscricao = Format(Inscricao, "00000000-00")
    End If
    BuscaContribuinte = Obrig.BuscaSujeitoPassivoObrigacao(Inscricao, Razao, Endereco, Documento, TipoInscricao)
    
End Function

Public Sub MenuPopUp(Formulario As Object, Grid As Object, BotaoMouse As Integer, MenuGeral As Object, MenuPopUp As Object, Caption As String)
    If Not Grid.SelectedItem Is Nothing Then
        If BotaoMouse = 2 Then
            MenuPopUp.Caption = Caption
            Formulario.PopupMenu MenuGeral
        End If
    End If
End Sub

Public Function CNPJCliente() As String
    Dim CgcPref As String
    CgcPref = UCase(Temp.PegaParametro(Bdados, "CGC CLIENTE"))
    CgcPref = Edita.TiraTudo(CgcPref)
    CNPJCliente = Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2)
End Function

Public Function MunicipioCliente() As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim CodMunicipio As String
    
    CodMunicipio = Temp.PegaParametro(Bdados, "MUNICIPIO")
    
    Sql = "SELECT TAB_MUNICIPIO.TMU_NOME" _
        & " FROM TAB_MUNICIPIO" _
        & " Where (TAB_MUNICIPIO.TMU_COD_MUNICIPIO = " & CodMunicipio & ")"
    
    If Bdados.AbreTabela(Sql, Rs) Then
        MunicipioCliente = "" & Rs!TMU_NOME
    End If
End Function

Public Function MunicipioUFCliente() As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim CodMunicipio As String
    
    CodMunicipio = Temp.PegaParametro(Bdados, "MUNICIPIO")
    
    Sql = "SELECT TAB_MUNICIPIO.TMU_NOME, TAB_UF.TUF_UF" _
        & " FROM TAB_MUNICIPIO INNER JOIN" _
        & " TAB_UF ON TAB_MUNICIPIO.TMU_TUF_COD_UF = TAB_UF.TUF_COD_UF" _
        & " Where (TAB_MUNICIPIO.TMU_COD_MUNICIPIO = " & CodMunicipio & ")"
    
    If Bdados.AbreTabela(Sql, Rs) Then
        MunicipioUFCliente = "" & Rs!TMU_NOME & " - " & Rs!TUF_UF
    End If
End Function

Public Function MunicipioEstadoCliente() As String
    'EX.: BELO HORIZONTE - MINAS GERAIS
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim CodMunicipio As String
    
    CodMunicipio = Temp.PegaParametro(Bdados, "MUNICIPIO")
    
    Sql = "SELECT TAB_MUNICIPIO.TMU_NOME, TAB_UF.TUF_NOME" _
        & " FROM TAB_MUNICIPIO INNER JOIN" _
        & " TAB_UF ON TAB_MUNICIPIO.TMU_TUF_COD_UF = TAB_UF.TUF_COD_UF" _
        & " Where (TAB_MUNICIPIO.TMU_COD_MUNICIPIO = " & CodMunicipio & ")"
    
    If Bdados.AbreTabela(Sql, Rs) Then
        MunicipioEstadoCliente = "" & Rs!TMU_NOME & " - " & Rs!TUF_NOME
    End If
    
End Function

Public Function Extenso(ByVal Valor As _
    Double, ByVal MoedaPlural As _
    String, ByVal MoedaSingular As _
    String) As String
    Dim StrValor As String, Negativo As Boolean
    Dim Buf As String, Parcial As Integer
    Dim Posicao As Integer, Unidades
    Dim Dezenas, Centenas, PotenciasSingular
    Dim PotenciasPlural

    Negativo = (Valor < 0)
    Valor = Abs(CDec(Valor))
    If Valor Then
        Unidades = Array(vbNullString, "Um", "Dois", _
        "Três", "Quatro", "Cinco", _
        "Seis", "Sete", "Oito", "Nove", _
        "Dez", "Onze", "Doze", "Treze", _
        "Quatorze", "Quinze", "Dezesseis", _
        "Dezessete", "Dezoito", "Dezenove")
        Dezenas = Array(vbNullString, vbNullString, _
        "Vinte", "Trinta", "Quarenta", _
        "Cinqüenta", "Sessenta", "Setenta", _
        "Oitenta", "Noventa")
        Centenas = Array(vbNullString, "Cento", _
        "Duzentos", "Trezentos", _
        "Quatrocentos", "Quinhentos", _
        "Seiscentos", "Setecentos", _
        "Oitocentos", "Novecentos")
        PotenciasSingular = Array(vbNullString, " Mil", _
        " Milhão", " Bilhão", _
        " Trilhão", " Quatrilhão")
        PotenciasPlural = Array(vbNullString, " Mil", _
        " Milhões", " Bilhões", _
        " Trilhões", " Quatrilhões")
    
        StrValor = Left(Format(Valor, String(18, "0") & _
        ".000"), 18)
        For Posicao = 1 To 18 Step 3
            Parcial = Val(Mid(StrValor, Posicao, 3))
            If Parcial Then
                If Parcial = 1 Then
                    Buf = "Um" & PotenciasSingular((18 - _
                    Posicao) \ 3)
                ElseIf Parcial = 100 Then
                    Buf = "Cem" & PotenciasSingular((18 - _
                    Posicao) \ 3)
                Else
                    Buf = Centenas(Parcial \ 100)
                    Parcial = Parcial Mod 100
                    If Parcial <> 0 And Buf <> vbNullString Then
                        Buf = Buf & " e "
                    End If
                    If Parcial < 20 Then
                        Buf = Buf & Unidades(Parcial)
                    Else
                        Buf = Buf & Dezenas(Parcial \ 10)
                        Parcial = Parcial Mod 10
                    If Parcial <> 0 And Buf <> vbNullString Then
                        Buf = Buf & " e "
                    End If
                    Buf = Buf & Unidades(Parcial)
                End If
                Buf = Buf & PotenciasPlural((18 - Posicao) \ 3)
            End If
            If Buf <> vbNullString Then
                If Extenso <> vbNullString Then
                    Parcial = Val(Mid(StrValor, Posicao, 3))
                    If Posicao = 16 And (Parcial < 100 Or _
                        (Parcial Mod 100) = 0) Then
                        Extenso = Extenso & " e "
                    Else
                        Extenso = Extenso & ", "
                    End If
                End If
                Extenso = Extenso & Buf
            End If
        End If
        Next
        If Extenso <> vbNullString Then
            If Negativo Then
                Extenso = "Menos " & Extenso
            End If
            If Int(Valor) = 1 Then
                Extenso = Extenso & " " & MoedaSingular
            Else
                Extenso = Extenso & " " & MoedaPlural
            End If
        End If
        Parcial = Int((Valor - Int(Valor)) * _
        100 + 0.1)
        If Parcial Then
            Buf = Extenso(Parcial, "Centavos", _
            "Centavo")
            If Extenso <> vbNullString Then
                Extenso = Extenso & " e "
            End If
            Extenso = Extenso & Buf
        End If
    End If
End Function

Public Function Checa_Obrigacao1_2(Im As String) As Boolean
    Dim Sql                                            As String
    
    Sql = " SELECT TOC_COD_OBRIGACAO"
    Sql = Sql & " From TAB_OBRIGACAO_CONTRIBUINTE, TAB_IMPOSTO, VIS_STATUS_OBRIGACAO, VIS_INSCRICAO"
    Sql = Sql & " Where TOC_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
    Sql = Sql & " AND TOC_INSCRICAO = VIN_INSCRICAO"
    Sql = Sql & " AND  TOC_STATUS_OBRIGACAO = TGE_CODIGO"
    Sql = Sql & " AND TOC_STATUS_OBRIGACAO in (1,2)"
    Sql = Sql & " AND VIN_INSCRICAO = " & Bdados.Converte(Im, tctexto)
    Checa_Obrigacao1_2 = Bdados.AbreTabela(Sql)
End Function
Public Function Calcula_UFM(Valor As Variant, Conversao As CalcUFM) As Variant
    Dim Sql                                 As String
    Dim Rs                                    As VSRecordset
    Dim UFM                                 As Variant
    
    'Pego o valor do UFM em Real na tab_Geral
        UFM = Edita.TrocaPic(Temp.PegaParametro(Bdados, "UFM"), ".", ",")
        
    Select Case Conversao
        Case Converete_UFM
            Calcula_UFM = Valor / UFM
        Case Converete_Real
            Calcula_UFM = (Valor * UFM)
    End Select
End Function

Public Function Verifica_Credenciamento(Im As String, data As Date, Ret As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "SELECT  max(Tgc_validade)"
    Sql = Sql & " From Tab_Contribuinte, tab_grafica_credenciada"
    Sql = Sql & " Where tgc_tci_im = tci_im And tgc_status = 0"
    Sql = Sql & " and tgc_tci_im = " & Bdados.Converte(Imposto.FormataInscricao(Im, InscContrib), tctexto)
    Sql = Sql & " and Tgc_validade <" & Bdados.Converte(data, TCDataHora)
    If Bdados.AbreTabela(Sql) Then
        If Not IsNull(Bdados.Tabela(0)) Then
            Ret = Bdados.Tabela(0)
            Verifica_Credenciamento = True
        Else
            Verifica_Credenciamento = False
        End If
    End If
End Function



