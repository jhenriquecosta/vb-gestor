Attribute VB_Name = "TMODCOMUM"
Option Explicit


Public User As String
Public Bdados As New VSDados
Public Edita As New VSTexto
Public Rpt As New VSRelatorio
Public Util As New VSUtil
Public Instala As New VSInstala
Public Seguranca As New VSSeguranca
Public Temp As New VSTemp
Public Aplicacoes As Aplicacoes
Public Imposto As VsTFuncoes.VSImposto
'Public CodPagamento As Double
Public MUN As String
Public CODMUN As String

'***********************Area de Instanciacoes dos Modulos*******************************
Public AplicacoesVTFuncoes As New VsTFuncoes.VsTFuncAplicacoes
'***********************---------------------------------*******************************
Public Const Const_Extrato As String = "EXTRATO"
Public Const Const_Notificacao As String = "NOTIFICA"
Public Const Const_Monetario As String = "#,##0.00"
Public Const Const_ImAvulso As String = "11000000-00"
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



Private Const cteEspacamentoColunas As Integer = 3
Private Const cteLinhasCabecalho As Integer = 11
Private Const cteLinhasRodape As Integer = 5

Public Sistema As String
Public Desc_Form As String
Public Cod_Sis As String

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

Sub Main()
    Set Aplicacoes = New Aplicacoes
    Set Imposto = New VSImposto
End Sub

Public Function BuscaCodigo(Tabela As String) As Long
    Dim RS As VSRecordset
    Dim ConsegAbrir As Boolean
    
    BuscaCodigo = 0
    If Bdados.AbreTabela(Tabela, RS) Then
        BuscaCodigo = "" & RS(0)
    End If
    Bdados.FechaTabela RS
    
End Function

Public Function BuscaIndiceCombo(Combo As ComboBox, Tabela As String, CampoCodigo As String, CampoNome As String, Indice As Integer) As Integer
    Dim SQL As String
    Dim RS As VSRecordset
    Dim i As Integer
    
    SQL = "SELECT " & CampoNome & " from " & Tabela & " where " & CampoCodigo & "=" & Indice
    If Bdados.AbreTabela(SQL, RS) Then
        For i = 0 To Combo.ListCount - 1
            Combo.ListIndex = i
            If Combo.Text = RS(0) Then
                BuscaIndiceCombo = i
                Bdados.FechaTabela RS
                Exit Function
            End If
        Next
    End If
    BuscaIndiceCombo = -1
    Bdados.FechaTabela RS
End Function

Public Function UltimoDiaDoMes(Data As Date) As Date
    UltimoDiaDoMes = DateAdd("d", -1, "01/" & Mid(DateAdd("m", 1, Data), 4))
End Function

Function PreencheEspaco(Texto, Tamanho As Byte) As String
        PreencheEspaco = Texto & Space(Tamanho - Len(Trim(Texto)))
End Function

Public Function CepCliente() As String
    CepCliente = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
End Function

Public Function FuncaoReal(Campo As String) As String
    FuncaoReal = "cast(" & Campo & " as decimal(14,2))"
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

Public Function ListIndexDe(Combo As ComboBox, Texto As String) As Integer
    Dim i As Integer
    For i = 0 To Combo.ListCount
        If Combo.List(i) = Texto Then
            ListIndexDe = i
            Exit Function
        End If
    Next
    ListIndexDe = -1
End Function

Public Function NomeDe(Oque As Byte, Codigo As String) As String
    Dim SQL As String
    SQL = "SELECT "
    Select Case Oque
        Case 0
            SQL = SQL & "tco_descricao_componente FROM Tab_Componente WHERE tco_tmu_cod_municipio =" & Aplicacoes.Codigo_Municipio & " and tco_cod_componente = "
            
        Case 1
            SQL = SQL & "tba_nome FROM tab_bairro WHERE tba_tmu_cod_municipio =" & Aplicacoes.Codigo_Municipio & " and TBA_COD_BAIRRO = "
        Case 2
            SQL = SQL & "ttl_nome from tab_tipo_logr where ttl_cod_tip_logr="
    End Select
    
    SQL = SQL & Codigo & IIf(Oque = 1, " and tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio, "")
    
    Dim RS As VSRecordset
    If Bdados.AbreTabela(SQL, RS) Then
        NomeDe = RS(0)
    End If
    Bdados.FechaTabela RS
    
End Function

Public Function CarregaEnderecoImovel(Ic As String, Optional ByRef Endereco As Object, Optional ByRef Im As Object) As Boolean
    Dim SQL As String
    Dim RS As VSRecordset
    
    If Trim(Ic) = "" Then Exit Function
    SQL = "select ttl_nome,tlg_nome,tba_nome,tim_numero," & _
    "tim_tci_im ,TIM_SITUACAO_LOTE   from TAB_IMOVEL,TAB_BAIRRO," & _
    " TAB_LOGRADOURO,TAB_TIPO_LOGR " & _
    " where tim_ic='" & Ic & _
    "' AND tim_tlg_cod_logradouro = " & _
    " TAB_LOGRADOURO.tlg_cod_logradouro AND tlg_ttl_cod_tip_logr = ttl_cod_tip_logr AND " & _
    " tlg_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio & " AND TBA_TMU_COD_MUNICIPIO =" & _
    Aplicacoes.Codigo_Municipio & "  AND TIM_TBA_COD_BAIRRO = TBA_COD_BAIRRO"
    If Bdados.AbreTabela(SQL, RS) Then
        If "" & RS!TIM_SITUACAO_LOTE = 1 Then
            Informa "Imóvel desativado."
            Exit Function
        End If
        If Not Endereco Is Nothing Then Endereco = RS(0) & " " & RS(1) & " " & RS(2) & " " & RS(3)
        If Not Im Is Nothing Then Im = RS!tim_tci_im
        CarregaEnderecoImovel = True
    Else
        Informa "Imovel não cadastrado."
    End If
    Bdados.FechaTabela RS

End Function
    
Public Function BuscaNaGeral(Tabela As String, Registro As Integer) As String
    Dim SQL As String
    Dim RS As VSRecordset
    
    SQL = "select TGE_NOME from tab_geral where TGE_TIPO = (SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME ='" & Tabela & "' ) and TGE_CODIGO =" & Registro
    BuscaNaGeral = ""
    If Bdados.AbreTabela(SQL, RS) Then
        BuscaNaGeral = RS(0)
    End If
    RS.Fechar
End Function

Public Function BuscaContribuinte(Inscricao As String, Razao As Object, Endereco As Object, Optional Documento As String) As String
    Dim Obrig As New Obrigacao
    If Len(Trim(Inscricao)) = 10 And IsNumeric(Inscricao) Then Inscricao = Format(Inscricao, "00000000-00")
    BuscaContribuinte = Obrig.BuscaSujeitoPassivoObrigacao(Inscricao, Razao, Endereco, Documento)
End Function
