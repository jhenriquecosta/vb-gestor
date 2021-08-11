VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Sub GravaUnidade(IC As String, CodComponente As String, CodGrupo As String, Unidade As String, Valor As Double, SubUnidade As Integer)
    Dim Valores As String
    Dim Campos As String
    On Error Resume Next
    Valores = Bdados.PreparaValor(Bdados.Converte(IC, TCTexto), CodComponente, Bdados.Converte(Valor, TCDuplo), CodGrupo, Unidade, SubUnidade)
    Campos = "tdi_tim_ic,tdi_tco_cod_componente,tdi_valor_item,tdi_tgc_cod_grupo,tdi_tim_ic_unidade,tdi_tim_sub_unidade"
    Call Bdados.InsereDados("Tab_Detalhe_Imovel", Valores, Campos)
End Sub


Function InsereTerritorio(Logradouro As String, InscricaoCadastral As String, Unidade As Integer, InscricaoMunicipal As String, Optional InscCadastralCondominio As String, Optional SubUnidade As Integer = 0) As Boolean
    Dim Campos As String
    Dim Valores As String
    Dim Rs As VSRecordset
    
    InsereTerritorio = True
    'Vou cadastrar o imovel
    Campos = "tim_ic, TIM_UNIDADE, tim_tci_im , tim_tlg_cod_logradouro, " _
        & "tim_numero,  tim_complemento, tim_loteamento, tim_quadra," _
        & "tim_lote, tim_cep,tim_tipo_imovel," _
        & "tim_ocupante,tim_cgc_cpf_ocupante,tim_ic_anterior"
    If Boletim <> tbo_Condominio Then
        Valores = Bdados.PreparaValor(Bdados.Converte(InscricaoCadastral, TCTexto), Unidade, InscricaoMunicipal, _
           Logradouro, Trim(txtNumero), Trim(txtComplemento), _
            Trim(txtLoteamento), Trim(txtQuadra), Trim(txtLote), _
            Trim(Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")), cboTipoImovel.ListIndex + 1, txtOcupante, txtCpfOcupante, Nvl(txtIcAnterior, 0))
    Else
        Valores = Bdados.PreparaValor(Bdados.Converte(InscricaoCadastral, TCTexto), Unidade, InscricaoMunicipal, Logradouro, Trim(txtNumeroBc), Trim(txtComplementoBc), Trim(txtLoteamentoBc), Trim(txtQuadraBc), Trim(txtLoteBc), Trim(txtCepImBc), cboTipoImovelBc.ListIndex + 1, txtOcupanteBc, txtCpfOcupanteBc, Nvl(txtInscAnteriorBC, 0))
    End If
    If Trim(InscCadastralCondominio) <> "" Then
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(InscCadastralCondominio, TCTexto))
        Campos = Campos & ",TIM_IC_CONDOMINIO"
    End If
    Valores = Valores & Bdados.PreparaValor(Nvl(txtCodMens, 0), txtCodBairro, Nvl(txtZona, 1))
    Campos = Campos & ",TIM_COD_MENSAGEM,TIM_TBA_COD_BAIRRO,TIM_ZONA"
    Campos = Campos & ",TIM_AFORAMENTO_NUMERO,TIM_AFORAMENTO_FICHA,TIM_AFORAMENTO_LIVRO,TIM_AFORAMENTO_FOLHA,TIM_AFORAMENTO_REGISTRO,TIM_DATA_REGISTRO,TIM_SUB_UNIDADE,TIM_TUS_COD_USUARIO,TIM_DATA_CADASTRO,TIM_TIPO_BOLETIM"
    Valores = Valores & Bdados.PreparaValor(txtNumAforamento, txtFichaAforamento, txtLivroAforamento, txtFolhaAforamento, txtRegistro, txtDtRegistro, SubUnidade, Aplicacoes.Usuario, Format(Date, "DD/MM/YYYY"), IIf(Boletim <> tbo_Condominio, 0, 1))
    If Trim(txtDataAforamento) <> "" Then
        Campos = Campos & ",TIM_AFORAMENTO_DATA"
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(txtDataAforamento, TCDataHora))
    End If
    Call Bdados.GravaDados("TAB_IMOVEL" _
        , Valores, Campos, "tim_ic='" & InscricaoCadastral & "' AND tim_unidade =" & Unidade & " AND TIM_SUB_UNIDADE = " & SubUnidade)
    Bdados.FechaTabela Rs
End Function

Function InsereContrib(InscricaoMunicipal As String, SitCadastral As String) As Boolean
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim Campos As String
    Dim Valores As String
    InsereContrib = True
    If Not Cadastro.ContribuinteHabilitado(InscricaoMunicipal, SitCadastral) Then
        Call Util.Avisa("O Contribuinte está " & SitCadastral & " e não pode adquirir novos imóveis.")
        Screen.MousePointer = 0
        txtFatorFixo.Tag = ""
        InsereContrib = False
        Exit Function
    End If
    'Verifica se gravou contribuinte
    Sql = "Select tci_nome from tab_contribuinte where tci_nome ='" & Trim(txtNomeContrib) & "'"
    If Not Bdados.AbreTabela(Sql, Rs) Then
        InscricaoMunicipal = Cadastro.GeraInscMunicipal(Right(Date, 1), 11, 1)
    End If
    
    Campos = "tci_im, tci_nome , tci_logradouro, " _
        & "tci_nome_logradouro, tci_numero, tci_complemento," _
        & "tci_bairro, tci_cep , tci_cidade, tci_UF, tci_data_cadastro," _
        & "tci_tsc_cod_sit_cad,tci_tus_cod_usuario,tci_tipo_contribuinte,TCI_CGC_CPF"
        
    
    Valores = Bdados.PreparaValor(InscricaoMunicipal, Trim(txtNomeContrib), _
        txtNomeTipoLogrContrib, Trim(txtNomeLogrContrib), Trim(txtNumeroContrib), _
        Trim(txtCompContrib), Trim(txtBairroContrib), Trim(txtCep), _
        Trim(txtMunic), txtUf, Bdados.Converte(Date, TCDataHora), 1, Aplicacoes.Usuario, 0, txtCpfCgc)
    
    Call Bdados.GravaDados("Tab_Contribuinte" _
     , Valores, Campos, "tci_im='" & InscricaoMunicipal & "'")
    Bdados.FechaTabela Rs
End Function

Public Function LoteCadastrado(Distrito As String, Setor As String, Quadra As String) As Boolean
    Dim Sql As String
    Sql = "SELECT TCL_TUS_COD_USUARIO,TCL_DATA_ABERTURA FROM TAB_CONTROLE_LOTE WHERE " _
            & "TCL_DISTRITO = '" & Distrito _
            & "' AND TCL_SETOR='" & Setor _
            & "' AND TCL_QUADRA = '" & Quadra & "'"
     If Bdados.AbreTabela(Sql) Then LoteCadastrado = True
     
End Function

Public Function LoteFechado(Distrito As String, Setor As String, Quadra As String) As Boolean
Dim Sql As String
    Sql = "SELECT TCL_STATUS FROM TAB_CONTROLE_LOTE WHERE " _
            & "TCL_DISTRITO = '" & Distrito _
            & "' AND TCL_SETOR='" & Setor _
            & "' AND TCL_QUADRA = '" & Quadra & "'"
     If Bdados.AbreTabela(Sql) Then
        LoteFechado = IIf(Rs!TCL_STATUS = 1, False, True)
     Else
        LoteFechado = False
     End If
End Function

Public Function FechaLote(Distrito As String, Setor As String, Quadra As String) As Boolean
    Dim Sql As String
    Sql = "SELECT TCL_STATUS FROM TAB_CONTROLE_LOTE WHERE " _
            & "TCL_DISTRITO = '" & Distrito _
            & "' AND TCL_SETOR='" & Setor _
            & "' AND TCL_QUADRA = '" & Quadra & "'"
     If Bdados.AbreTabela(Sql) Then
        LoteFechado = IIf(Rs!TCL_STATUS = 1, False, True)
     Else
        LoteFechado = False
     End If
End Function




