VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eContribuinte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Im As String
Private m_CgcCpf As String
Private m_Nome As String
Private m_Fantasia As String
Private m_Logradouro As String
Private m_NomeLogradouro As String
Private m_Numero As String
Private m_Complemento As String
Private m_Bairro As String
Private m_Cep As String
Private m_Cidade As String
Private m_Uf As String
Private m_DataCadastro As String
Private m_DataModificacao As String
Private m_CodSitCadastral As Integer
Private m_CodUsuario As String
Private m_InicioAtividade As String
Private m_CodGrupo As Integer
Private m_CodNatureza As Integer
Private m_CodAtividade As Double
Private m_CodAtivPoder As Integer
Private m_Estabelecido As Integer
Private m_GrupoAtividade As Double
Private m_TipoContribuinte As Integer
Private m_TipoRecolhimentoIss As Integer
Private m_Ruc As String
Private m_Conselho As String
Private m_ImovelProprio As Integer
Private m_Registro As String
Private m_NumEmpregado As Integer
Private m_PorteEmpresa As Integer
Private m_CodAtividadeSec As Double
Private m_CodAtividadeTerc As Double
Private m_NivelEscolar As Integer
Private m_Protocolo As String
Private m_Ic As String
Private m_Isento As Integer
Private m_FatorAlvara As Double
Private m_CodRamo As Double
Private m_FoneFax As String
Private m_CNH As String
Private m_Categoria As String
Private m_Autorizacao As Integer
Private m_PontoRecepcao As String
Private m_CodHorario As Integer
Private m_CodLogradouro As String
Private m_CodBairro As String


Public Property Get CodBairro() As String
    CodBairro = m_CodBairro
End Property

Public Property Let CodBairro(ByVal Value As String)
    m_CodBairro = Value
End Property

Public Property Get CodLogradouro() As String
    CodLogradouro = m_CodLogradouro
End Property

Public Property Let CodLogradouro(ByVal Value As String)
    m_CodLogradouro = Value
End Property

Public Property Get CodHorario() As Integer
    CodHorario = m_CodHorario
End Property

Public Property Let CodHorario(ByVal Value As Integer)
    m_CodHorario = Value
End Property

Public Property Get PontoRecepcao() As String
    PontoRecepcao = m_PontoRecepcao
End Property

Public Property Let PontoRecepcao(ByVal Value As String)
    m_PontoRecepcao = Value
End Property

Public Property Get Autorizacao() As Integer
    Autorizacao = m_Autorizacao
End Property

Public Property Let Autorizacao(ByVal Value As Integer)
    m_Autorizacao = Value
End Property

Public Property Get Categoria() As String
    Categoria = m_Categoria
End Property

Public Property Let Categoria(ByVal Value As String)
    m_Categoria = Value
End Property

Public Property Get CNH() As String
    CNH = m_CNH
End Property

Public Property Let CNH(ByVal Value As String)
    m_CNH = Value
End Property

Public Property Get FoneFax() As String
    FoneFax = m_FoneFax
End Property

Public Property Let FoneFax(ByVal Value As String)
    m_FoneFax = Value
End Property

Public Property Get CodRamo() As Double
    CodRamo = m_CodRamo
End Property

Public Property Let CodRamo(ByVal Value As Double)
    m_CodRamo = Value
End Property

Public Property Get FatorAlvara() As Double
    FatorAlvara = m_FatorAlvara
End Property

Public Property Let FatorAlvara(ByVal Value As Double)
    m_FatorAlvara = Value
End Property

Public Property Get Isento() As Integer
    Isento = m_Isento
End Property

Public Property Let Isento(ByVal Value As Integer)
    m_Isento = Value
End Property

Public Property Get Ic() As String
    Ic = m_Ic
End Property

Public Property Let Ic(ByVal Value As String)
    m_Ic = Value
End Property

Public Property Get Protocolo() As String
    Protocolo = m_Protocolo
End Property

Public Property Let Protocolo(ByVal Value As String)
    m_Protocolo = Value
End Property

Public Property Get NivelEscolar() As Integer
    NivelEscolar = m_NivelEscolar
End Property

Public Property Let NivelEscolar(ByVal Value As Integer)
    m_NivelEscolar = Value
End Property

Public Property Get CodAtividadeTerc() As Double
    CodAtividadeTerc = m_CodAtividadeTerc
End Property

Public Property Let CodAtividadeTerc(ByVal Value As Double)
    m_CodAtividadeTerc = Value
End Property

Public Property Get CodAtividadeSec() As Double
    CodAtividadeSec = m_CodAtividadeSec
End Property

Public Property Let CodAtividadeSec(ByVal Value As Double)
    m_CodAtividadeSec = Value
End Property

Public Property Get PorteEmpresa() As Integer
    PorteEmpresa = m_PorteEmpresa
End Property

Public Property Let PorteEmpresa(ByVal Value As Integer)
    m_PorteEmpresa = Value
End Property

Public Property Get NumEmpregado() As Integer
    NumEmpregado = m_NumEmpregado
End Property

Public Property Let NumEmpregado(ByVal Value As Integer)
    m_NumEmpregado = Value
End Property

Public Property Get Registro() As String
    Registro = m_Registro
End Property

Public Property Let Registro(ByVal Value As String)
    m_Registro = Value
End Property

Public Property Get ImovelProprio() As Integer
    ImovelProprio = m_ImovelProprio
End Property

Public Property Let ImovelProprio(ByVal Value As Integer)
    m_ImovelProprio = Value
End Property

Public Property Get Conselho() As String
    Conselho = m_Conselho
End Property

Public Property Let Conselho(ByVal Value As String)
    m_Conselho = Value
End Property

Public Property Get Ruc() As String
    Ruc = m_Ruc
End Property

Public Property Let Ruc(ByVal Value As String)
    m_Ruc = Value
End Property

Public Property Get TipoRecolhimentoIss() As Integer
    TipoRecolhimentoIss = m_TipoRecolhimentoIss
End Property

Public Property Let TipoRecolhimentoIss(ByVal Value As Integer)
    m_TipoRecolhimentoIss = Value
End Property

Public Property Get TipoContribuinte() As Integer
    TipoContribuinte = m_TipoContribuinte
End Property

Public Property Let TipoContribuinte(ByVal Value As Integer)
    m_TipoContribuinte = Value
End Property

Public Property Get GrupoAtividade() As Double
    GrupoAtividade = m_GrupoAtividade
End Property

Public Property Let GrupoAtividade(ByVal Value As Double)
    m_GrupoAtividade = Value
End Property

Public Property Get Estabelecido() As Integer
    Estabelecido = m_Estabelecido
End Property

Public Property Let Estabelecido(ByVal Value As Integer)
    m_Estabelecido = Value
End Property

Public Property Get CodAtivPoder() As Integer
    CodAtivPoder = m_CodAtivPoder
End Property

Public Property Let CodAtivPoder(ByVal Value As Integer)
    m_CodAtivPoder = Value
End Property

Public Property Get CodAtividade() As Double
    CodAtividade = m_CodAtividade
End Property

Public Property Let CodAtividade(ByVal Value As Double)
    m_CodAtividade = Value
End Property

Public Property Get CodNatureza() As Integer
    CodNatureza = m_CodNatureza
End Property

Public Property Let CodNatureza(ByVal Value As Integer)
    m_CodNatureza = Value
End Property

Public Property Get CodGrupo() As Integer
    CodGrupo = m_CodGrupo
End Property

Public Property Let CodGrupo(ByVal Value As Integer)
    m_CodGrupo = Value
End Property

Public Property Get InicioAtividade() As String
    InicioAtividade = m_InicioAtividade
End Property

Public Property Let InicioAtividade(ByVal Value As String)
    m_InicioAtividade = Value
End Property

Public Property Get CodUsuario() As String
    CodUsuario = m_CodUsuario
End Property

Public Property Let CodUsuario(ByVal Value As String)
    m_CodUsuario = Value
End Property

Public Property Get CodSitCadastral() As Integer
    CodSitCadastral = m_CodSitCadastral
End Property

Public Property Let CodSitCadastral(ByVal Value As Integer)
    m_CodSitCadastral = Value
End Property

Public Property Get DataModificacao() As String
    DataModificacao = m_DataModificacao
End Property

Public Property Let DataModificacao(ByVal Value As String)
    m_DataModificacao = Value
End Property

Public Property Get DataCadastro() As String
    DataCadastro = m_DataCadastro
End Property

Public Property Let DataCadastro(ByVal Value As String)
    m_DataCadastro = Value
End Property

Public Property Get Uf() As String
    Uf = m_Uf
End Property

Public Property Let Uf(ByVal Value As String)
    m_Uf = Value
End Property

Public Property Get Cidade() As String
    Cidade = m_Cidade
End Property

Public Property Let Cidade(ByVal Value As String)
    m_Cidade = Value
End Property

Public Property Get Cep() As String
    Cep = m_Cep
End Property

Public Property Let Cep(ByVal Value As String)
    m_Cep = Value
End Property

Public Property Get Bairro() As String
    Bairro = m_Bairro
End Property

Public Property Let Bairro(ByVal Value As String)
    m_Bairro = Value
End Property

Public Property Get Complemento() As String
    Complemento = m_Complemento
End Property

Public Property Let Complemento(ByVal Value As String)
    m_Complemento = Value
End Property

Public Property Get Numero() As String
    Numero = m_Numero
End Property

Public Property Let Numero(ByVal Value As String)
    m_Numero = Value
End Property

Public Property Get NomeLogradouro() As String
    NomeLogradouro = m_NomeLogradouro
End Property

Public Property Let NomeLogradouro(ByVal Value As String)
    m_NomeLogradouro = Value
End Property

Public Property Get Logradouro() As String
    Logradouro = m_Logradouro
End Property

Public Property Let Logradouro(ByVal Value As String)
    m_Logradouro = Value
End Property

Public Property Get Fantasia() As String
    Fantasia = m_Fantasia
End Property

Public Property Let Fantasia(ByVal Value As String)
    m_Fantasia = Value
End Property

Public Property Get Nome() As String
    Nome = m_Nome
End Property

Public Property Let Nome(ByVal Value As String)
    m_Nome = Value
End Property

Public Property Get CgcCpf() As String
    CgcCpf = m_CgcCpf
End Property

Public Property Let CgcCpf(ByVal Value As String)
    m_CgcCpf = Value
End Property

Public Property Get Im() As String
    Im = m_Im
End Property

Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Function Buscar(Optional Im As String, Optional Cpf As String, Optional Detalhes As Boolean = True) As Boolean
    Dim SQL As String, sqlaux As String
    Dim RS As VSRecordset
    
    SQL = "Select * from tab_contribuinte "
    If Trim(Im) <> "" Then
        sqlaux = sqlaux & " and tci_im = '" & Im & "'"
    End If
    If Trim(Cpf) <> "" Then
        sqlaux = sqlaux & " and TCI_CGC_CPF = '" & Cpf & "'"
    End If
    If sqlaux <> "" Then
        SQL = SQL & " where " & Right(sqlaux, Len(sqlaux) - 4)
    Else
        Exit Function
    End If
    If Bdados.AbreTabela(SQL, RS) Then
        Buscar = True
        m_Im = "" & RS!tci_im
        m_CgcCpf = "" & RS!TCI_CGC_CPF
        m_Nome = "" & RS!tci_nome
        m_Fantasia = "" & RS!tci_fantasia
        m_Logradouro = "" & RS!tci_logradouro
        m_NomeLogradouro = "" & RS!tci_nome_logradouro
        m_Numero = "" & RS!tci_numero
        m_Complemento = "" & RS!tci_COMPLEMENTO
        m_Bairro = "" & RS!tci_bairro
        m_Cep = "" & RS!tci_cep
        m_Cidade = "" & RS!tci_cidade
        m_Uf = "" & RS!tci_UF
        m_DataCadastro = "" & RS!tci_data_cadastro
        m_DataModificacao = "" & RS!tci_data_modific
        m_CodSitCadastral = "" & RS!tci_tsc_cod_sit_cad
        m_CodUsuario = "" & RS!tci_tus_cod_usuario
        m_Registro = "" & RS!tci_registro
        m_CodAtividade = "" & RS!tci_tae_cae
        m_InicioAtividade = "" & RS!tci_inicio_atividade
        If Not IsNull(RS!TCI_FONE_FAX) And Trim(RS!TCI_FONE_FAX) <> "" Then m_FoneFax = "" & RS!TCI_FONE_FAX
        If Detalhes = True Then
            
            m_CodGrupo = "" & RS!tci_tga_cod_grupo
            m_CodNatureza = "" & RS!tci_tnj_cod_natureza
            m_CodAtivPoder = "" & RS!tci_tap_cod_ativ_poder
            m_Estabelecido = "" & RS!tci_estab
            m_GrupoAtividade = "" & RS!tci_grupo_cae
            m_TipoContribuinte = "" & RS!tci_tipo_contribuinte
            m_TipoRecolhimentoIss = "" & RS!TCI_TIPO_RECOLHIMENTO_ISS
            m_Ruc = "" & RS!tci_ruc
            m_Conselho = "" & RS!tci_conselho
            m_ImovelProprio = "" & RS!tci_imovel_proprio
            m_NumEmpregado = "" & RS!tci_num_empregado
            m_PorteEmpresa = "" & RS!tci_porte_empresa
            m_CodAtividadeSec = "" & RS!tci_tae_cae_secund
            m_CodAtividadeTerc = "" & RS!tci_tae_cae_terc
            m_NivelEscolar = "" & RS!tci_nivel_escolar
            m_Protocolo = "" & RS!tci_protocolo
            m_Ic = "" & RS!tci_tim_ic
            If Not IsNull(RS!TCI_ISENTO) Then m_Isento = "" & RS!TCI_ISENTO
            If Not IsNull(RS!TCI_FATOR_ALVARA) Then m_FatorAlvara = "" & RS!TCI_FATOR_ALVARA
            If Not IsNull(RS!TCI_TRA_COD_RAMO) Then m_CodRamo = "" & RS!TCI_TRA_COD_RAMO
            If Not IsNull(RS!TCI_CNH) And Trim(RS!TCI_CNH) <> "" Then m_CNH = "" & RS!TCI_CNH
            If Not IsNull(RS!TCI_CATEGORIA) And Trim(RS!TCI_CATEGORIA) <> "" Then m_Categoria = "" & RS!TCI_CATEGORIA
            If Not IsNull(RS!TCI_AUTORIZACAO) And Trim(RS!TCI_AUTORIZACAO) <> "" Then m_Autorizacao = "" & RS!TCI_AUTORIZACAO
            If Not IsNull(RS!TCI_PONTO_RECEPCAO) And Trim(RS!TCI_PONTO_RECEPCAO) <> "" Then m_PontoRecepcao = "" & RS!TCI_PONTO_RECEPCAO
            If Not IsNull(RS!TCI_THF_COD_HORARIO) And Trim(RS!TCI_THF_COD_HORARIO) <> "" Then m_CodHorario = "" & RS!TCI_THF_COD_HORARIO
            If Not IsNull(RS!TCI_COD_LOGRADOURO) And Trim(RS!TCI_COD_LOGRADOURO) <> "" Then m_CodLogradouro = "" & RS!TCI_COD_LOGRADOURO
            If Not IsNull(RS!TCI_COD_BAIRRO) And Trim(RS!TCI_COD_BAIRRO) <> "" Then m_CodBairro = "" & RS!TCI_COD_BAIRRO
        End If
    End If
End Function

Public Function Salvar() As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Campos = "tci_im,tci_cgc_cpf, tci_nome, tci_fantasia, tci_logradouro, tci_nome_logradouro, tci_numero, tci_complemento," & _
                " tci_bairro, tci_cep, tci_cidade, tci_UF, tci_data_cadastro, tci_data_modific, tci_tsc_cod_sit_cad, tci_tus_cod_usuario," & _
                " tci_inicio_atividade, tci_tga_cod_grupo, tci_tnj_cod_natureza, tci_tae_cae, tci_tap_cod_ativ_poder, tci_estab, " & _
                " tci_grupo_cae, tci_tipo_contribuinte, tci_tipo_recolhimento_iss, tci_ruc, tci_conselho, " & _
                " tci_imovel_proprio, tci_registro, tci_num_empregado, tci_porte_empresa, tci_tae_cae_secund, tci_tae_cae_terc," & _
                " tci_nivel_escolar, tci_protocolo, tci_ISENTO, tci_FATOR_ALVARA, tci_TRA_COD_RAMO, tci_FONE_FAX, " & _
                " tci_CNH, tci_CATEGORIA, tci_AUTORIZACAO, tci_PONTO_RECEPCAO, tci_THF_COD_HORARIO, tci_COD_LOGRADOURO, tci_COD_BAIRRO"
    If Trim(m_Ic) <> "" Then Campos = Campos & ",tci_tim_ic"
    If Trim(m_Ic) <> "" Then Valores = Valores & Bdados.PreparaValor(Bdados.Converte(m_Ic, tctexto))
    If Trim(m_DataCadastro) = "" Then m_DataCadastro = Format(Date, "dd/mm/yyyy")
    If Trim(m_DataModificacao) = "" Then m_DataModificacao = Format(Date, "dd/mm/yyyy")
    If Trim(m_InicioAtividade) = "" Then m_InicioAtividade = Format(Date, "dd/mm/yyyy")
    Valores = Bdados.PreparaValor(m_Im, m_CgcCpf, m_Nome, m_Fantasia, m_Logradouro, m_NomeLogradouro, m_Numero, m_Complemento, _
                  m_Bairro, m_Cep, m_Cidade, m_Uf, m_DataCadastro, m_DataModificacao, m_CodSitCadastral, m_CodUsuario, _
                  m_InicioAtividade, m_CodGrupo, m_CodNatureza, m_CodAtividade, m_CodAtivPoder, m_Estabelecido, _
                  m_GrupoAtividade, m_TipoContribuinte, m_TipoRecolhimentoIss, m_Ruc, m_Conselho, m_ImovelProprio, _
                  m_Registro, m_NumEmpregado, m_PorteEmpresa, m_CodAtividadeSec, m_CodAtividadeTerc, _
                  m_NivelEscolar, m_Protocolo, m_Isento, m_FatorAlvara, m_CodRamo, m_FoneFax, _
                  m_CNH, m_Categoria, m_Autorizacao, m_PontoRecepcao, m_CodHorario, m_CodLogradouro, m_CodBairro)
    If Trim(m_Im) <> "" Then
        Condicao = "tci_im = '" & m_Im & "'"
    ElseIf Trim(m_CgcCpf) <> "" Then
        Condicao = "tci_cgc_cpf = '" & m_CgcCpf & "'"
    End If
    If Bdados.GravaDados("Tab_contribuinte", Valores, Campos, Condicao) Then
        Salvar = True
    End If
End Function

Public Function Excluir(Im As String) As Boolean
    If Bdados.DeletaDados("tab_contribuinte", "tci_im = '" & Im & "' or tci_cgc_cpf = '" & Im & "'") Then Excluir = True
End Function

Public Function PreencherGrd(Grid As Object, Optional Im As String, Optional Nome As String, Optional TipoContrib As Byte = 0) As Boolean
    Dim RS As VSRecordset
    Dim SQL As String
    SQL = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf as CPF_CGC from Tab_Contribuinte where tci_tsc_cod_sit_cad =1"
    If Trim(Nome) <> "" Then
        SQL = SQL & " AND tci_nome like '%" & Nome & "%'"
    End If
    If Trim(Im) <> "" Then
        SQL = SQL & " and tci_im = '" & Im & "' or tci_cgc_cpf = '" & Im & "'"
    End If
    If TipoContrib <> 0 Then
        SQL = SQL & " and tci_tipo_contribuinte = " & TipoContrib
    End If
    If Grid.Preencher(Bdados, SQL, 1400, 4000, 2000) Then PreencherGrd = True
End Function

Public Function PreencherComboTipoLogr(Combo As Object) As Boolean
    Dim SQL As String
    SQL = "Select DISTINCT(ttl_nome) From Tab_Tipo_Logr"
    If Combo.Preencher(Bdados, SQL) Then PreencherComboTipoLogr = True
End Function

Public Function VerificaTEMImovel(InsMunicipal As String) As Boolean
    Dim SQL As String                   'VERIFICA SE TEM IMOVEL CADASTRADO PRA INXCRIÇÃO MUNICIPAL
    Dim RS As VSRecordset
    SQL = "Select tim_ic from Tab_Imovel where tim_tci_im = '" & InsMunicipal & "'"
    If Bdados.AbreTabela(SQL, RS) Then VerificaTEMImovel = True
End Function

Public Function VerificaTEMDebito(InsMunicipal As String) As Boolean
    Dim SQL As String                   'VERIFICA SE TEM debito PRA INXCRIÇÃO MUNICIPAL
    Dim RS As VSRecordset
    SQL = "Select * FROM tab_geracao_tributo WHERE tgt_im = '" & InsMunicipal & "' and tgt_data_pagamento is null"
    If Bdados.AbreTabela(SQL, RS) Then VerificaTEMDebito = True
End Function

Public Sub PreencherCboNaturezaJur(Combo As Object)
    Dim SQL As String
    SQL = "Select tnj_natureza, TNJ_COD_NATUREZA From Tab_Natureza_Juridica"
    Combo.Preencher Bdados, SQL
End Sub
    
Public Sub PreencherCboClasseAtividade(Combo As Object)
    Dim SQL As String
    SQL = "Select tga_nome, tga_cod_grupo From Tab_Grupo_Atividade"
    Combo.Preencher Bdados, SQL
End Sub

Public Sub PreencherCboAtividadePoder(Combo As Object)
    Dim SQL As String
    SQL = "Select tap_poder, tap_cod_ativ_poder From Tab_atividade_Poder"
    Combo.Preencher Bdados, SQL
End Sub

Public Sub PreencherCboAtividadeEcon(Combo As Object)
    Dim SQL As String
    SQL = "Select tae_nome, TAE_CAE From Tab_atividade_Economica"
    Combo.Preencher Bdados, SQL
End Sub

Public Sub PreencherCboSitCad(Combo As Object)
    Dim SQL As String
    SQL = "Select tsc_nome, tsc_cod_sit_cad  From Tab_sit_cadastral order by tsc_cod_sit_cad"
    Combo.Preencher Bdados, SQL
End Sub


Public Function ExibirFicha(grd As Object, Im As String, Contribuinte As String, Optional PeriodoInicial As String, Optional PeriodoFinal As String, Optional Situacao As String, Optional NFInicio As String, Optional NFFim As String) As Boolean
    On Error GoTo trata
    Dim SQL As String, pos As Integer
    
    Select Case Situacao
        Case "NÃO PAGO"
            SQL = "SELECT "
            SQL = SQL & Bdados.ParteTexto("tgt_periodo", MidVs, 5, 2, True) & Bdados.Concatena & "'/'" & Bdados.Concatena & Bdados.ParteTexto("tgt_periodo", MidVs, 1, 4, True) & " as Periodo,"
            SQL = SQL & " tgt_cod_pagamento as DAM,"
            SQL = SQL & " tip_sigla_imposto AS Tributo,"
            SQL = SQL & " tdd_num_nota_inicial AS NFInicial,"
            SQL = SQL & " tdd_num_nota_final as NFFinal,"
            SQL = SQL & FuncaoReal("tdd_total_nota") & " as TotalNF,"
            SQL = SQL & FuncaoReal("tdd_total_material_reducao") & " as Deduções,"
            SQL = SQL & FuncaoReal("tgt_valor_tributo") & " as Original,"
            SQL = SQL & "'' as Juros,"
            SQL = SQL & "'' as Multa,"
            SQL = SQL & "'' as DataPagto,"
            SQL = SQL & "'' as ValorPago,"
            SQL = SQL & "tdd_obs As Obs"
            SQL = SQL & " From TAB_GERACAO_TRIBUTO, Tab_Detalhe_Dam, TAB_IMPOSTO"
            SQL = SQL & " Where tgt_cod_pagamento *= tdd_tgt_cod_pagamento"
            SQL = SQL & " AND tgt_tip_cod_imposto=tip_cod_imposto"
            SQL = SQL & " AND tgt_cod_pagamento NOT IN (SELECT tdr_tgt_cod_pagamento FROM TAB_DARM_RECEBIDO WHERE tdr_im=tgt_im AND tdr_tip_cod_imposto=tgt_tip_cod_imposto AND tdr_periodo=tgt_periodo)"
            SQL = SQL & " AND tgt_tip_cod_imposto LIKE '111305%'"
            SQL = SQL & " AND tgt_im='" & Im & "'"
        Case "PAGO"
            SQL = "SELECT "
            SQL = SQL & Bdados.ParteTexto("tgt_periodo", MidVs, 5, 2, True) & Bdados.Concatena & "'/'" & Bdados.Concatena & Bdados.ParteTexto("tgt_periodo", MidVs, 1, 4, True) & " as Periodo,"
            SQL = SQL & " tgt_cod_pagamento as DAM,"
            SQL = SQL & " tip_sigla_imposto AS Tributo,"
            SQL = SQL & " tdd_num_nota_inicial AS NFInicial,"
            SQL = SQL & " tdd_num_nota_final as NFFinal,"
            SQL = SQL & FuncaoReal("tdd_total_nota") & " as TotalNF,"
            SQL = SQL & FuncaoReal("tdd_total_material_reducao") & " as Deduções,"
            SQL = SQL & FuncaoReal("tgt_valor_tributo") & " as Original,"
            SQL = SQL & FuncaoReal("tdr_valor_real_juros") & " as Juros,"
            SQL = SQL & FuncaoReal("tdr_valor_real_multa") & " as Multa,"
            SQL = SQL & "tdr_data_pagamento as DataPagto,"
            SQL = SQL & FuncaoReal("tdr_valor_real_pago") & " as ValorPago,"
            SQL = SQL & "tdd_obs As Obs"
            SQL = SQL & " From TAB_GERACAO_TRIBUTO, Tab_Detalhe_Dam, TAB_DARM_RECEBIDO, TAB_IMPOSTO"
            SQL = SQL & " Where tgt_cod_pagamento *= tdd_tgt_cod_pagamento"
            SQL = SQL & " AND tgt_tip_cod_imposto=tip_cod_imposto"
            SQL = SQL & " AND tgt_cod_pagamento = tdr_tgt_cod_pagamento"
            SQL = SQL & " AND tgt_tip_cod_imposto LIKE '111305%"
            SQL = SQL & " AND tgt_im='" & Im & "'"
        Case Else
            SQL = "SELECT "
            SQL = SQL & Bdados.ParteTexto("tgt_periodo", MidVs, 5, 2, True) & Bdados.Concatena & "'/'" & Bdados.Concatena & Bdados.ParteTexto("tgt_periodo", MidVs, 1, 4, True) & " as Periodo,"
            SQL = SQL & " tgt_cod_pagamento as DAM,"
            SQL = SQL & " tip_sigla_imposto AS Tributo,"
            SQL = SQL & " tdd_num_nota_inicial AS NFInicial,"
            SQL = SQL & " tdd_num_nota_final as NFFinal,"
            SQL = SQL & FuncaoReal("tdd_total_nota") & " as TotalNF,"
            SQL = SQL & FuncaoReal("tdd_total_material_reducao") & " as Deduções,"
            SQL = SQL & FuncaoReal("tgt_valor_tributo") & " as Original,"
            SQL = SQL & FuncaoReal("tdr_valor_real_juros") & " as Juros,"
            SQL = SQL & FuncaoReal("tdr_valor_real_multa") & " as Multa,"
            SQL = SQL & "tdr_data_pagamento as DataPagto,"
            SQL = SQL & FuncaoReal("tdr_valor_real_pago") & " as ValorPago,"
            SQL = SQL & "tdd_obs As Obs"
            SQL = SQL & " From TAB_GERACAO_TRIBUTO, Tab_Detalhe_Dam, TAB_DARM_RECEBIDO, TAB_IMPOSTO"
            SQL = SQL & " Where tgt_cod_pagamento *= tdd_tgt_cod_pagamento"
            SQL = SQL & " AND tgt_tip_cod_imposto=tip_cod_imposto"
            SQL = SQL & " AND tgt_cod_pagamento *= tdr_tgt_cod_pagamento"
            SQL = SQL & " AND tgt_tip_cod_imposto LIKE '111305%'"
            SQL = SQL & " AND tgt_im='" & Im & "'"
    End Select
    
    If PeriodoInicial <> "" Then
        SQL = SQL & " AND tgt_periodo>=" & Util.ParseString(PeriodoInicial, "/", 2) & Util.ParseString(PeriodoInicial, "/", 1)
    End If
    If PeriodoFinal <> "" Then
        SQL = SQL & " AND tgt_periodo<=" & Util.ParseString(PeriodoFinal, "/", 2) & Util.ParseString(PeriodoFinal, "/", 1)
    End If
    If NFInicio <> "" Then
        SQL = SQL & " AND tdd_num_nota_inicial>=" & NFInicio
    End If
    If NFFim <> "" Then
        SQL = SQL & " AND tdd_num_nota_final<=" & NFFim
    End If
    SQL = SQL & " ORDER BY tgt_periodo, tgt_tip_cod_imposto"
    If grd.Preencher(Bdados, SQL) Then
        ExibirFicha = True
        grd.Mensagem = "Total lançado : " & Format(grd.Colunas(8).Soma, "currency") & " x Total arrecadado : " & Format(grd.Colunas(9).Soma + grd.Colunas(10).Soma + grd.Colunas(12).Soma, "currency") & " (" & Format((grd.Colunas(9).Soma + grd.Colunas(10).Soma + grd.Colunas(12).Soma) / grd.Colunas(8).Soma, "percent") & ")"
    End If
    Exit Function
trata:
    Erro Err.Description
End Function

Public Function GravarHistorico(CodMudanca As String, Im As String) As Boolean
    Dim SQL As String
    SQL = "Select tab_contribuinte.* , " & Bdados.Converte(Date, TCDataHora) & ",'" & _
            Format(Time, "HH:MM:SS") & "'," & CodMudanca & _
            " from tab_contribuinte where tci_im = '" & Im & "'"
    If Bdados.Executa("Insert into tab_contribuinte_historico " & SQL) Then GravarHistorico = True
End Function

Public Function BuscarContribuintes(grd As Object, Optional Im As String, Optional Cgc As String, Optional Razao As String, Optional Fantasia As String, _
                        Optional TipoLogr As String, Optional Logr As String, Optional Bairro As String, Optional Numero As String, _
                        Optional Compl As String, Optional Cidade As String, Optional AtivEconomica As String, Optional ObrigIss As String, _
                        Optional Periodo1 As String, Optional Periodo2 As String, Optional Ponto As String, Optional ByRef RetFiltroRpt As String) As Boolean
    Dim SQL As String
    Dim Condicao As String
    Dim FiltroRpt As String
    SQL = "SELECT tci_im as IM, tci_cgc_cpf as CNPJ, tci_nome as Razão, TCI_FANTASIA AS Fantasia," & _
            " tci_logradouro as Logr, tci_nome_logradouro as Nome, tci_numero as [Nº], tci_complemento as Complemento," & _
            " tci_bairro as Bairro,tci_cidade as Cidade,tci_uf as UF, tae_nome as [Atividade Economica]," & _
            " tci_autorizacao as Autorizacao, tci_fator_alvara as [Fator], tci_INICIO_ATIVIDADE as [Abertura]," & _
            " tci_tus_cod_usuario as Usuario " & _
            " FROM Tab_Contribuinte,Tab_Atividade_Economica" & _
            " where tci_tae_cae=tae_cae"
            
            FiltroRpt = "{TAB_CONTRIBUINTE.tci_tsc_cod_sit_cad} = 1 and {TAB_CONTRIBUINTE.tci_tipo_contribuinte} >0 AND "
            If Trim(Cgc) <> "" Then
                Condicao = " and tci_cgc_cpf ='" & Cgc & "'"
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.tci_cgc_cpf} ='" & Cgc & "' AND "
            End If
            If Trim(Im) <> "" Then
                Condicao = " and tci_im ='" & Im & "'"
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.tci_im} ='" & Im & "' AND "
            End If
            If Trim(Razao) <> "" Then
                Condicao = Condicao & " and (tci_nome like '%" & UCase(Razao) & "%')"
                FiltroRpt = FiltroRpt & " ({TAB_CONTRIBUINTE.tci_nome} like '*" & UCase(Razao) & "*') and "
            End If
            If Trim(Fantasia) <> "" Then
                Condicao = Condicao & " and tci_fantasia like '%" & UCase(Fantasia) & "%'"
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.tci_fantasia} like '*" & UCase(Fantasia) & "*' and "
            End If
            If Trim(TipoLogr) <> "" Then
                Condicao = Condicao & " and tci_logradouro = '" & TipoLogr & "'"
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.tci_logradouro} '" & TipoLogr & "' AND "
            End If
            If Trim(Logr) <> "" Then
                Condicao = Condicao & " and tci_nome_logradouro like '%" & UCase(Logr) & "%'"
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.tci_nome_logradouro} like '*" & UCase(Logr) & "*' AND"
            End If
            If Trim(Bairro) <> "" Then
                Condicao = Condicao & " and tci_bairro like '%" & UCase(Bairro) & "%'"
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.tci_bairro} like '*" & UCase(Bairro) & "*' AND "
            End If
            If Trim(Numero) <> "" Then
                Condicao = Condicao & " and tci_numero = '" & Numero & "'"
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.tci_numero} = '" & Numero & "' AND "
            End If
            If Trim(Compl) <> "" Then
                Condicao = Condicao & " and tci_complemento like '%" & UCase(Compl) & "%'"
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.tci_complemento} like '*" & UCase(Compl) & "*' AND "
            End If
            If Trim(Cidade) <> "" Then
                Condicao = Condicao & " and TCi_CIDADE ='" & UCase(Cidade) & "'"
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.TCI_CIDADE} = '" & UCase(Cidade) & "' AND "
            End If
            If Trim(AtivEconomica) <> "" Then
                Condicao = Condicao & " and tci_tae_cae =" & AtivEconomica
                FiltroRpt = " {TAB_CONTRIBUINTE.tci_tae_cae} =" & AtivEconomica & " AND "
            End If
            If Trim(ObrigIss) <> "" Then
                Condicao = Condicao & " and tci_tipo_recolhimento_iss =" & ObrigIss
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.tci_tae_cae} =" & ObrigIss
            End If
            If Trim(Periodo1) <> "" And Trim(Periodo2) <> "" Then
                Condicao = Condicao & " and tci_data_cadastro >= " & Bdados.Converte(Periodo1, TCDataHora) & " And tci_data_cadastro <= " & Bdados.Converte(Periodo2, TCDataHora)
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.tci_data_cadastro} >=" & Bdados.Converte(Periodo1, TCDataHora) & " AND {TAB_CONTRIBUINTE.tci_data_cadastro} <= " & Bdados.Converte(Periodo2, TCDataHora)
            End If
            If Trim$(Ponto) <> "" Then
                Condicao = Condicao & " and TCI_PONTO_RECEPCAO='" & Ponto & "'"
                FiltroRpt = FiltroRpt & " {TAB_CONTRIBUINTE.TCI_PONTO_RECEPCAO} ='" & Ponto & "'"
            End If
            FiltroRpt = Mid(FiltroRpt, 1, Len(FiltroRpt) - 4)
            SQL = SQL & Condicao
            RetFiltroRpt = FiltroRpt
            If grd.Preencher(Bdados, SQL) Then BuscarContribuintes = True
End Function

Public Function BuscarAcessoEmtDam() As Boolean
    Dim SQL As String
    Dim RS As VSRecordset
    SQL = "SELECT * FROM TAB_ACESSO_USUARIO WHERE TAU_TMO_COD_MODULO ='TCOB' and TAU_TFO_COD_FORMULARIO =102 AND TAU_TUS_COD_USUARIO='" & Aplicacoes.Usuario & "'"
    If Bdados.AbreTabela(SQL, RS) Then BuscarAcessoEmtDam = True
End Function
