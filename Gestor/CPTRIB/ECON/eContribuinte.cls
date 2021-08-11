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
Private m_Obs As String
Private m_Im As String
Private m_ImAnterior As String
Private m_ImAuxiliar As String
Private m_CgcCpf As String
Private m_Nome As String
Private m_Fantasia As String
Private m_Logradouro As String
Private m_NomeLogradouro As String
Private m_Numero As String
Private m_Complemento As String
Private m_Bairro As String
Private M_Data As String
Private m_Motivo As String
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

Private m_AreaEstabelecimento As Double

Private m_CodAtivPoder As Integer
Private m_Estabelecido As Integer
Private m_GrupoAtividade As Double
Private m_TipoContribuinte As Integer
Private m_TipoRecolhimentoIss As Integer
Private m_Ruc As String
Private m_Rg As String
Private m_Conselho As String
Private m_ImovelProprio As Integer
Private m_Registro As String
Public AtividadeTransporte As Double
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
Private m_TipoCadastro As String
Private m_Matriz_Filial As String
Private m_Data_Encerramento As String
Private m_Data_Reabertura As String
Public Telefone           As String
Public m_Tela             As String
Public Email As String
Private m_VariavelAnuncio As String
Private m_VariavelQuantidadeItem As String
Private m_VariavelDataFormatura As String
Private m_VariavelFuncionarioSUS As String
Private m_Sit_Alvara As String
Private m_InicioPrestacaoServico As String

'Private m_crc As String

Public Property Get AreaEstabelecimento() As Double
    AreaEstabelecimento = m_AreaEstabelecimento
End Property

Public Property Let AreaEstabelecimento(ByVal Value As Double)
    m_AreaEstabelecimento = Value
End Property

Public Property Get Motivo() As String
    Motivo = m_Motivo
End Property

Public Property Let Motivo(ByVal Value As String)
    m_Motivo = Value
End Property
Public Property Get Obs() As String
    Obs = m_Obs
End Property

Public Property Let Obs(ByVal Value As String)
    m_Obs = Value
End Property

Public Property Get data() As String
    data = M_Data
End Property

Public Property Let data(ByVal Value As String)
    M_Data = Value
End Property



Public Property Get SituacaoAlvara() As String
    SituacaoAlvara = m_Sit_Alvara
End Property

Public Property Let SituacaoAlvara(ByVal Value As String)
    m_Sit_Alvara = Value
End Property

Public Property Get InicioPrestacaoServico() As String
    InicioPrestacaoServico = m_InicioPrestacaoServico
End Property

Public Property Let InicioPrestacaoServico(ByVal Value As String)
    m_InicioPrestacaoServico = Value
End Property



Public Property Get Nome_Tela() As String
    Nome_Tela = m_Tela
End Property

Public Property Let Nome_Tela(ByVal Value As String)
    m_Tela = Value
End Property


Public Property Get VariavelFuncionarioSUS() As String
    VariavelFuncionarioSUS = m_VariavelFuncionarioSUS
End Property

Public Property Let VariavelFuncionarioSUS(ByVal Value As String)
    m_VariavelFuncionarioSUS = Value
End Property


Public Property Get VariavelDataFormatura() As String
    VariavelDataFormatura = m_VariavelDataFormatura
End Property

Public Property Let VariavelDataFormatura(ByVal Value As String)
    m_VariavelDataFormatura = Value
End Property


Public Property Get VariavelQuantidadeItem() As String
    VariavelQuantidadeItem = m_VariavelQuantidadeItem
End Property

Public Property Let VariavelQuantidadeItem(ByVal Value As String)
    m_VariavelQuantidadeItem = Value
End Property


Public Property Get VariavelAnuncio() As String
    VariavelAnuncio = m_VariavelAnuncio
End Property

Public Property Let VariavelAnuncio(ByVal Value As String)
    m_VariavelAnuncio = Value
End Property


Public Property Get Data_Reabertura() As String
    Data_Reabertura = m_Data_Reabertura
End Property

Public Property Let Data_Reabertura(ByVal Value As String)
    m_Data_Reabertura = Value
End Property


Public Property Get Data_Encerramento() As String
    Data_Encerramento = m_Data_Encerramento
End Property

Public Property Let Data_Encerramento(ByVal Value As String)
    m_Data_Encerramento = Value
End Property


Public Property Get Matriz_Filial() As String
    Matriz_Filial = m_Matriz_Filial
End Property

Public Property Let Matriz_Filial(ByVal Value As String)
    m_Matriz_Filial = Value
End Property


Public Property Get TipoCadastro() As String
    TipoCadastro = m_TipoCadastro
End Property

Public Property Let TipoCadastro(ByVal Value As String)
    m_TipoCadastro = Value
End Property


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

Public Property Let Rg(ByVal Value As String)
    m_Rg = Value
End Property

Public Property Get Rg() As String
    Rg = m_Rg
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

Public Property Get ImAuxiliar() As String
    ImAuxiliar = m_ImAuxiliar
End Property

Public Property Let ImAuxiliar(ByVal Value As String)
    m_ImAuxiliar = Value
End Property

Public Property Get ImAnterior() As String
    ImAnterior = m_ImAnterior
End Property

Public Property Let ImAnterior(ByVal Value As String)
    m_ImAnterior = Value
End Property

Public Function Buscar(Optional Im As String, Optional Cpf As String, Optional Detalhes As Boolean = True) As Boolean
    On Error GoTo TrataErro
    Dim Sql As String, sqlaux As String
    Dim Rs As VSRecordset
    Sql = "Select * from tab_contribuinte "
    If Trim(Im) <> "" Then
        sqlaux = sqlaux & " and tci_im = '" & Im & "'"
    End If
    If Trim(Cpf) <> "" Then
        sqlaux = sqlaux & " and TCI_CGC_CPF = '" & Cpf & "'"
    End If
    If sqlaux <> "" Then
        Sql = Sql & " where " & Right(sqlaux, Len(sqlaux) - 4)
    Else
        Exit Function
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        Buscar = True
        If IsNull(Rs!TCI_AREA_ESTABELECIMENTO) Then
            m_AreaEstabelecimento = 0
        Else
            m_AreaEstabelecimento = CDbl(Rs!TCI_AREA_ESTABELECIMENTO)
        End If
        'm_crc = "" & rs!TCI_CRC
        
        
        m_Rg = "" & Rs!TCI_RG
        m_Im = "" & Rs!TCI_IM
        m_ImAnterior = "" & Rs!TCI_INSCRICAO_ANTERIOR
        m_ImAuxiliar = "" & Rs!TCI_IM_AUXILIAR
        m_CgcCpf = "" & Rs!tci_cgc_cpf
        m_InicioPrestacaoServico = "" & Rs.Fields("TCI_INICIO_PREST_SERV")
        m_Sit_Alvara = "" & Rs.Fields("TCI_SIT_ALVARA")
        m_Nome = "" & Rs!tci_nome
        m_Fantasia = "" & Rs!tci_fantasia
        m_Logradouro = "" & Rs!tci_logradouro
        m_Matriz_Filial = "" & Rs.Fields("tci_matriz_filial")
        m_Data_Encerramento = "" & Rs.Fields("tci_data_encerramento")
        m_Data_Reabertura = "" & Rs.Fields("tci_data_reabertura")
        m_TipoCadastro = "" & Rs.Fields("TCI_TIPO_CADASTRO")
        m_NomeLogradouro = "" & Rs!tci_nome_logradouro
        m_Numero = "" & Rs!tci_numero
        m_Complemento = "" & Rs!tci_COMPLEMENTO
        Telefone = "" & Rs!tci_telefone
        Email = "" & Rs!tci_email
        m_Bairro = "" & Rs!tci_bairro
        m_Cep = "" & Rs!tci_cep
        m_Cidade = "" & Rs!tci_cidade
        m_Uf = "" & Rs!tci_UF
        m_DataCadastro = "" & Rs!tci_data_cadastro
        m_DataModificacao = "" & Rs!tci_data_modific
        m_CodSitCadastral = "" & Rs!tci_tsc_cod_sit_cad
        m_CodUsuario = "" & Rs!tci_tus_cod_usuario
        AtividadeTransporte = Nvl("" & Rs!TCI_TAE_CAE_TRANSPORTE, 0)
        m_Registro = "" & Rs!tci_registro
        If Not IsNull(Rs!TCI_FONE_FAX) And Trim(Rs!TCI_FONE_FAX) <> "" Then m_FoneFax = "" & Rs!TCI_FONE_FAX
        If Detalhes = True Then
            m_InicioAtividade = "" & Rs!tci_inicio_atividade
            m_CodGrupo = "" & Rs!tci_tga_cod_grupo
            m_CodNatureza = "" & Rs!tci_tnj_cod_natureza
            m_CodAtividade = "" & Rs!tci_tae_cae
            m_CodAtivPoder = "" & Rs!tci_tap_cod_ativ_poder
            m_Estabelecido = "" & Nvl("" & Rs!tci_estab, 0)
            If Not IsNull(Rs!tci_grupo_cae) Then
                m_GrupoAtividade = "" & Val(Rs!tci_grupo_cae)
            End If
            m_TipoContribuinte = "" & Rs!tci_tipo_contribuinte
            m_TipoRecolhimentoIss = "" & Rs!TCI_TIPO_RECOLHIMENTO_ISS
            m_Ruc = "" & Rs!tci_ruc
            m_Conselho = "" & Rs!tci_conselho
            m_ImovelProprio = "" & Nvl("" & Rs!tci_imovel_proprio, 0)
            m_NumEmpregado = "" & Nvl("" & Rs!tci_num_empregado, 0)
            m_PorteEmpresa = "" & Rs!tci_porte_empresa
            m_CodAtividadeSec = "" & Rs!tci_tae_cae_secund
            m_CodAtividadeTerc = Nvl("" & Rs!tci_tae_cae_terc, 0)
            m_NivelEscolar = Nvl("" & Rs!tci_nivel_escolar, 0)
            m_Protocolo = "" & Rs!tci_protocolo
            m_Ic = "" & Rs!tci_tim_ic
            m_Obs = "" & Rs!tci_observacao
            If Not IsNull(Rs!tci_ISENTO) Then m_Isento = "" & Rs!tci_ISENTO
            If Not IsNull(Rs!TCI_FATOR_ALVARA) Then m_FatorAlvara = "" & Rs!TCI_FATOR_ALVARA
            If Not IsNull(Rs!TCI_TRA_COD_RAMO) Then m_CodRamo = "" & Rs!TCI_TRA_COD_RAMO
            If Not IsNull(Rs!TCI_CNH) And Trim(Rs!TCI_CNH) <> "" Then m_CNH = "" & Rs!TCI_CNH
            If Not IsNull(Rs!TCI_CATEGORIA) And Trim(Rs!TCI_CATEGORIA) <> "" Then m_Categoria = "" & Rs!TCI_CATEGORIA
            If Not IsNull(Rs!TCI_AUTORIZACAO) And Trim(Rs!TCI_AUTORIZACAO) <> "" Then m_Autorizacao = "" & Rs!TCI_AUTORIZACAO
            If Not IsNull(Rs!TCI_PONTO_RECEPCAO) And Trim(Rs!TCI_PONTO_RECEPCAO) <> "" Then m_PontoRecepcao = "" & Rs!TCI_PONTO_RECEPCAO
            If Not IsNull(Rs!TCI_THF_COD_HORARIO) And Trim(Rs!TCI_THF_COD_HORARIO) <> "" Then m_CodHorario = "" & Rs!TCI_THF_COD_HORARIO
            If Not IsNull(Rs!TCI_COD_LOGRADOURO) And Trim(Rs!TCI_COD_LOGRADOURO) <> "" Then m_CodLogradouro = "" & Rs!TCI_COD_LOGRADOURO
            If Not IsNull(Rs!TCI_COD_BAIRRO) And Trim(Rs!TCI_COD_BAIRRO) <> "" Then m_CodBairro = "" & Rs!TCI_COD_BAIRRO
            If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                Dim RsVariaveis As VSRecordset
                'PEGO OS DADOS DAS VARIAVEIS...
                Sql = "SELECT * FROM TAB_VARIAVEL_ECONOMICO WHERE TVE_TCI_IM      = " & Bdados.Converte(m_Im, tctexto)
                If Bdados.AbreTabela(Sql, RsVariaveis) Then
                    m_VariavelAnuncio = "" & RsVariaveis.Fields("TVE_ANUNCIOS")
                    m_VariavelDataFormatura = "" & RsVariaveis.Fields("TVE_DATA_FORMATURA")
                    m_VariavelFuncionarioSUS = "" & RsVariaveis.Fields("TVE_FUNCIONARIO_SUS")
                    m_VariavelQuantidadeItem = "" & RsVariaveis.Fields("TVE_QUANTIDADE_ITEM")
                End If
            End If
        End If
    End If
    
    Exit Function
TrataErro:
    If Err.Number = 13 Then 'VALOR NULO
        Resume Next
    Else
        Util.Erro Err.Description
        Exit Function
        Resume
    End If
End Function


Public Function BuscarHistorico(Optional Im As String, Optional Cpf As String, Optional Detalhes As Boolean = True, Optional data As String, Optional Motivo As String) As Boolean
    On Error GoTo TrataErro
    Dim Sql As String, sqlaux As String
    Dim Rs As VSRecordset
    Sql = "Select * from tab_contribuinte_Historico "
    If Trim(Im) <> "" Then
        sqlaux = sqlaux & " and tci_im = '" & Im & "'"
    End If
    If data <> "" Then
        sqlaux = sqlaux & " and tc_data = " & Bdados.Converte(data, TCDataHora)
    End If
    If Trim(Cpf) <> "" Then
        sqlaux = sqlaux & " and TCI_CGC_CPF = '" & Cpf & "'"
    End If
    If Motivo <> "" Then
        sqlaux = sqlaux & " and tci_motivo = '" & Motivo & "'"
    End If
    If sqlaux <> "" Then
        Sql = Sql & " where " & Right(sqlaux, Len(sqlaux) - 4)
    Else
        Exit Function
    End If
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscarHistorico = True
        m_AreaEstabelecimento = CDbl(Rs!TCI_AREA_ESTABELECIMENTO)
        'm_CRC = "" & rs!TCI_CRC
        
        m_Im = "" & Rs!TCI_IM
        m_ImAuxiliar = "" & Rs!TCI_IM_AUXILIAR
        m_ImAnterior = "" & Rs!TCI_INSCRICAO_ANTERIOR
        m_CgcCpf = "" & Rs!tci_cgc_cpf
        m_InicioPrestacaoServico = "" & Rs.Fields("TCI_INICIO_PREST_SERV")
        m_Sit_Alvara = "" & Rs.Fields("TCI_SIT_ALVARA")
        m_Nome = "" & Rs!tci_nome
        m_Fantasia = "" & Rs!tci_fantasia
        m_Logradouro = "" & Rs!tci_logradouro
        m_Matriz_Filial = "" & Rs.Fields("tci_matriz_filial")
        m_Data_Encerramento = "" & Rs.Fields("tci_data_encerramento")
        m_Data_Reabertura = "" & Rs.Fields("tci_data_reabertura")
        m_TipoCadastro = "" & Rs.Fields("TCI_TIPO_CADASTRO")
        m_NomeLogradouro = "" & Rs!tci_nome_logradouro
        m_Numero = "" & Rs!tci_numero
        m_Complemento = "" & Rs!tci_COMPLEMENTO
        Telefone = "" & Rs!tci_telefone
        Email = "" & Rs!tci_email
        m_Bairro = "" & Rs!tci_bairro
        m_Cep = "" & Rs!tci_cep
        m_Cidade = "" & Rs!tci_cidade
        m_Uf = "" & Rs!tci_UF
        m_DataCadastro = "" & Rs!tci_data_cadastro
        m_DataModificacao = "" & Rs!tci_data_modific
        m_CodSitCadastral = "" & Rs!tci_tsc_cod_sit_cad
        m_CodUsuario = "" & Rs!tci_tus_cod_usuario
        AtividadeTransporte = Nvl("" & Rs!TCI_TAE_CAE_TRANSPORTE, 0)
        m_Registro = "" & Rs!tci_registro
        If Not IsNull(Rs!TCI_FONE_FAX) And Trim(Rs!TCI_FONE_FAX) <> "" Then m_FoneFax = "" & Rs!TCI_FONE_FAX
        If Detalhes = True Then
            m_InicioAtividade = "" & Rs!tci_inicio_atividade
            m_CodGrupo = "" & Rs!tci_tga_cod_grupo
            m_CodNatureza = "" & Rs!tci_tnj_cod_natureza
            m_CodAtividade = "" & Rs!tci_tae_cae
            m_CodAtivPoder = "" & Rs!tci_tap_cod_ativ_poder
            m_Estabelecido = "" & Nvl("" & Rs!tci_estab, 0)
            If Not IsNull(Rs!tci_grupo_cae) Then
                m_GrupoAtividade = "" & Val(Rs!tci_grupo_cae)
            End If
            m_TipoContribuinte = "" & Rs!tci_tipo_contribuinte
            m_TipoRecolhimentoIss = "" & Rs!TCI_TIPO_RECOLHIMENTO_ISS
            m_Ruc = "" & Rs!tci_ruc
            m_Conselho = "" & Rs!tci_conselho
            m_ImovelProprio = "" & Nvl("" & Rs!tci_imovel_proprio, 0)
            m_NumEmpregado = "" & Nvl("" & Rs!tci_num_empregado, 0)
            m_PorteEmpresa = "" & Rs!tci_porte_empresa
            m_CodAtividadeSec = "" & Rs!tci_tae_cae_secund
            m_CodAtividadeTerc = Nvl("" & Rs!tci_tae_cae_terc, 0)
            m_NivelEscolar = Nvl("" & Rs!tci_nivel_escolar, 0)
            m_Protocolo = "" & Rs!tci_protocolo
            m_Ic = "" & Rs!tci_tim_ic
            If Not IsNull(Rs!tci_ISENTO) Then m_Isento = "" & Rs!tci_ISENTO
            If Not IsNull(Rs!TCI_FATOR_ALVARA) Then m_FatorAlvara = "" & Rs!TCI_FATOR_ALVARA
            If Not IsNull(Rs!TCI_TRA_COD_RAMO) Then m_CodRamo = "" & Rs!TCI_TRA_COD_RAMO
            If Not IsNull(Rs!TCI_CNH) And Trim(Rs!TCI_CNH) <> "" Then m_CNH = "" & Rs!TCI_CNH
            If Not IsNull(Rs!TCI_CATEGORIA) And Trim(Rs!TCI_CATEGORIA) <> "" Then m_Categoria = "" & Rs!TCI_CATEGORIA
            If Not IsNull(Rs!TCI_AUTORIZACAO) And Trim(Rs!TCI_AUTORIZACAO) <> "" Then m_Autorizacao = "" & Rs!TCI_AUTORIZACAO
            If Not IsNull(Rs!TCI_PONTO_RECEPCAO) And Trim(Rs!TCI_PONTO_RECEPCAO) <> "" Then m_PontoRecepcao = "" & Rs!TCI_PONTO_RECEPCAO
            If Not IsNull(Rs!TCI_THF_COD_HORARIO) And Trim(Rs!TCI_THF_COD_HORARIO) <> "" Then m_CodHorario = "" & Rs!TCI_THF_COD_HORARIO
            If Not IsNull(Rs!TCI_COD_LOGRADOURO) And Trim(Rs!TCI_COD_LOGRADOURO) <> "" Then m_CodLogradouro = "" & Rs!TCI_COD_LOGRADOURO
            If Not IsNull(Rs!TCI_COD_BAIRRO) And Trim(Rs!TCI_COD_BAIRRO) <> "" Then m_CodBairro = "" & Rs!TCI_COD_BAIRRO
            If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                Dim RsVariaveis As VSRecordset
                'PEGO OS DADOS DAS VARIAVEIS...
                Sql = "SELECT * FROM TAB_VARIAVEL_ECONOMICO WHERE TVE_TCI_IM      = " & Bdados.Converte(m_Im, tctexto)
                If Bdados.AbreTabela(Sql, RsVariaveis) Then
                    m_VariavelAnuncio = "" & RsVariaveis.Fields("TVE_ANUNCIOS")
                    m_VariavelDataFormatura = "" & RsVariaveis.Fields("TVE_DATA_FORMATURA")
                    m_VariavelFuncionarioSUS = "" & RsVariaveis.Fields("TVE_FUNCIONARIO_SUS")
                    m_VariavelQuantidadeItem = "" & RsVariaveis.Fields("TVE_QUANTIDADE_ITEM")
                End If
            End If
        End If
    End If
    
    Exit Function
TrataErro:
    If Err.Number = 13 Then 'VALOR NULO
        Resume Next
    Else
        Util.Erro Err.Description
        Exit Function
        Resume
    End If
End Function


Public Function Salvar() As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Campos = "tci_im,tci_im_auxiliar,tci_cgc_cpf, tci_nome, tci_fantasia, tci_logradouro, tci_nome_logradouro, tci_numero, tci_complemento," & _
                " tci_bairro, tci_cep, tci_cidade, tci_UF, tci_data_cadastro, tci_data_modific, tci_tsc_cod_sit_cad, tci_tus_cod_usuario," & _
                " tci_inicio_atividade, tci_tga_cod_grupo, tci_tnj_cod_natureza, tci_tae_cae, tci_tap_cod_ativ_poder, tci_estab, " & _
                " tci_grupo_cae, tci_tipo_contribuinte, tci_tipo_recolhimento_iss, tci_ruc, tci_conselho, " & _
                " tci_imovel_proprio, tci_registro, tci_num_empregado, tci_porte_empresa, tci_tae_cae_secund, tci_tae_cae_terc," & _
                " tci_nivel_escolar, tci_protocolo, tci_ISENTO, tci_FATOR_ALVARA, tci_TRA_COD_RAMO, tci_FONE_FAX, " & _
                " tci_CNH, tci_CATEGORIA, tci_AUTORIZACAO, tci_PONTO_RECEPCAO, tci_THF_COD_HORARIO, tci_COD_LOGRADOURO, tci_COD_BAIRRO,TCI_TAE_CAE_TRANSPORTE,tci_tipo_cadastro,tci_matriz_filial,tci_tim_ic,tci_telefone,tci_email,tci_rg,tci_inscricao_anterior,tci_observacao,tci_area_estabelecimento"
                
    If m_InicioPrestacaoServico <> "" Then Campos = Campos & " , TCI_INICIO_PREST_SERV"
    If m_Sit_Alvara <> "" Then Campos = Campos & " , TCI_SIT_ALVARA"
    
    'If m_Data_Reabertura <> "" Then Campos = Campos & "  , tci_data_reabertura "
    
    If Trim(Format(m_DataCadastro, "dd\mm\yyyy")) = "" Then m_DataCadastro = Format(Date, "dd/mm/yyyy")
    If Trim(Format(m_DataModificacao, "dd\mm\yyyy")) = "" Then m_DataModificacao = Format(Date, "dd/mm/yyyy")
    If Trim(Format(m_InicioAtividade, "dd\mm\yyyy")) = "" Then m_InicioAtividade = Format(Date, "dd/mm/yyyy")
    
    m_CodBairro = PegaCodBairro(m_Bairro)
    
    Valores = Bdados.PreparaValor(m_Im, m_ImAuxiliar, Bdados.Converte(m_CgcCpf, tctexto), m_Nome, m_Fantasia, m_Logradouro, m_NomeLogradouro, m_Numero, m_Complemento, _
                  m_Bairro, Bdados.Converte(m_Cep, tctexto), m_Cidade, m_Uf, m_DataCadastro, m_DataModificacao, m_CodSitCadastral, m_CodUsuario, _
                  m_InicioAtividade, m_CodGrupo, m_CodNatureza, m_CodAtividade, m_CodAtivPoder, m_Estabelecido, _
                  m_GrupoAtividade, m_TipoContribuinte, m_TipoRecolhimentoIss, m_Ruc, m_Conselho, m_ImovelProprio, _
                  m_Registro, m_NumEmpregado, m_PorteEmpresa, m_CodAtividadeSec, m_CodAtividadeTerc, _
                  m_NivelEscolar, m_Protocolo, m_Isento, m_FatorAlvara, m_CodRamo, m_FoneFax, _
                  m_CNH, m_Categoria, m_Autorizacao, m_PontoRecepcao, m_CodHorario, m_CodLogradouro, m_CodBairro, AtividadeTransporte, m_TipoCadastro, m_Matriz_Filial, Bdados.Converte(m_Ic, tctexto), m_FoneFax, Email, m_Rg, m_ImAnterior, m_Obs, m_AreaEstabelecimento)
                  
   ' If m_Data_Encerramento <> "" Then Valores = Valores & Bdados.PreparaValor(Bdados.Converte(m_Data_Encerramento, TCDataHora))
   ' If m_Data_Reabertura <> "" Then Valores = Valores & Bdados.Converte(m_Data_Reabertura, TCDataHora)
    
    If m_InicioPrestacaoServico <> "" Then Valores = Valores & Bdados.PreparaValor(Bdados.Converte(m_InicioPrestacaoServico, TCDataHora))
    If m_Sit_Alvara <> "" Then Valores = Valores & Bdados.PreparaValor(Bdados.Converte(m_Sit_Alvara, tctexto))
     
    
    If Trim(m_Im) <> "" Then
        Condicao = "tci_im = '" & m_Im & "'"
    ElseIf Trim(m_CgcCpf) <> "" Then
        Condicao = "tci_cgc_cpf = '" & m_CgcCpf & "'"
    End If
    
    If Bdados.GravaDados("Tab_contribuinte", Valores, Campos, Condicao) Then
        'Gravo o Historico...
        Campos = Campos & ",TC_DATA,TCI_MOTIVO"
        Valores = Valores & Bdados.PreparaValor(M_Data, Bdados.Converte(m_Motivo, tctexto))
        Bdados.InsereDados "Tab_contribuinte_Historico", Valores, Campos
        If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
            If m_Tela <> "" Then
                Call Grava_Movimento
            End If
            Call Grava_Variaveis
        End If
        Salvar = True
    End If
    Exit Function
    Resume
End Function
Public Function PegaCodBairro(NomeBairro As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "Select tba_cod_bairro from tab_bairro where tba_nome = '" & NomeBairro & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        PegaCodBairro = Rs.Fields("tba_cod_bairro")
    End If
    
End Function

Private Function Grava_Movimento() As Boolean
    Dim Campos     As String
    Dim Valores    As String
    Dim Processo   As String
    Dim Conta      As ContaCorrente
    
    Set Conta = New ContaCorrente
    Processo = Conta.GeraCodPagamento(83)
    
    Campos = "TMC_CODIGO_MOVIMENTACAO,TMC_TIPO,TMC_TCI_IM,TMC_DATA"
    Valores = Bdados.PreparaValor(Processo, Bdados.Converte("1", tctexto), Bdados.Converte(m_Im, tctexto), m_InicioAtividade)
    If Bdados.GravaDados("TAB_MOVIMENTO_CONTRIBUINTE", Valores, Campos, "TMC_CODIGO_MOVIMENTACAO = '" & Processo & "'") Then
        'ALTERO OS CAMPOS TMC_LIBERACAO_ALVARA,TMC_LIBERACAO_SANITARIA
        Campos = " TMC_LIBERACAO_ALVARA,TMC_LIBERACAO_SANITARIA"
        Valores = Bdados.PreparaValor(Bdados.Converte("0", tctexto), Bdados.Converte("0", tctexto))
        Grava_Movimento = Bdados.GravaDados("TAB_MOVIMENTO_CONTRIBUINTE", Valores, Campos, "TMC_CODIGO_MOVIMENTACAO = '" & Processo & "'")
    End If
End Function
Private Function Grava_Variaveis() As Boolean
    Dim Valores  As String
    Dim Campos   As String
    Dim Condicao As String
    
    'TVE_TCI_IM
    'TVE_ANUNCIOS
    'TVE_QUANTIDADE_ITEM
    'TVE_DATA_FORMATURA
    'TVE_FUNCIONARIO_SUS
    '======================================================================================================================================
    
    Campos = "TVE_TCI_IM,TVE_ANUNCIOS,TVE_QUANTIDADE_ITEM,TVE_FUNCIONARIO_SUS"
    Valores = Bdados.PreparaValor(m_Im, Bdados.Converte(m_VariavelAnuncio, TCMonetario), m_VariavelQuantidadeItem, m_VariavelFuncionarioSUS)
    If m_VariavelDataFormatura <> "" Then
        Campos = Campos & " , TVE_DATA_FORMATURA"
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(m_VariavelDataFormatura, TCDataHora))
    End If
    
    
    Condicao = "TVE_TCI_IM =  " & Bdados.Converte(m_Im, tctexto)
    Grava_Variaveis = Bdados.GravaDados("TAB_VARIAVEL_ECONOMICO", Valores, Campos, Condicao)
    Bdados.GravaDados "TAB_VARIAVEL_ECONOMICO_HIST", Valores, Campos, Condicao
    If Grava_Variaveis = False Then
        Erro "Erro ao gravar variáveis"
    End If
    '======================================================================================================================================
End Function
Public Function SalvarReduzido() As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    Campos = "tci_im,tci_cgc_cpf, tci_nome, tci_fantasia, tci_logradouro, tci_nome_logradouro, tci_numero, tci_complemento," & _
                " tci_bairro, tci_cep, tci_cidade, tci_UF, tci_data_modific,tci_tus_cod_usuario,tci_tsc_cod_sit_cad,tci_area_estabelecimento,tci_tae_cae"
                
    If Trim(m_DataModificacao) = "" Then m_DataModificacao = Format(Date, "dd/mm/yyyy")
    Valores = Bdados.PreparaValor(m_Im, m_CgcCpf, m_Nome, m_Fantasia, m_Logradouro, m_NomeLogradouro, m_Numero, m_Complemento, _
                  m_Bairro, m_Cep, m_Cidade, m_Uf, m_DataModificacao, AplicacoesVTFuncoes.Usuario, m_CodSitCadastral, m_AreaEstabelecimento, m_CodAtividade)
    If Trim(m_Im) <> "" Then
        Condicao = "tci_im = '" & m_Im & "'"
    ElseIf Trim(m_CgcCpf) <> "" Then
        Condicao = "tci_cgc_cpf = '" & m_CgcCpf & "'"
    End If
    If Trim(Telefone) <> "" Then
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Telefone, tctexto))
        Campos = Campos & ",tci_telefone"
    End If
    If Trim(Email) <> "" Then
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Email, tctexto))
        Campos = Campos & ",tci_email"
    End If
    If Bdados.GravaDados("Tab_contribuinte", Valores, Campos, Condicao) Then
        SalvarReduzido = True
    End If
End Function


Public Function Excluir(Im As String) As Boolean
    If Bdados.DeletaDados("tab_contribuinte", "tci_im = '" & Im & "' or tci_cgc_cpf = '" & Im & "'") Then Excluir = True
End Function

Public Function PreencherGrd(Grid As Object, Optional Im As String, Optional Nome As String, Optional TipoContrib As Byte = 0, Optional CpfCgc As String) As Boolean
    Dim Rs As VSRecordset
    Dim Sql As String
    Sql = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf as CPF_CGC from Tab_Contribuinte where tci_tsc_cod_sit_cad =1"
    If Trim(Nome) <> "" Then
        Sql = Sql & " AND tci_nome like '%" & Nome & "%'"
    End If
    If Trim(Im) <> "" Then
        Sql = Sql & " and tci_im = '" & Im & "' or tci_cgc_cpf = '" & Im & "'"
    End If
    If TipoContrib <> 0 Then
        Sql = Sql & " and tci_tipo_contribuinte = " & TipoContrib
    End If
    If CpfCgc <> "" Then
        Sql = Sql & " and tci_cgc_cpf = '" & CpfCgc & "'"
    End If
    
    If Grid.Preencher(Bdados, Sql, 1400, 5000, 3000) Then PreencherGrd = True
End Function

Public Function PreencherComboTipoLogr(Combo As Object) As Boolean
    Dim Sql As String
    Sql = "Select DISTINCT(ttl_nome) From Tab_Tipo_Logr"
    If Combo.Preencher(Bdados, Sql) Then PreencherComboTipoLogr = True
End Function

Public Function VerificaTEMImovel(InsMunicipal As String) As Boolean
    Dim Sql As String                   'VERIFICA SE TEM IMOVEL CADASTRADO PRA INXCRIÇÃO MUNICIPAL
    Dim Rs As VSRecordset
    Sql = "Select tim_ic from Tab_Imovel where tim_tci_im = '" & InsMunicipal & "'"
    If Bdados.AbreTabela(Sql, Rs) Then VerificaTEMImovel = True
End Function

Public Function VerificaTEMDebito(InsMunicipal As String) As Boolean
    Dim Sql As String                   'VERIFICA SE TEM debito PRA INXCRIÇÃO MUNICIPAL
    Dim Rs As VSRecordset
    Sql = "Select * FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_INSCRICAO = '" & InsMunicipal & "' and TOC_STATUS_OBRIGACAO = 2 "
    If Bdados.AbreTabela(Sql, Rs) Then VerificaTEMDebito = True
End Function

Public Sub PreencherCboNaturezaJur(Combo As Object)
    Dim Sql As String
    Sql = "Select tnj_natureza, TNJ_COD_NATUREZA From Tab_Natureza_Juridica"
    Combo.Preencher Bdados, Sql
End Sub
    
Public Sub PreencherCboClasseAtividade(Combo As Object)
    Dim Sql As String
    Sql = "Select tga_nome, tga_cod_grupo From Tab_Grupo_Atividade"
    Combo.Preencher Bdados, Sql
End Sub

Public Sub PreencherCboAtividadePoder(Combo As Object)
    Dim Sql As String
    Sql = "Select tap_poder, tap_cod_ativ_poder From Tab_atividade_Poder"
    Combo.Preencher Bdados, Sql
End Sub

Public Sub PreencherCboAtividadeEcon(Combo As Object)
    Dim Sql As String
    Sql = "Select tae_nome, TAE_CAE From Tab_atividade_Economica"
    Combo.Preencher Bdados, Sql
End Sub

Public Sub PreencherCboSitCad(Combo As Object)
    Dim Sql As String
    Sql = "Select tsc_nome, tsc_cod_sit_cad  From Tab_sit_cadastral order by tsc_cod_sit_cad"
    Combo.Preencher Bdados, Sql
End Sub


Public Function ExibirFicha(grd As Object, Im As String, Contribuinte As String, Optional PeriodoInicial As String, Optional PeriodoFinal As String, Optional Situacao As String, Optional NFInicio As String, Optional NFFim As String) As Boolean
    On Error GoTo trata
    Dim Sql As String, Pos As Integer
    
    Select Case Situacao
        Case "NÃO PAGO"
            Sql = "SELECT "
            Sql = Sql & Bdados.ParteTexto("tgt_periodo", MidVs, 5, 2, True) & Bdados.Concatena & "'/'" & Bdados.Concatena & Bdados.ParteTexto("tgt_periodo", MidVs, 1, 4, True) & " as Periodo,"
            Sql = Sql & " tgt_cod_pagamento as DAM,"
            Sql = Sql & " tip_sigla_imposto AS Tributo,"
            Sql = Sql & " tdd_num_nota_inicial AS NFInicial,"
            Sql = Sql & " tdd_num_nota_final as NFFinal,"
            Sql = Sql & FuncaoReal("tdd_total_nota") & " as TotalNF,"
            Sql = Sql & FuncaoReal("tdd_total_material_reducao") & " as Deduções,"
            Sql = Sql & FuncaoReal("tgt_valor_tributo") & " as Original,"
            Sql = Sql & "'' as Juros,"
            Sql = Sql & "'' as Multa,"
            Sql = Sql & "'' as DataPagto,"
            Sql = Sql & "'' as ValorPago,"
            Sql = Sql & "tdd_obs As Obs"
            Sql = Sql & " From TAB_GERACAO_TRIBUTO, Tab_Detalhe_Dam, TAB_IMPOSTO"
            Sql = Sql & " Where tgt_cod_pagamento *= tdd_tgt_cod_pagamento"
            Sql = Sql & " AND tgt_tip_cod_imposto=tip_cod_imposto"
            Sql = Sql & " AND tgt_cod_pagamento NOT IN (SELECT tdr_tgt_cod_pagamento FROM TAB_DARM_RECEBIDO WHERE tdr_im=tgt_im AND tdr_tip_cod_imposto=tgt_tip_cod_imposto AND tdr_periodo=tgt_periodo)"
            Sql = Sql & " AND tgt_tip_cod_imposto LIKE '111305%'"
            Sql = Sql & " AND tgt_im='" & Im & "'"
        Case "PAGO"
            Sql = "SELECT "
            Sql = Sql & Bdados.ParteTexto("tgt_periodo", MidVs, 5, 2, True) & Bdados.Concatena & "'/'" & Bdados.Concatena & Bdados.ParteTexto("tgt_periodo", MidVs, 1, 4, True) & " as Periodo,"
            Sql = Sql & " tgt_cod_pagamento as DAM,"
            Sql = Sql & " tip_sigla_imposto AS Tributo,"
            Sql = Sql & " tdd_num_nota_inicial AS NFInicial,"
            Sql = Sql & " tdd_num_nota_final as NFFinal,"
            Sql = Sql & FuncaoReal("tdd_total_nota") & " as TotalNF,"
            Sql = Sql & FuncaoReal("tdd_total_material_reducao") & " as Deduções,"
            Sql = Sql & FuncaoReal("tgt_valor_tributo") & " as Original,"
            Sql = Sql & FuncaoReal("tdr_valor_real_juros") & " as Juros,"
            Sql = Sql & FuncaoReal("tdr_valor_real_multa") & " as Multa,"
            Sql = Sql & "tdr_data_pagamento as DataPagto,"
            Sql = Sql & FuncaoReal("tdr_valor_real_pago") & " as ValorPago,"
            Sql = Sql & "tdd_obs As Obs"
            Sql = Sql & " From TAB_GERACAO_TRIBUTO, Tab_Detalhe_Dam, TAB_DARM_RECEBIDO, TAB_IMPOSTO"
            Sql = Sql & " Where tgt_cod_pagamento *= tdd_tgt_cod_pagamento"
            Sql = Sql & " AND tgt_tip_cod_imposto=tip_cod_imposto"
            Sql = Sql & " AND tgt_cod_pagamento = tdr_tgt_cod_pagamento"
            Sql = Sql & " AND tgt_tip_cod_imposto LIKE '111305%"
            Sql = Sql & " AND tgt_im='" & Im & "'"
        Case Else
            Sql = "SELECT "
            Sql = Sql & Bdados.ParteTexto("tgt_periodo", MidVs, 5, 2, True) & Bdados.Concatena & "'/'" & Bdados.Concatena & Bdados.ParteTexto("tgt_periodo", MidVs, 1, 4, True) & " as Periodo,"
            Sql = Sql & " tgt_cod_pagamento as DAM,"
            Sql = Sql & " tip_sigla_imposto AS Tributo,"
            Sql = Sql & " tdd_num_nota_inicial AS NFInicial,"
            Sql = Sql & " tdd_num_nota_final as NFFinal,"
            Sql = Sql & FuncaoReal("tdd_total_nota") & " as TotalNF,"
            Sql = Sql & FuncaoReal("tdd_total_material_reducao") & " as Deduções,"
            Sql = Sql & FuncaoReal("tgt_valor_tributo") & " as Original,"
            Sql = Sql & FuncaoReal("tdr_valor_real_juros") & " as Juros,"
            Sql = Sql & FuncaoReal("tdr_valor_real_multa") & " as Multa,"
            Sql = Sql & "tdr_data_pagamento as DataPagto,"
            Sql = Sql & FuncaoReal("tdr_valor_real_pago") & " as ValorPago,"
            Sql = Sql & "tdd_obs As Obs"
            Sql = Sql & " From TAB_GERACAO_TRIBUTO, Tab_Detalhe_Dam, TAB_DARM_RECEBIDO, TAB_IMPOSTO"
            Sql = Sql & " Where tgt_cod_pagamento *= tdd_tgt_cod_pagamento"
            Sql = Sql & " AND tgt_tip_cod_imposto=tip_cod_imposto"
            Sql = Sql & " AND tgt_cod_pagamento *= tdr_tgt_cod_pagamento"
            Sql = Sql & " AND tgt_tip_cod_imposto LIKE '111305%'"
            Sql = Sql & " AND tgt_im='" & Im & "'"
    End Select
    
    If PeriodoInicial <> "" Then
        Sql = Sql & " AND tgt_periodo>=" & Util.ParseString(PeriodoInicial, "/", 2) & Util.ParseString(PeriodoInicial, "/", 1)
    End If
    If PeriodoFinal <> "" Then
        Sql = Sql & " AND tgt_periodo<=" & Util.ParseString(PeriodoFinal, "/", 2) & Util.ParseString(PeriodoFinal, "/", 1)
    End If
    If NFInicio <> "" Then
        Sql = Sql & " AND tdd_num_nota_inicial>=" & NFInicio
    End If
    If NFFim <> "" Then
        Sql = Sql & " AND tdd_num_nota_final<=" & NFFim
    End If
    Sql = Sql & " ORDER BY tgt_periodo, tgt_tip_cod_imposto"
    If grd.Preencher(Bdados, Sql) Then
        ExibirFicha = True
        grd.Mensagem = "Total lançado : " & Format(grd.Colunas(8).Soma, "currency") & " x Total arrecadado : " & Format(grd.Colunas(9).Soma + grd.Colunas(10).Soma + grd.Colunas(12).Soma, "currency") & " (" & Format((grd.Colunas(9).Soma + grd.Colunas(10).Soma + grd.Colunas(12).Soma) / IIf(grd.Colunas(8).Soma = 0, 1, grd.Colunas(8).Soma), "percent") & ")"
    End If
    Exit Function
trata:
    Erro Err.Description
    Exit Function
    Resume
End Function

Public Function GravarHistorico(CodMudanca As String, Im As String) As Boolean
    Dim Sql As String
    Sql = "Select tab_contribuinte.* , " & Bdados.Converte(Date, TCDataHora) & ",'" & _
            Format(Time, "HH:MM:SS") & "'," & CodMudanca & _
            " from tab_contribuinte where tci_im = '" & Im & "'"
    If Bdados.Executa("Insert into tab_contribuinte_historico " & Sql) Then GravarHistorico = True
End Function

Public Function BuscarContribuintes(grd As Object, Optional Im As String, Optional Cgc As String, Optional Razao As String, Optional Fantasia As String, _
                        Optional TipoLogr As String, Optional Logr As String, Optional Bairro As String, Optional Numero As String, _
                        Optional Compl As String, Optional Cidade As String, Optional AtivEconomica As String, Optional ObrigIss As String, _
                        Optional Periodo1 As String, Optional Periodo2 As String, Optional Ponto As String, Optional ByRef RetFiltroRpt As String, Optional IAnterior As String, Optional TipoCadastro As String, Optional SitCadastro As String, Optional DataFormatura As String, Optional FuncionarioSUS As String, Optional PorteEmpresa As String) As Boolean
    Dim Sql As String
    Dim Condicao As String
    Dim FiltroRpt As String
    'alterado por henrique/ 30 jan de 2014
    'consulta com where onde deveria existir um left join
        Sql = "SELECT tci_im as IM, tci_cgc_cpf as CNPJ, tci_nome as Razão, TCI_FANTASIA AS Fantasia," & _
                " tci_logradouro as Logr, tci_nome_logradouro as Nome, tci_numero as [Nº], tci_complemento as Complemento," & _
                " tci_bairro as Bairro,tci_cidade as Cidade,tci_uf as UF, Tab_Atividade_Economica.tae_nome as [Atividade Economica]," & _
                " tci_autorizacao as Autorizacao, tci_fator_alvara as [Fator], tci_INICIO_ATIVIDADE as [Abertura]," & _
                " tci_tus_cod_usuario as Usuario,tci_tipo_cadastro as [Tipo Cadastro],tci_IM_Auxiliar as [Insc Municipal]" & _
                " FROM Tab_Contribuinte left join Tab_Atividade_Economica" & _
                " on Tab_Contribuinte.tci_tae_cae=Tab_Atividade_Economica.tae_cae" & _
                " where 1=1"
            'FiltroRpt = " {TAB_CONTRIBUINTE.tci_tsc_cod_sit_cad} = 1 and {TAB_CONTRIBUINTE.tci_tipo_contribuinte} >0 AND "
            FiltroRpt = "  1 = 1 "
            If DataFormatura <> "" Then
                Sql = Sql & " and TVE_DATA_FORMATURA = " & Bdados.Converte(DataFormatura, TCDataHora)
                FiltroRpt = FiltroRpt & " and {TAB_VARIAVEL_ECONOMICO.TVE_DATA_FORMATURA}  <= #" & DataFormatura & "#"
            End If
            If PorteEmpresa <> "" Then
                Sql = Sql & " and TCI_PORTE_EMPRESA = '" & PorteEmpresa & "'"
                FiltroRpt = FiltroRpt & " and {TAB_CONTRIBUINTE.TCI_PORTE_EMPRESA}  = " & PorteEmpresa
            End If
            If FuncionarioSUS <> "" Then
                Sql = Sql & " and TVE_FUNCIONARIO_SUS = '" & FuncionarioSUS & "'"
                FiltroRpt = FiltroRpt & "   and  {VIS_FUNCIONARIO_SUS.tge_codigo} = " & FuncionarioSUS
            End If
            If Trim(Cgc) <> "" Then
                Condicao = " and tci_cgc_cpf ='" & Cgc & "'"
                FiltroRpt = FiltroRpt & " and {TAB_CONTRIBUINTE.tci_cgc_cpf} ='" & Cgc & "'"
            End If
            If Trim(Im) <> "" Then
                Condicao = " and tci_im ='" & Im & "'"
                FiltroRpt = FiltroRpt & " and  {TAB_CONTRIBUINTE.tci_im} ='" & Im & "'"
            End If
            If Trim(IAnterior) <> "" Then
                Condicao = " and tci_IM_Auxiliar ='" & Imposto.FormataInscricao(IAnterior, InscContrib) & "'"
                FiltroRpt = FiltroRpt & " and  {TAB_CONTRIBUINTE.tci_IM_Auxiliar} ='" & Imposto.FormataInscricao(IAnterior, InscContrib) & "'"
            End If
            If Trim(Razao) <> "" Then
                Condicao = Condicao & " and (tci_nome like '" & UCase(Razao) & "%')"
                FiltroRpt = FiltroRpt & " and  ({TAB_CONTRIBUINTE.tci_nome} like '" & UCase(Razao) & "*')"
            End If
            If Trim(TipoCadastro) <> "" Then
                Condicao = Condicao & " and tci_tipo_cadastro ='" & TipoCadastro & "'"
                FiltroRpt = FiltroRpt & "  and {TAB_CONTRIBUINTE.tci_tipo_cadastro} = " & TipoCadastro & ""
            End If
            If SitCadastro <> "" Then
                Condicao = Condicao & " AND TCI_TSC_COD_SIT_CAD = '" & SitCadastro & "'"
                FiltroRpt = FiltroRpt & " and  {TAB_CONTRIBUINTE.TCI_TSC_COD_SIT_CAD} = " & SitCadastro & " "
            End If
            If Trim(Fantasia) <> "" Then
                Condicao = Condicao & " and tci_fantasia like '" & UCase(Fantasia) & "%'"
                FiltroRpt = FiltroRpt & " and {TAB_CONTRIBUINTE.tci_fantasia} like '*" & UCase(Fantasia) & "*'"
            End If
            If Trim(TipoLogr) <> "" Then
                Condicao = Condicao & " and tci_logradouro = '" & TipoLogr & "'"
                FiltroRpt = FiltroRpt & "  and {TAB_CONTRIBUINTE.tci_logradouro} =  '" & TipoLogr & "'"
            End If
            If Trim(Logr) <> "" Then
                Condicao = Condicao & " and tci_nome_logradouro like '" & UCase(Logr) & "%'"
                FiltroRpt = FiltroRpt & "  and {TAB_CONTRIBUINTE.tci_nome_logradouro} like '*" & UCase(Logr) & "*'"
            End If
            If Trim(Bairro) <> "" Then
                Condicao = Condicao & " and tci_bairro like '" & UCase(Bairro) & "%'"
                FiltroRpt = FiltroRpt & " and  {TAB_CONTRIBUINTE.tci_bairro} like '*" & UCase(Bairro) & "*'"
            End If
            If Trim(Numero) <> "" Then
                Condicao = Condicao & " and tci_numero = '" & Numero & "'"
                FiltroRpt = FiltroRpt & "  and {TAB_CONTRIBUINTE.tci_numero} = '" & Numero & "'"
            End If
            If Trim(Compl) <> "" Then
                Condicao = Condicao & " and tci_complemento like '" & UCase(Compl) & "%'"
                FiltroRpt = FiltroRpt & " and  {TAB_CONTRIBUINTE.tci_complemento} like '*" & UCase(Compl) & "*'"
            End If
            If Trim(Cidade) <> "" Then
                Condicao = Condicao & " and TCi_CIDADE ='" & UCase(Cidade) & "'"
                FiltroRpt = FiltroRpt & "  and {TAB_CONTRIBUINTE.TCI_CIDADE} = '" & UCase(Cidade) & "'"
            End If
            If Trim(AtivEconomica) <> "" Then
                Condicao = Condicao & " and tci_tae_cae =" & AtivEconomica
                FiltroRpt = FiltroRpt & " and  {Tab_Contribuinte.tci_tae_cae} =" & AtivEconomica
            End If
            If Trim(ObrigIss) <> "" Then
                Condicao = Condicao & " and tci_tipo_recolhimento_iss =" & ObrigIss
                FiltroRpt = FiltroRpt & " and  {TAB_CONTRIBUINTE.tci_tipo_recolhimento_iss} =" & ObrigIss
            End If
            If Trim(Periodo1) <> "" And Trim(Periodo2) <> "" Then
                Condicao = Condicao & " and tci_data_cadastro >= " & Bdados.Converte(Periodo1, TCDataHora) & " And tci_data_cadastro <= " & Bdados.Converte(Periodo2, TCDataHora)
                FiltroRpt = FiltroRpt & "  and {TAB_CONTRIBUINTE.tci_data_cadastro} >=" & Bdados.Converte(Periodo1, TCDataHora) & " AND {TAB_CONTRIBUINTE.tci_data_cadastro} <= " & Bdados.Converte(Periodo2, TCDataHora)
            End If
            If Trim$(Ponto) <> "" Then
                Condicao = Condicao & " and TCI_PONTO_RECEPCAO='" & Ponto & "'"
                FiltroRpt = FiltroRpt & " and  {TAB_CONTRIBUINTE.TCI_PONTO_RECEPCAO} ='" & Ponto & "'"
            End If
            
            Sql = Sql & Condicao
            RetFiltroRpt = FiltroRpt
            If grd.Preencher(Bdados, Sql) Then BuscarContribuintes = True
End Function

Public Function BuscarContribuintesHistorico(grd As Object, Optional Im As String, Optional Cgc As String, Optional Razao As String, Optional Fantasia As String, _
                        Optional TipoLogr As String, Optional Logr As String, Optional Bairro As String, Optional Numero As String, _
                        Optional Compl As String, Optional Cidade As String, Optional AtivEconomica As String, Optional ObrigIss As String, _
                        Optional Periodo1 As String, Optional Periodo2 As String, Optional Ponto As String, Optional ByRef RetFiltroRpt As String, Optional IAnterior As String, Optional TipoCadastro As String, Optional SitCadastro As String) As Boolean
    Dim Sql As String
    Dim Condicao As String
    Dim FiltroRpt As String
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "SELECT tci_im as IM, tci_cgc_cpf as CNPJ, tci_nome as Razão, TCI_FANTASIA AS Fantasia," & _
                " tci_logradouro as Logr, tci_nome_logradouro as Nome, tci_numero as [Nº], tci_complemento as Complemento," & _
                " tci_bairro as Bairro,tci_cidade as Cidade,tci_uf as UF, tae_nome as [Atividade Economica]," & _
                " tci_autorizacao as Autorizacao, tci_fator_alvara as [Fator], tci_INICIO_ATIVIDADE as [Abertura]," & _
                " tci_tus_cod_usuario as Usuario,tci_tipo_cadastro as [Tipo Cadastro],tci_IM_Auxiliar as [Insc Municipal]" & _
                " FROM Tab_Contribuinte_Historico,Tab_Atividade_Economica" & _
                " where tci_tae_cae=tae_cae"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "SELECT tci_im as IM, tci_cgc_cpf as CNPJ, tci_nome as Razão, TCI_FANTASIA AS Fantasia,"
        Sql = Sql & " tci_logradouro as Logr, tci_nome_logradouro as Nome, tci_numero as Número, tci_complemento as Complemento,"
        Sql = Sql & " tci_bairro as Bairro,tci_cidade as Cidade,tci_uf as UF, tae_nome as Atividade_Economica,"
        Sql = Sql & "  tci_autorizacao as Autorizacao, tci_fator_alvara as Fator, tci_INICIO_ATIVIDADE as Abertura,"
        Sql = Sql & " tci_tus_cod_usuario as Usuario,tci_tipo_cadastro as Tipo_Cadastro,tci_IM_Auxiliar as Insc_Municipal,TCI_DATA_CADASTRO AS Data_Cadastro"
        If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
            Sql = Sql & " ,TVE_ANUNCIOS AS Anúncios,"
            Sql = Sql & " TVE_QUANTIDADE_ITEM as QTD,"
            Sql = Sql & " TVE_DATA_FORMATURA as Formatura,"
            Sql = Sql & " vis_funcionario_sus.tge_nome as Funcionário_SUS"
        End If
        Sql = Sql & " ,TC_DATA AS Data,tci_motivo as Motivo"
        Sql = Sql & " FROM  Tab_Contribuinte_Historico,Tab_Atividade_Economica"
        If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
            Sql = Sql & " ,TAB_VARIAVEL_ECONOMICO,vis_funcionario_sus"
            Sql = Sql & " where tci_tae_cae = Tab_Atividade_Economica.tae_cae  "
            Sql = Sql & " AND tci_im = TVE_TCI_IM"
            Sql = Sql & " and TVE_FUNCIONARIO_SUS = vis_funcionario_sus.tge_codigo"
        Else
            Sql = Sql & " where tci_tae_cae = tae_cae "
        End If
      
    End If
            'FiltroRpt = " {TAB_CONTRIBUINTE.tci_tsc_cod_sit_cad} = 1 and {TAB_CONTRIBUINTE.tci_tipo_contribuinte} >0 AND "
            FiltroRpt = "  1 = 1 "
            If Trim(Cgc) <> "" Then
                Condicao = " and tci_cgc_cpf ='" & Cgc & "'"
                FiltroRpt = FiltroRpt & " AND  {TAB_CONTRIBUINTE_HISTORICO.tci_cgc_cpf} ='" & Cgc & "'"
            End If
            If Trim(Im) <> "" Then
                Condicao = " and tci_im ='" & Im & "'"
                FiltroRpt = FiltroRpt & " AND  {TAB_CONTRIBUINTE_HISTORICO.tci_im} ='" & Im & "'"
            End If
            If Trim(IAnterior) <> "" Then
                Condicao = " and tci_IM_Auxiliar ='" & Imposto.FormataInscricao(IAnterior, InscContrib) & "'"
                FiltroRpt = FiltroRpt & " AND  {TAB_CONTRIBUINTE_HISTORICO.tci_IM_Auxiliar} ='" & Imposto.FormataInscricao(IAnterior, InscContrib) & "'"
            End If
            If Trim(Razao) <> "" Then
                Condicao = Condicao & " and (tci_nome like '" & UCase(Razao) & "%')"
                FiltroRpt = FiltroRpt & " AND  {TAB_CONTRIBUINTE_HISTORICO.tci_nome} like '" & UCase(Razao) & "*'"
            End If
            If Trim(TipoCadastro) <> "" Then
                Condicao = Condicao & " and tci_tipo_cadastro ='" & TipoCadastro & "'"
                FiltroRpt = FiltroRpt & " AND  {TAB_CONTRIBUINTE_HISTORICO.tci_tipo_cadastro} = " & TipoCadastro
            End If
            If SitCadastro <> "" Then
                Condicao = Condicao & " AND TCI_TSC_COD_SIT_CAD = '" & SitCadastro & "'"
                FiltroRpt = FiltroRpt & "  AND {TAB_CONTRIBUINTE_HISTORICO.TCI_TSC_COD_SIT_CAD} = " & SitCadastro
            End If
            If Trim(Fantasia) <> "" Then
                Condicao = Condicao & " and tci_fantasia like '" & UCase(Fantasia) & "%'"
                FiltroRpt = FiltroRpt & "  AND {TAB_CONTRIBUINTE_HISTORICO.tci_fantasia} like '*" & UCase(Fantasia) & "*'"
            End If
            If Trim(TipoLogr) <> "" Then
                Condicao = Condicao & " and tci_logradouro = '" & TipoLogr & "'"
                FiltroRpt = FiltroRpt & " AND  {TAB_CONTRIBUINTE_HISTORICO.tci_logradouro} =  '" & TipoLogr & "'"
            End If
            If Trim(Logr) <> "" Then
                Condicao = Condicao & " and tci_nome_logradouro like '" & UCase(Logr) & "%'"
                FiltroRpt = FiltroRpt & "  AND {TAB_CONTRIBUINTE_HISTORICO.tci_nome_logradouro} like '*" & UCase(Logr) & "*'"
            End If
            If Trim(Bairro) <> "" Then
                Condicao = Condicao & " and tci_bairro like '" & UCase(Bairro) & "%'"
                FiltroRpt = FiltroRpt & "  AND {TAB_CONTRIBUINTE_HISTORICO.tci_bairro} like '*" & UCase(Bairro) & "*'"
            End If
            If Trim(Numero) <> "" Then
                Condicao = Condicao & " and tci_numero = '" & Numero & "'"
                FiltroRpt = FiltroRpt & " AND  {TAB_CONTRIBUINTE_HISTORICO.tci_numero} = '" & Numero & "'"
            End If
            If Trim(Compl) <> "" Then
                Condicao = Condicao & " and tci_complemento like '" & UCase(Compl) & "%'"
                FiltroRpt = FiltroRpt & "  AND {TAB_CONTRIBUINTE_HISTORICO.tci_complemento} like '*" & UCase(Compl) & "*'"
            End If
            If Trim(Cidade) <> "" Then
                Condicao = Condicao & " and TCi_CIDADE ='" & UCase(Cidade) & "'"
                FiltroRpt = FiltroRpt & "  AND {TAB_CONTRIBUINTE_HISTORICO.TCI_CIDADE} = '" & UCase(Cidade) & "'"
            End If
            If Trim(AtivEconomica) <> "" Then
                Condicao = Condicao & " and tci_tae_cae =" & AtivEconomica
                FiltroRpt = "  AND {TAB_CONTRIBUINTE_HISTORICO.tci_tae_cae} =" & AtivEconomica
            End If
            If Trim(ObrigIss) <> "" Then
                Condicao = Condicao & " and tci_tipo_recolhimento_iss =" & ObrigIss
                FiltroRpt = FiltroRpt & " AND  {TAB_CONTRIBUINTE_HISTORICO.tci_tipo_recolhimento_iss} =" & ObrigIss
            End If
            If Trim(Periodo1) <> "" And Trim(Periodo2) <> "" Then
                Condicao = Condicao & " and tci_data_cadastro >= " & Bdados.Converte(Periodo1, TCDataHora) & " And tci_data_cadastro <= " & Bdados.Converte(Periodo2, TCDataHora)
                FiltroRpt = FiltroRpt & "  AND {TAB_CONTRIBUINTE.tci_data_cadastro} >=" & Bdados.Converte(Periodo1, TCDataHora) & " AND {TAB_CONTRIBUINTE.tci_data_cadastro} <= " & Bdados.Converte(Periodo2, TCDataHora)
            End If
            If Trim$(Ponto) <> "" Then
                Condicao = Condicao & " and TCI_PONTO_RECEPCAO='" & Ponto & "'"
                FiltroRpt = FiltroRpt & " AND  {TAB_CONTRIBUINTE.TCI_PONTO_RECEPCAO} ='" & Ponto & "'"
            End If
            
            Sql = Sql & Condicao
            RetFiltroRpt = FiltroRpt
            If grd.Preencher(Bdados, Sql) Then BuscarContribuintesHistorico = True
End Function

Public Function BuscarAcessoEmtDam() As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "SELECT * FROM TAB_ACESSO_USUARIO WHERE TAU_TMO_COD_MODULO ='TCOB' and TAU_TFO_COD_FORMULARIO =102 AND TAU_TUS_COD_USUARIO='" & Aplicacoes.Usuario & "'"
    If Bdados.AbreTabela(Sql, Rs) Then BuscarAcessoEmtDam = True
End Function
Public Function Grava_Atividade_Secundaria(Grid As Object, Im As String) As Boolean
    Dim i                                       As Integer
    Dim RegistrosAfetados                       As Integer
    Dim Sql                                     As String
    Dim Valores                                 As String
    Dim Condicao                                As String
    Dim Campos                                  As String
    Dim PosGrupo                                As Integer
    Dim PosAtividade                            As Integer
            
    Condicao = "TAS_TCI_IM = '" & Im & "'"
    Bdados.AbreTrans
    If Bdados.DeletaDados("TAB_ATIVIDADE_SECUNDARIA", Condicao) Then
        For i = 1 To Grid.ListItems.Count
            Campos = "TAS_TCI_IM,TAS_TGA_COD_GRUPO,TAS_TAE_CAE"
            PosGrupo = Trim(InStr(Grid.ListItems(i), " - ") - 1)
            PosAtividade = Trim(InStr(Grid.ListItems(i).SubItems(1), " - ") - 1)
            Valores = Bdados.PreparaValor(Im, Left(Grid.ListItems(i), PosGrupo), Left(Grid.ListItems(i).SubItems(1), PosAtividade))
            If Bdados.InsereDados("TAB_ATIVIDADE_SECUNDARIA", Valores, Campos) Then
                RegistrosAfetados = RegistrosAfetados + 1
            End If
        Next
    End If
    If RegistrosAfetados = Grid.ListItems.Count Then
        Bdados.GravaTrans
        Grava_Atividade_Secundaria = True
    Else
        Bdados.CancelaTrans
        Grava_Atividade_Secundaria = True
    End If
End Function
Public Sub Adiciona_Atividade_Secundaria(Grid As Object, Grupo As Object, atividade As Object)
    Dim RetIm                                 As String
    Dim i                                     As Byte
    Dim Index                                 As Integer
    Dim PosGrupo                              As Integer
    Dim PosAtividade                          As Integer
    Dim Sql                                   As String
    Dim Rs                                    As VSRecordset
    
    If Grupo.ListIndex = -1 Then
        Util.Avisa "Selecione Classificação da Atividade."
        Grupo.SetFocus
        Exit Sub
    End If
    
    If atividade.ListIndex = -1 Then
        Util.Avisa "Selecione Atividade."
        atividade.SetFocus
        Exit Sub
    End If
    
    'Checo se foi inserida uma ATIVIDADE...
    For i = 1 To Grid.ListItems.Count
        PosGrupo = Trim(InStr(Grid.ListItems(i), " - ") - 1)
        PosAtividade = Trim(InStr(Grid.ListItems(i).SubItems(1), " - ") - 1)
        If CStr(Grupo.Coluna(1).Valor) = Left(Grid.ListItems(i), PosGrupo) Then
            If CStr(atividade.Coluna(1).Valor) = Left(Grid.ListItems(i).SubItems(1), PosAtividade) Then
                Avisa "Atividade " & Grid.ListItems(i).SubItems(1) & " já foi inserida para essa classificaçao."
                atividade.SetFocus
                Exit Sub
            End If
        End If
    Next
    
    Index = Grid.ListItems.Count + 1
    Grid.ListItems.Add Index, , Grupo.Coluna(1).Valor & " - " & Grupo.Text
    Grid.ListItems.Item(Index).SubItems(1) = atividade.Coluna(1).Valor & " - " & atividade.Text
    atividade.ListIndex = -1
    atividade.SetFocus
End Sub
Public Sub Seta_Atividade_Secundaria(Grid As Object, Grupo As Object, atividade As Object)
    If Grid.ListItems.Count >= 1 Then
        Dim PosGrupo          As Integer
        Dim PosAtividade      As Integer
        Dim Contador          As Integer
        
        PosGrupo = Trim(InStr(Grid.SelectedItem, " - ") - 1)
        PosAtividade = Trim(InStr(Grid.SelectedItem.SubItems(1), " - ") - 1)
        Grupo.SetarLinha Left(Grid.SelectedItem, PosGrupo)
        Grupo.SetarLinha Left(Grid.SelectedItem.SubItems(1), PosAtividade)
        Grid.ListItems.Remove Grid.SelectedItem.Index
    End If
End Sub

Public Sub Apaga_Atividade_Secundaria(Grid As Object)
    If Grid.ListItems.Count >= 1 Then
        Grid.ListItems.Remove Grid.SelectedItem.Index
    End If
End Sub
Public Sub PreencherAtividadeSecundarias(Grid As Object, Im As String)
    Dim Sql                                   As String
    Dim Rs                                    As VSRecordset
    Dim RsAtividade                           As VSRecordset
    Dim Item                                  As Object
    
    Sql = "Select * from tab_Atividade_Secundaria where TAS_TCI_IM = '" & Im & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        Grid.ListItems.Clear
        Do Until Rs.EOF
            Set Item = Grid.ListItems.Add(Grid.ListItems.Count + 1, , Rs.Fields("TAS_TGA_COD_GRUPO") & " - " & Pega_Descricao_Grupo_Atividade(Rs.Fields("TAS_TGA_COD_GRUPO")))
            Item.SubItems(1) = Rs.Fields("TAS_TAE_CAE") & " - " & Imposto.BuscaNomeCAE(Rs.Fields("TAS_TAE_CAE"))
            Rs.MoveNext
        Loop
    End If
End Sub
Public Function Pega_Descricao_Grupo_Atividade(Codigo As String) As String
    Dim Sql As String
    Sql = "Select TGA_NOME from tab_grupo_atividade wheRe TGA_COD_GRUPO = '" & Codigo & "'"
    If Bdados.AbreTabela(Sql) Then
        Pega_Descricao_Grupo_Atividade = Bdados.Tabela(0)
    Else
        Pega_Descricao_Grupo_Atividade = ""
    End If
End Function

