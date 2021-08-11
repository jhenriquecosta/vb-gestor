VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImovel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_IC As String
Private m_Unidade As Double
Private m_IM As String
Private m_CodLogradouro As String
Private m_Numero As String
Private m_Complemento As String
Private m_CEP As String
Private m_AnoAquis As Integer
Private m_Valor As Double
Private m_TipoImovel As Integer
Private m_Ocupante As String
Private m_CPFCGCOcupante As String
Private m_ICCondominio As String
Private m_CodBairro As Integer
Private m_CodMensagem As Integer
Private m_Zona As Integer
Private m_ValorTerreno As Double
Private m_ValorEdificado As Double
Private m_Loteamento As String
Private m_Secao As String
Private m_Quadra As String
Private m_Lote As String
Private m_CodSitCad As Integer
Private m_Aforado As String
Private m_Obs As String
Private m_AforamentoRegistro As String
Private m_DataRegistro As String
Private m_AforamentoNumero As String
Private m_AforamentoLivro As String
Private m_AforamentoFolha As String
Private m_AforamentoFicha As String
Private m_AforamentoData As String
Private m_MotivoAlteracao As String
Private m_ValorTerrenoMercado As Double
Private m_ValorEdificacaoMercado As Double
Private m_SituacaoLote As Integer
Private m_SubUnidade As String
Private m_Usuario As String
Private m_DataCadastro As String
Private m_TipoBoletim As String
Private m_ICAnterior As String

Private m_pac As Boolean

Public Property Get ICAnterior() As String
    ICAnterior = m_ICAnterior
End Property

Public Property Let ICAnterior(ByVal Value As String)
    m_ICAnterior = Value
End Property

Public Property Get TipoBoletim() As String
    TipoBoletim = m_TipoBoletim
End Property

Public Property Let TipoBoletim(ByVal Value As String)
    m_TipoBoletim = Value
End Property

Public Property Get Pac() As Boolean
    Pac = m_pac
End Property

Public Property Let Pac(ByVal Value As Boolean)
    m_pac = Value
End Property

Public Property Get DataCadastro() As String
    DataCadastro = m_DataCadastro
End Property

Public Property Let DataCadastro(ByVal Value As String)
    m_DataCadastro = Value
End Property

Public Property Get Usuario() As String
    Usuario = m_Usuario
End Property

Public Property Let Usuario(ByVal Value As String)
    m_Usuario = Value
End Property

Public Property Get SubUnidade() As String
    SubUnidade = m_SubUnidade
End Property

Public Property Let SubUnidade(ByVal Value As String)
    m_SubUnidade = Value
End Property

Public Property Get SituacaoLote() As Integer
    SituacaoLote = m_SituacaoLote
End Property

Public Property Let SituacaoLote(ByVal Value As Integer)
    m_SituacaoLote = Value
End Property

Public Property Get ValorEdificacaoMercado() As Double
    ValorEdificacaoMercado = m_ValorEdificacaoMercado
End Property

Public Property Let ValorEdificacaoMercado(ByVal Value As Double)
    m_ValorEdificacaoMercado = Value
End Property

Public Property Get ValorTerrenoMercado() As Double
    ValorTerrenoMercado = m_ValorTerrenoMercado
End Property

Public Property Let ValorTerrenoMercado(ByVal Value As Double)
    m_ValorTerrenoMercado = Value
End Property

Public Property Get MotivoAlteracao() As String
    MotivoAlteracao = m_MotivoAlteracao
End Property

Public Property Let MotivoAlteracao(ByVal Value As String)
    m_MotivoAlteracao = Value
End Property

Public Property Get AforamentoData() As String
    AforamentoData = m_AforamentoData
End Property

Public Property Let AforamentoData(ByVal Value As String)
    m_AforamentoData = Value
End Property

Public Property Get AforamentoFicha() As String
    AforamentoFicha = m_AforamentoFicha
End Property

Public Property Let AforamentoFicha(ByVal Value As String)
    m_AforamentoFicha = Value
End Property

Public Property Get AforamentoFolha() As String
    AforamentoFolha = m_AforamentoFolha
End Property

Public Property Let AforamentoFolha(ByVal Value As String)
    m_AforamentoFolha = Value
End Property

Public Property Get AforamentoLivro() As String
    AforamentoLivro = m_AforamentoLivro
End Property

Public Property Let AforamentoLivro(ByVal Value As String)
    m_AforamentoLivro = Value
End Property

Public Property Get AforamentoNumero() As String
    AforamentoNumero = m_AforamentoNumero
End Property

Public Property Let AforamentoNumero(ByVal Value As String)
    m_AforamentoNumero = Value
End Property

Public Property Get DataRegistro() As String
    DataRegistro = m_DataRegistro
End Property

Public Property Let DataRegistro(ByVal Value As String)
    m_DataRegistro = Value
End Property

Public Property Get AforamentoRegistro() As String
    AforamentoRegistro = m_AforamentoRegistro
End Property

Public Property Let AforamentoRegistro(ByVal Value As String)
    m_AforamentoRegistro = Value
End Property

Public Property Get Obs() As String
    Obs = m_Obs
End Property

Public Property Let Obs(ByVal Value As String)
    m_Obs = Value
End Property

Public Property Get Aforado() As String
    Aforado = m_Aforado
End Property

Public Property Let Aforado(ByVal Value As String)
    m_Aforado = Value
End Property

Public Property Get CodSitCad() As Integer
    CodSitCad = m_CodSitCad
End Property

Public Property Let CodSitCad(ByVal Value As Integer)
    m_CodSitCad = Value
End Property

Public Property Get Lote() As String
    Lote = m_Lote
End Property

Public Property Let Lote(ByVal Value As String)
    m_Lote = Value
End Property

Public Property Get Quadra() As String
    Quadra = m_Quadra
End Property

Public Property Let Quadra(ByVal Value As String)
    m_Quadra = Value
End Property

Public Property Get Secao() As String
    Secao = m_Secao
End Property

Public Property Let Secao(ByVal Value As String)
    m_Secao = Value
End Property

Public Property Get Loteamento() As String
    Loteamento = m_Loteamento
End Property

Public Property Let Loteamento(ByVal Value As String)
    m_Loteamento = Value
End Property

Public Property Get ValorEdificado() As Double
    ValorEdificado = m_ValorEdificado
End Property

Public Property Let ValorEdificado(ByVal Value As Double)
    m_ValorEdificado = Value
End Property

Public Property Get ValorTerreno() As Double
    ValorTerreno = m_ValorTerreno
End Property

Public Property Let ValorTerreno(ByVal Value As Double)
    m_ValorTerreno = Value
End Property

Public Property Get Zona() As Integer
    Zona = m_Zona
End Property

Public Property Let Zona(ByVal Value As Integer)
    m_Zona = Value
End Property

Public Property Get CodMensagem() As Integer
    CodMensagem = m_CodMensagem
End Property

Public Property Let CodMensagem(ByVal Value As Integer)
    m_CodMensagem = Value
End Property

Public Property Get CodBairro() As Integer
    CodBairro = m_CodBairro
End Property

Public Property Let CodBairro(ByVal Value As Integer)
    m_CodBairro = Value
End Property

Public Property Get ICCondominio() As String
    ICCondominio = m_ICCondominio
End Property

Public Property Let ICCondominio(ByVal Value As String)
    m_ICCondominio = Value
End Property

Public Property Get CPFCGCOcupante() As String
    CPFCGCOcupante = m_CPFCGCOcupante
End Property

Public Property Let CPFCGCOcupante(ByVal Value As String)
    m_CPFCGCOcupante = Value
End Property

Public Property Get Ocupante() As String
    Ocupante = m_Ocupante
End Property

Public Property Let Ocupante(ByVal Value As String)
    m_Ocupante = Value
End Property

Public Property Get TipoImovel() As Integer
    TipoImovel = m_TipoImovel
End Property

Public Property Let TipoImovel(ByVal Value As Integer)
    m_TipoImovel = Value
End Property

Public Property Get Valor() As Double
    Valor = m_Valor
End Property

Public Property Let Valor(ByVal Value As Double)
    m_Valor = Value
End Property

Public Property Get AnoAquis() As Integer
    AnoAquis = m_AnoAquis
End Property

Public Property Let AnoAquis(ByVal Value As Integer)
    m_AnoAquis = Value
End Property

Public Property Get CEP() As String
    CEP = m_CEP
End Property

Public Property Let CEP(ByVal Value As String)
    m_CEP = Value
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

Public Property Get CodLogradouro() As String
    CodLogradouro = m_CodLogradouro
End Property

Public Property Let CodLogradouro(ByVal Value As String)
    m_CodLogradouro = Value
End Property

Public Property Get Im() As String
    Im = m_IM
End Property

Public Property Let Im(ByVal Value As String)
    m_IM = Value
End Property

Public Property Get Unidade() As Double
    Unidade = m_Unidade
End Property

Public Property Let Unidade(ByVal Value As Double)
    m_Unidade = Value
End Property

Public Property Get Ic() As String
    Ic = m_IC
End Property

Public Property Let Ic(ByVal Value As String)
    m_IC = Value
End Property
Public Function Buscar(Ic As String) As Boolean
    Dim Sql As String, rs As VSRecordset
    Sql = "select * from tab_imovel where tim_ic = '" & Ic & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        m_IC = "" & rs!TIM_IC
        m_Unidade = "" & rs!tim_unidade
        m_IM = "" & rs!tim_tci_im
        m_ICAnterior = "" & rs!tim_ic_anterior
        m_CodLogradouro = "" & rs!tim_tlg_cod_logradouro
        m_Numero = "" & rs!tim_numero
        m_Complemento = "" & rs!tim_complemento
        m_CEP = "" & rs!tim_cep
        m_AnoAquis = "" & rs!tim_ano_aquis
        m_Valor = "" & rs!tim_valor
        m_TipoImovel = "" & rs!tim_tipo_imovel
        m_Ocupante = "" & rs!tim_ocupante
        m_CPFCGCOcupante = "" & rs!tim_cgc_cpf_ocupante
        m_ICCondominio = "" & rs!tim_IC_CONDOMINIO
        m_CodBairro = "" & rs!tim_TBA_COD_BAIRRO
        m_CodMensagem = "" & rs!tim_COD_MENSAGEM
        m_Zona = "" & rs!tim_ZONA
        m_ValorTerreno = "" & rs!tim_VALOR_TERRENO
        m_ValorEdificado = "" & rs!tim_VALOR_EDIFIC
        m_Loteamento = "" & rs!tim_loteamento
        m_Secao = "" & rs!tim_secao
        m_Quadra = "" & rs!tim_QUADRA
        m_Lote = "" & rs!tim_lote
        m_CodSitCad = "" & rs!tim_tsc_cod_sit_cad
        m_Aforado = "" & rs!tim_aforado
        m_Obs = "" & rs!tim_obs
        m_AforamentoRegistro = "" & rs!tim_AFORAMENTO_REGISTRO
        m_DataRegistro = "" & rs!tim_DATA_REGISTRO
        m_AforamentoNumero = "" & rs!tim_AFORAMENTO_NUMERO
        m_AforamentoFicha = "" & rs!tim_AFORAMENTO_FICHA
        m_AforamentoLivro = "" & rs!tim_AFORAMENTO_LIVRO
        m_AforamentoFolha = "" & rs!tim_AFORAMENTO_FOLHA
        m_AforamentoData = "" & rs!tim_AFORAMENTO_DATA
        m_MotivoAlteracao = "" & rs!tim_MOTIVO_ALTERACAO
        m_ValorTerrenoMercado = "" & rs!tim_VALOR_TERRENO_MERCADO
        m_ValorEdificacaoMercado = "" & rs!tim_VALOR_EDIFICACAO_MERCADO
        m_SituacaoLote = "" & rs!TIM_SITUACAO_LOTE
        m_SubUnidade = "" & rs!tim_SUB_UNIDADE
        m_Usuario = "" & rs!tim_TUS_COD_USUARIO
        m_DataCadastro = "" & rs!tim_DATA_CADASTRO
        m_TipoBoletim = "" & rs!tim_TIPO_BOLETIM
        m_pac = rs!tim_Pac
        Buscar = True
    End If
End Function


