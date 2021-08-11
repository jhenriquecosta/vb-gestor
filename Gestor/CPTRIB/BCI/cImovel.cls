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
Private m_IcAnterior As String
Private m_CodLogradouro As String
Private m_Numero As String
Private m_Complemento As String
Private m_CEP As String
Private m_AnoAquisicao As Integer
Private m_Valor As Double
Private m_TipoImovel As Integer
Private m_Ocupante As String
Private m_CgcCpfOcupante As String
Private m_IcCondominio As String
Private m_CodBairro As Integer
Private m_CodMensagem As Integer
Private m_Zona As Integer
Private m_ValorTerreno As Double
Private m_ValorEdific As Double
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
Private m_AforamentoFicha As String
Private m_AforamentoLivro As String
Private m_AforamentoFolha As String
Private m_AforamentoData As String
Private m_MotivoAlteracao As String
Private m_ValorTerrenoMercado As String
Private m_ValorEdificacaoMercado As String
Private m_SituacaoLote As Integer
Private m_SubUnidade As String
Private m_DataCadastro As String
Private m_TipoBoletim As String
Private m_Usuario As String

Public Property Get Usuario() As String
    Usuario = m_Usuario
End Property

Public Property Let Usuario(ByVal Value As String)
    m_Usuario = Value
End Property

Public Property Get TipoBoletim() As String
    TipoBoletim = m_TipoBoletim
End Property

Public Property Let TipoBoletim(ByVal Value As String)
    m_TipoBoletim = Value
End Property

Public Property Get DataCadastro() As String
    DataCadastro = m_DataCadastro
End Property

Public Property Let DataCadastro(ByVal Value As String)
    m_DataCadastro = Value
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

Public Property Get ValorEdificacaoMercado() As String
    ValorEdificacaoMercado = m_ValorEdificacaoMercado
End Property

Public Property Let ValorEdificacaoMercado(ByVal Value As String)
    m_ValorEdificacaoMercado = Value
End Property

Public Property Get ValorTerrenoMercado() As String
    ValorTerrenoMercado = m_ValorTerrenoMercado
End Property

Public Property Let ValorTerrenoMercado(ByVal Value As String)
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

Public Property Get AforamentoFicha() As String
    AforamentoFicha = m_AforamentoFicha
End Property

Public Property Let AforamentoFicha(ByVal Value As String)
    m_AforamentoFicha = Value
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

Public Property Get ValorEdific() As Double
    ValorEdific = m_ValorEdific
End Property

Public Property Let ValorEdific(ByVal Value As Double)
    m_ValorEdific = Value
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

Public Property Get IcCondominio() As String
    IcCondominio = m_IcCondominio
End Property

Public Property Let IcCondominio(ByVal Value As String)
    m_IcCondominio = Value
End Property

Public Property Get CgcCpfOcupante() As String
    CgcCpfOcupante = m_CgcCpfOcupante
End Property

Public Property Let CgcCpfOcupante(ByVal Value As String)
    m_CgcCpfOcupante = Value
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

Public Property Get AnoAquisicao() As Integer
    AnoAquisicao = m_AnoAquisicao
End Property

Public Property Let AnoAquisicao(ByVal Value As Integer)
    m_AnoAquisicao = Value
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

Public Property Get IcAnterior() As String
    IcAnterior = m_IcAnterior
End Property

Public Property Let IcAnterior(ByVal Value As String)
    m_IcAnterior = Value
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
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "select * from tab_imovel where tim_ic = '" & Ic & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        m_IC = "" & Rs!tim_ic
        m_Unidade = "" & Rs!tim_unidade
        m_IM = "" & Rs!tim_tci_im
        m_IcAnterior = "" & Rs!tim_ic_anterior
        m_CodLogradouro = "" & Rs!tim_tlg_cod_logradouro
        m_Numero = "" & Rs!tim_numero
        m_Complemento = "" & Rs!tim_complemento
        m_CEP = "" & Rs!tim_cep
        m_AnoAquisicao = "" & Rs!tim_ano_aquis
        m_Valor = "" & Rs!tim_valor
        m_TipoImovel = "" & Rs!tim_tipo_imovel
        m_Ocupante = "" & Rs!tim_ocupante
        m_CgcCpfOcupante = "" & Rs!tim_cgc_cpf_ocupante
        m_IcCondominio = "" & Rs!tim_IC_CONDOMINIO
        m_CodBairro = "" & Rs!tim_TBA_COD_BAIRRO
        m_CodMensagem = "" & Rs!tim_COD_MENSAGEM
        m_Zona = "" & Rs!tim_ZONA
        m_ValorTerreno = "" & Rs!tim_VALOR_TERRENO
        m_ValorEdific = "" & Rs!tim_VALOR_EDIFIC
        m_Loteamento = "" & Rs!Tim_loteamento
        m_Secao = "" & Rs!tim_secao
        m_Quadra = "" & Rs!tim_quadra
        m_Lote = "" & Rs!tim_Lote
        m_CodSitCad = "" & Rs!tim_tsc_cod_sit_cad
        m_Aforado = "" & Rs!tim_aforado
        m_Obs = "" & Rs!tim_obs
        m_AforamentoRegistro = "" & Rs!tim_AFORAMENTO_REGISTRO
        m_DataRegistro = "" & Rs!tim_DATA_REGISTRO
        m_AforamentoNumero = "" & Rs!tim_AFORAMENTO_NUMERO
        m_AforamentoFicha = "" & Rs!tim_AFORAMENTO_FICHA
        m_AforamentoLivro = "" & Rs!tim_AFORAMENTO_LIVRO
        m_AforamentoFolha = "" & Rs!tim_AFORAMENTO_FOLHA
        m_AforamentoData = "" & Rs!tim_AFORAMENTO_DATA
        m_MotivoAlteracao = "" & Rs!tim_MOTIVO_ALTERACAO
        m_ValorTerrenoMercado = "" & Rs!tim_VALOR_TERRENO_MERCADO
        m_ValorEdificacaoMercado = "" & Rs!tim_VALOR_EDIFICACAO_MERCADO
        m_SituacaoLote = "" & Rs!TIM_SITUACAO_LOTE
        m_SubUnidade = "" & Rs!tim_SUB_UNIDADE
        m_Usuario = "" & Rs!tim_TUS_COD_USUARIO
        m_DataCadastro = "" & Rs!tim_DATA_CADASTRO
        m_TipoBoletim = "" & Rs!tim_TIPO_BOLETIM
        Buscar = True
    End If
End Function

Public Function TemEmpresaNoImovel(Inscricao As String) As Boolean
    Dim rsIm As VSRecordset
    Dim Sql As String
    
    Sql = "Select tci_tim_ic from Tab_Contribuinte where tci_tim_ic ='" & Inscricao & "'  and tci_tsc_cod_sit_cad =1"
    If Bdados.AbreTabela(Sql, rsIm) Then
        TemEmpresaNoImovel = True
    Else
        TemEmpresaNoImovel = False
    End If
    Bdados.FechaTabela rsIm
End Function

Public Function PreencherGrid(Grid As Object, Optional Ic As String, Optional Im As String, Optional Tipologr As String, Optional Logr As String, _
                                            Optional Bairro As String, Optional Contrib As String, Optional Loteamento As String, Optional Quadra As String, _
                                            Optional Aforado As String, Optional AnoAquis As String, Optional Valor As String) As Boolean
    Dim Sql As String
    Dim Condicao As String
    Sql = "SELECT tim_ic  as IC, " & _
                        " tim_tci_im as IM," & _
                        " tci_nome as Contribuinte," & _
                        " TTL_NOME as Logr," & _
                        " tlg_nome as Nome," & _
                        " tim_numero as [Nº], " & _
                        " TBA_NOME as Bairro," & _
                        " tim_valor as [Valor(R$)] " & _
                        " FROM Vis_Imovel "
            If Trim(Ic) <> "" Then
                Condicao = " and tim_ic ='" & Ic & "'"
            End If
            If Trim(Im) <> "" Then
                Condicao = Condicao & " and tim_tci_im = '" & Im & "'"
            End If
            If Trim(Tipologr) <> "" Then
                Condicao = Condicao & " and TTL_NOME = '" & Tipologr & "'"
            End If
            If Trim(Logr) <> "" Then
                Condicao = Condicao & " and tlg_nome = '" & Logr & "'"
            End If
            If Trim(Bairro) <> "" Then
                Condicao = Condicao & " and TBA_NOME = '" & Bairro & "'"
            End If
            If Trim(Contrib) <> "" Then
                Condicao = Condicao & " and (tci_nome like '" & Contrib & "%' or tci_nome like '%" & Contrib & "%')"
            End If
            If Trim(Loteamento) <> "" Then
                Condicao = Condicao & " and mid(tim_ic,3,2) = '" & Loteamento & "'"
            End If
            If Trim(Quadra) <> "" Then
                Condicao = Condicao & " and mid(tim_ic,5,4) = '" & Quadra & "'"
            End If
            If Trim(Aforado) <> "" Then
                Condicao = Condicao & " and tim_aforado = '" & Left(Aforado, 1) & "'"
            End If
            If Trim(AnoAquis) <> "" Then
                Condicao = Condicao & " and tim_ano_aquis = '" & AnoAquis & "'"
            End If
            If Trim(Valor) <> "" Then
                Condicao = Condicao & " and tim_valor = " & Valor
            End If
            If Condicao <> "" Then
                Sql = Sql & " where " & Right(Condicao, Len(Condicao) - 4)
            End If
            If Grid.Preencher(Bdados, Sql) Then PreencherGrid = True
End Function

Public Function BuscarVisImovel(Ic As String, Optional Tipologr As String, Optional Logr As String, Optional Numero As String, Optional Complemento As String, _
                                            Optional Bairro As String, Optional Loteamento As String, Optional Lote As String, Optional Quadra As String, Optional Secao As String, _
                                            Optional Im As String, Optional AnoAq As String, Optional Aforado As String, Optional Valor As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "Select * from Vis_Imovel Where tim_Ic = '" & Ic & "'"
    Sql = Sql & " AND TBA_TMU_COD_MUNICIPIO = " & Aplicacoes.Codigo_Municipio & " AND tlg_tmu_cod_municipio = " & Aplicacoes.Codigo_Municipio
    If Bdados.AbreTabela(Sql, Rs) = True Then
        'daddos do imovel
        Tipologr = "" & Rs!TTL_NOME
        Logr = "" & Rs!tlg_nome
        Numero = "" & Rs!tim_numero
        Complemento = "" & Rs!tim_complemento
        Bairro = "" & Rs!TBA_NOME
        Loteamento = "" & Rs!Tim_loteamento
        Lote = "" & Rs!tim_Lote
        Quadra = "" & Rs!tim_quadra
        Secao = "" & Rs!tim_secao
        'dados do contribuinte
        Im = "" & Rs!tim_tci_im
        AnoAq = "" & Rs!tim_ano_aquis
        Aforado = "" & Rs!tim_aforado
        Valor = Rs!tim_valor
        BuscarVisImovel = True
    End If
End Function

Public Function Excluir(Ic As String) As Boolean
    If Bdados.DeletaDados("TAB_IMOVEL", "tim_ic = '" & Ic & "'") Then Excluir = True
End Function

Public Function Gravar(Ic As String) As Boolean
    Dim Campos As String, Valores As String
    Campos = "tim_ic,tim_unidade, tim_tci_im , tim_tlg_cod_logradouro, " _
                & "tim_numero,  tim_complemento, tim_loteamento, tim_secao,tim_quadra," _
                & "tim_lote, tim_cep, tim_ano_aquis , tim_aforado, tim_valor,tim_tsc_cod_sit_cad,tim_ic_anterior,tim_tba_cod_bairro"
    Valores = Bdados.PreparaValor(Ic, 0, m_IM, _
                m_CodLogradouro, m_Numero, m_Complemento, _
                m_Loteamento, m_Secao, m_Quadra, m_Lote, _
                m_CEP, m_AnoAquisicao, m_Aforado, m_Valor, 1, m_IcAnterior, m_CodBairro)
    If Bdados.GravaDados("TAB_IMOVEL", Valores, Campos, "TIM_IC = '" & Ic & "'") Then
        Gravar = True
    End If
End Function


Public Function preenchercomponente(Combo As Object, Tag As Integer) As Boolean
    Dim Sql As String
    Sql = "Select tco_descricao_componente, tco_cod_componente, tco_grupo " & _
             " From Tab_Componente " & _
             " Where tco_grupo = " & Tag & " and tco_tmu_cod_municipio =" & Aplicacoes.Codigo_Municipio & " order by tco_cod_componente asc"
    Combo.Preencher Bdados, Sql
End Function

Public Function BuscarComponente(Ic As String, Formulario As Object) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Controle As Control
    '1
    Sql = "SELECT * FROM TAB_DETALHE_IMOVEL WHERE tdi_tim_ic = '" & Ic & "' AND tdi_tgc_cod_grupo < 100"
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            For Each Controle In Formulario.Controls
                If IsNumeric(Controle.Tag) Then
                    If Controle.Tag < 100 Then
                        If CInt(Controle.Tag) = CInt(Rs!tdi_tgc_cod_grupo) Then
                            Controle.SetarLinha Rs!tdi_tco_cod_componente, 1
                            Exit For
                        End If
                    End If
                End If
            Next
            Rs.MoveNext
        Loop
    End If
    Sql = "SELECT * FROM TAB_DETALHE_IMOVEL WHERE tdi_tim_ic = '" & Ic & "' AND tdi_tgc_cod_grupo >= 100"
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Do While Not Rs.EOF
            For Each Controle In Formulario.Controls
                If IsNumeric(Controle.Tag) Then
                    If Controle.Tag >= 100 Then
                        If CInt(Rs!tdi_tco_cod_componente) = CInt(Controle.Tag) Or CInt(Rs!tdi_tco_cod_componente) = CInt(Controle.Tag) + 100 Then
                            Controle.Text = Rs!TDI_VALOR_ITEM
                            Exit For
                        End If
                    End If
                End If
            Next
            Rs.MoveNext
        Loop
    End If
End Function

