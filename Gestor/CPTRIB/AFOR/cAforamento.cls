VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAforamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_IC As String
Private m_NumOrdem As Integer
Private m_DataAforamento As String
Private m_Livro As Integer
Private m_Folha As Integer
Private m_IMCedente As String
Private m_IMAdquirinte As String
Private m_TamFrente As Double
Private m_LimFrente As String
Private m_TamDireita As Double
Private m_LimDireita As String
Private m_TamEsquerda As Double
Private m_LimEsquerda As String
Private m_TamFundos As Double
Private m_LimFundos As String
Private m_TestemunhaUm As String
Private m_CPFUm As String
Private m_TestemunhaDois As String
Private m_CPFDois As String
Private m_Usuario As String
Private m_destinacaoTerritorial As String
Private m_EstadoCivilAdquirinte As String
Private m_doc As String
Private m_RG As String
Private m_Profissao As String
Public Observacao As String
Private Const cteTestadaPrincipal As String = "TESTADA PRINCIPAL"

Public Property Get Profissao() As String
    Profissao = m_Profissao
End Property

Public Property Let Profissao(ByVal Value As String)
    m_Profissao = Value
End Property

Public Property Get Rg() As String
    Rg = m_RG
End Property

Public Property Let Rg(ByVal Value As String)
    m_RG = Value
End Property

Public Property Get Doc() As String
    Doc = m_doc
End Property

Public Property Let Doc(ByVal Value As String)
    m_doc = Value
End Property


Public Property Get EstadoCivilAdquirinte() As String
    EstadoCivilAdquirinte = m_EstadoCivilAdquirinte
End Property

Public Property Let EstadoCivilAdquirinte(ByVal Value As String)
    m_EstadoCivilAdquirinte = Value
End Property

Public Property Get destinacaoTerritorial() As String
    destinacaoTerritorial = m_destinacaoTerritorial
End Property

Public Property Let destinacaoTerritorial(ByVal Value As String)
    m_destinacaoTerritorial = Value
End Property

Public Property Get Usuario() As String
    Usuario = m_Usuario
End Property

Public Property Let Usuario(ByVal Value As String)
    m_Usuario = Value
End Property

Public Property Get CPFDois() As String
    CPFDois = m_CPFDois
End Property

Public Property Let CPFDois(ByVal Value As String)
    m_CPFDois = Value
End Property

Public Property Get TestemunhaDois() As String
    TestemunhaDois = m_TestemunhaDois
End Property

Public Property Let TestemunhaDois(ByVal Value As String)
    m_TestemunhaDois = Value
End Property

Public Property Get CPFUm() As String
    CPFUm = m_CPFUm
End Property

Public Property Let CPFUm(ByVal Value As String)
    m_CPFUm = Value
End Property

Public Property Get TestemunhaUm() As String
    TestemunhaUm = m_TestemunhaUm
End Property

Public Property Let TestemunhaUm(ByVal Value As String)
    m_TestemunhaUm = Value
End Property

Public Property Get LimFundos() As String
    LimFundos = m_LimFundos
End Property

Public Property Let LimFundos(ByVal Value As String)
    m_LimFundos = Value
End Property

Public Property Get TamFundos() As Double
    TamFundos = m_TamFundos
End Property

Public Property Let TamFundos(ByVal Value As Double)
    m_TamFundos = Value
End Property

Public Property Get LimEsquerda() As String
    LimEsquerda = m_LimEsquerda
End Property

Public Property Let LimEsquerda(ByVal Value As String)
    m_LimEsquerda = Value
End Property

Public Property Get TamEsquerda() As Double
    TamEsquerda = m_TamEsquerda
End Property

Public Property Let TamEsquerda(ByVal Value As Double)
    m_TamEsquerda = Value
End Property

Public Property Get LimDireita() As String
    LimDireita = m_LimDireita
End Property

Public Property Let LimDireita(ByVal Value As String)
    m_LimDireita = Value
End Property

Public Property Get TamDireita() As Double
    TamDireita = m_TamDireita
End Property

Public Property Let TamDireita(ByVal Value As Double)
    m_TamDireita = Value
End Property

Public Property Get LimFrente() As String
    LimFrente = m_LimFrente
End Property

Public Property Let LimFrente(ByVal Value As String)
    m_LimFrente = Value
End Property

Public Property Get TamFrente() As Double
    TamFrente = m_TamFrente
End Property

Public Property Let TamFrente(ByVal Value As Double)
    m_TamFrente = Value
End Property


Public Property Get IMAdquirinte() As String
    IMAdquirinte = m_IMAdquirinte
End Property

Public Property Let IMAdquirinte(ByVal Value As String)
    m_IMAdquirinte = Value
End Property

Public Property Get IMCedente() As String
    IMCedente = m_IMCedente
End Property

Public Property Let IMCedente(ByVal Value As String)
    m_IMCedente = Value
End Property

Public Property Get Folha() As Integer
    Folha = m_Folha
End Property

Public Property Let Folha(ByVal Value As Integer)
    m_Folha = Value
End Property

Public Property Get Livro() As Integer
    Livro = m_Livro
End Property

Public Property Let Livro(ByVal Value As Integer)
    m_Livro = Value
End Property

Public Property Get DataAforamento() As String
    DataAforamento = m_DataAforamento
End Property

Public Property Let DataAforamento(ByVal Value As String)
    m_DataAforamento = Value
End Property

Public Property Get NumOrdem() As Integer
    NumOrdem = m_NumOrdem
End Property

Public Property Let NumOrdem(ByVal Value As Integer)
    m_NumOrdem = Value
End Property

Public Property Get Ic() As String
    Ic = m_IC
End Property

Public Property Let Ic(ByVal Value As String)
    m_IC = Value
End Property

Public Function Buscar(Ic As String, NumOrdem As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "select TAF_TIM_IC, TAF_NUM_ORDEM, TAF_DATA_AFORAMENTO," & _
        " TAF_LIVRO, TAF_FOLHA, TAF_TCI_IM_CEDENTE," & _
        " TAF_TCI_IM_ADQUIRENTE, TAF_TAM_FRENTE, TAF_LIM_FRENTE," & _
        " TAF_TAM_DIREITA, TAF_LIM_DIREITA, TAF_TAM_ESQUERDA," & _
        " TAF_LIM_ESQUERDA, TAF_TAM_FUNDOS, TAF_LIM_FUNDOS," & _
        " TAF_TESTEMUNHA_UM, TAF_CPF_UM, TAF_TESTEMUNHA_DOIS," & _
        " TAF_CPF_DOIS, TAF_TUS_COD_USUARIO, TAF_DESTINACAO_TERRITORIAL," & _
        " TAF_ESTADO_CIVIL_ADQUIRENTE,TAF_TIPO_DOC,taf_rg,taf_profissao,TAF_OBSERVACAO " & _
        " From dbo.TAB_AFORAMENTO" & _
        " where TAF_TIM_IC = '" & Ic & "' and" & _
        " TAF_NUM_ORDEM = '" & NumOrdem & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        m_IC = "" & Rs!TAF_TIM_IC
        m_NumOrdem = "" & Rs!taf_num_ordem
        m_DataAforamento = "" & Rs!TAF_DATA_AFORAMENTO
        m_Livro = "" & Rs!TAF_LIVRO
        m_Profissao = "" & Rs!taf_Profissao
        m_RG = "" & Rs!taf_Rg
        m_Folha = "" & Rs!TAF_FOLHA
        m_IMCedente = "" & Rs!TAF_TCI_IM_CEDENTE
        m_IMAdquirinte = "" & Rs!TAF_TCI_IM_ADQUIRENTE
        m_TamFrente = "" & Rs!TAF_TAM_FRENTE
        m_LimFrente = "" & Rs!TAF_LIM_FRENTE
        m_TamDireita = "" & Rs!TAF_TAM_DIREITA
        m_LimDireita = "" & Rs!TAF_LIM_DIREITA
        m_TamEsquerda = "" & Rs!TAF_TAM_ESQUERDA
        m_LimEsquerda = "" & Rs!TAF_LIM_ESQUERDA
        m_TamFundos = "" & Rs!TAF_TAM_FUNDOS
        m_LimFundos = "" & Rs!TAF_LIM_FUNDOS
        m_TestemunhaUm = "" & Rs!TAF_TESTEMUNHA_UM
        m_CPFUm = "" & Rs!TAF_CPF_UM
        m_TestemunhaDois = "" & Rs!TAF_TESTEMUNHA_DOIS
        m_CPFDois = "" & Rs!TAF_CPF_DOIS
        m_Usuario = "" & Rs!TAF_TUS_COD_USUARIO
        m_destinacaoTerritorial = "" & Rs!TAF_DESTINACAO_TERRITORIAL
        m_EstadoCivilAdquirinte = "" & Rs!TAF_ESTADO_CIVIL_ADQUIRENTE
        m_doc = "" & Rs!TAF_TIPO_DOC
        Observacao = "" & Rs!TAF_OBSERVACAO
        Buscar = True
    End If
End Function

Public Function Excluir(Ic As String, NumOrdem As String) As Boolean
    Dim condicao As String, Sql As String, Rs As VSRecordset, Ordem As Integer
    Dim Campos As String, Valores As String, RsAfo As VSRecordset
    Dim IMAdquirin As String, Livro As String, Folha As String, Data As String
    condicao = "TAF_TIM_IC = '" & Ic & "' and TAF_NUM_ORDEM = '" & NumOrdem & "'"
    If Bdados.DeletaDados("TAB_AFORAMENTO", condicao) = True Then
        Excluir = True
    Else
        Excluir = False
        Exit Function
    End If
    
    Sql = "select max(taf_num_ordem) as ultimo from tab_aforamento where taf_tim_ic = '" & Ic & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        If "" & Rs!Ultimo <> "" Then
            Ordem = Rs!Ultimo
        End If
    End If
    
    Livro = 0: Folha = 0: Data = 0: IMAdquirin = 0
    If Ordem <> 0 Then
        Sql = "select  TAF_LIVRO," & _
                " TAF_FOLHA," & _
                " TAF_DATA_AFORAMENTO," & _
                " TAF_TCI_IM_ADQUIRENTE" & _
                " From tab_aforamento" & _
                " Where taf_tim_ic = '" & Ic & "' And taf_num_ordem = '" & Ordem & "'"
        If Bdados.AbreTabela(Sql, RsAfo) Then
            Livro = RsAfo!TAF_LIVRO
            Folha = RsAfo!TAF_FOLHA
            Data = RsAfo!TAF_DATA_AFORAMENTO
            IMAdquirin = RsAfo!TAF_TCI_IM_ADQUIRENTE
        End If
    End If
        
    Campos = "TIM_TCI_IM, TIM_AFORADO, TIM_AFORAMENTO_NUMERO, TIM_AFORAMENTO_LIVRO, TIM_AFORAMENTO_FOLHA, TIM_AFORAMENTO_DATA"
    Valores = Bdados.PreparaValor(IMAdquirin, 1, Ordem, Livro, Folha, Data)
    If Bdados.AtualizaDados("TAB_IMOVEL", Valores, Campos, "TIM_IC='" & Ic & "'") Then
        Excluir = True
    End If
End Function

Public Function Salvar(Ic As String, Ordem As String, Data As String, Ficha As String, Livro As String, _
            Folha As String, Total As String, IMCedente As String, IMAdquirente As String, _
            TamFrente As String, LimFrente As String, TamDireita As String, LimDireita As String, _
            TamEsquerda As String, LimEsquerda As String, TamFundos As String, LimFundos As String, _
            Testemunha1 As String, CPF1 As String, Testemunha2 As String, CPF2 As String, Destinacao As Integer, _
            EstadoCivil As Integer, Doc As String, Optional Rg As String, Optional Profissao As String, Optional Obs As String) As Boolean
    Dim Valores As String, Campos As String
    '1. TAF
        Folha = TiraPic(Folha, "/")
        Folha = TiraPic(Folha, "V")
        Folha = TiraPic(Folha, " ")
        Ordem = TiraPic(Ordem, "/")
        Ordem = TiraPic(Ordem, "V")
        Campos = "TAF_TIM_IC, TAF_NUM_ORDEM, TAF_DATA_AFORAMENTO, TAF_LIVRO, TAF_FOLHA,TAF_FICHA,TAF_TCI_IM_CEDENTE, TAF_TCI_IM_ADQUIRENTE, TAF_TAM_FRENTE, TAF_LIM_FRENTE, TAF_TAM_DIREITA, TAF_LIM_DIREITA, TAF_TAM_ESQUERDA, TAF_LIM_ESQUERDA, TAF_TAM_FUNDOS, TAF_LIM_FUNDOS, TAF_TESTEMUNHA_UM, TAF_CPF_UM, TAF_TESTEMUNHA_DOIS, TAF_CPF_DOIS, TAF_TUS_COD_USUARIO,TAF_DESTINACAO_TERRITORIAL,TAF_ESTADO_CIVIL_ADQUIRENTE,TAF_TIPO_DOC,TAF_RG,TAF_PROFISSAO,TAF_OBSERVACAO"
        Valores = Bdados.PreparaValor(Bdados.Converte(Trim$(Ic), tctexto), Ordem, Data, Livro, Folha, Ficha, IMCedente, IMAdquirente, TamFrente, LimFrente, TamDireita, LimDireita, TamEsquerda, LimEsquerda, TamFundos, LimFundos, Testemunha1, Bdados.Converte(CPF1, tctexto), Testemunha2, Bdados.Converte(CPF2, tctexto), Aplicacoes.Usuario, Destinacao, EstadoCivil, Doc, Rg, Profissao, Obs)
        If Bdados.GravaDados("TAB_AFORAMENTO", Valores, Campos, "TAF_NUM_ORDEM=" & Ordem) Then
            'TNC
            Campos = "tnc_ano, tnc_tipo_operacao, tnc_sequencia, tnc_descricao"
            Valores = Bdados.PreparaValor(0, 33, Ordem, "ORDEM AFORAMENTO")
            Bdados.GravaDados "TAB_NUM_CORRELATIVO", Valores, Campos, "tnc_tipo_operacao=33"
            '2. TIM
            Campos = "TIM_TCI_IM, TIM_AFORADO, TIM_AFORAMENTO_NUMERO, TIM_AFORAMENTO_FICHA, TIM_AFORAMENTO_LIVRO, TIM_AFORAMENTO_FOLHA, TIM_AFORAMENTO_DATA"
            Valores = Bdados.PreparaValor(IMAdquirente, 1, Ordem, Ficha, Livro, Folha, Data)
            If Bdados.AtualizaDados("TAB_IMOVEL", Valores, Campos, "TIM_IC='" & Ic & "'") Then
                '3. TAM
                If AtualizarLivro(CInt(Livro), CInt(Folha), CInt(Nvl(Total, 0))) Then
                    Salvar = True
                Else
                    Erro "Erro ao atualizar livro."
                End If
            Else
                Erro "Erro ao atualizar imovel."
            End If
        Else
            Erro "Erro ao gravar aforamento."
        End If
End Function

Public Function AtualizarLivro(Livro As Integer, Atual As Integer, Total As Integer) As Boolean
    Dim Campos As String
    Dim Valores As String
    Dim Situacao As StatusLivroAforamento
    
    '1.
    If Atual >= Total Then
        Situacao = slaLivroFechado
    Else
        Situacao = slaLivroAberto
    End If
    
    '2.
    Campos = "TAM_LIVRO, TAM_FOLHA_ATUAL, TAM_STATUS, TAM_TUS_COD_USUARIO"
    Valores = Bdados.PreparaValor(Livro, Atual, Situacao, Aplicacoes.Usuario)
    AtualizarLivro = Bdados.AtualizaDados("TAB_AFORAMENTO_MANUTENCAO", Valores, Campos, "TAM_LIVRO=" & Livro)
End Function

Public Function ProximoAforamento() As Integer
    Dim Ultimo As Integer
    Ultimo = Bdados.BuscaCodigo("SELECT TNC_SEQUENCIA+1 FROM TAB_NUM_CORRELATIVO WHERE TNC_TIPO_OPERACAO=33")
    If Ultimo = 0 Then
        ProximoAforamento = 1
    Else
        ProximoAforamento = Ultimo
    End If
End Function

Public Function BuscarComponenteImovel(Ic As String, Componente As String) As String
    Dim Sql As String, Rs As VSRecordset
    
    Sql = "SELECT TDI_VALOR_ITEM " & _
            " FROM VIS_DETALHE_IMOVEL " & _
            " WHERE TDI_TIM_IC = '" & Ic & "' AND" & _
                " TCO_DESCRICAO_COMPONENTE='" & Componente & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscarComponenteImovel = Rs!TDI_VALOR_ITEM
    End If
    Bdados.FechaTabela Rs
End Function

Public Function BuscarContribuinte(Im As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    
    If Trim$(Im) = "" Then Exit Function
    Sql = "SELECT TCI_NOME FROM TAB_CONTRIBUINTE WHERE TCI_IM='" & Im & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscarContribuinte = "" & Rs!tci_nome
    End If
    Bdados.FechaTabela Rs
End Function

Public Function BuscarLogradouro(Ic As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    
    If Trim$(Ic) = "" Then Exit Function
    Sql = "SELECT TTL_NOME, " & _
            " TLG_NOME, " & _
            " TIM_NUMERO," & _
            " TBA_NOME" & _
        " FROM TAB_TIPO_LOGR, TAB_LOGRADOURO, TAB_IMOVEL, TAB_BAIRRO" & _
        " WHERE TTL_COD_TIP_LOGR = TLG_TTL_COD_TIP_LOGR AND " & _
            " TLG_COD_LOGRADOURO = TIM_TLG_COD_LOGRADOURO AND " & _
            " TBA_COD_BAIRRO = TLG_TBA_COD_BAIRRO AND " & _
            " TIM_IC ='" & Ic & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscarLogradouro = ("" & Rs!TTL_NOME) & " " & ("" & Rs!tlg_nome) & ", " & ("" & Rs!tim_numero) & " - " & ("" & Rs!TBA_NOME)
    End If
    Bdados.FechaTabela Rs
End Function


Public Sub ProximaFolha(ByRef Livro As String, ByRef Folha As String, ByRef Total As String)
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "SELECT TAM_LIVRO, TAM_FOLHA_ATUAL, TAM_FOLHA_TOTAL" & _
            " FROM TAB_AFORAMENTO_MANUTENCAO" & _
            " WHERE TAM_STATUS=" & slaLivroAberto
    If Bdados.AbreTabela(Sql, Rs) Then
        If Rs!TAM_FOLHA_ATUAL >= Rs!TAM_FOLHA_TOTAL Then
            AtualizarLivro Rs!TAM_LIVRO, Rs!TAM_FOLHA_TOTAL, Rs!TAM_FOLHA_TOTAL
            ProximaFolha Livro, Folha, Total
        Else
            Livro = Rs!TAM_LIVRO
            Folha = Rs!TAM_FOLHA_ATUAL + 1
            Total = Rs!TAM_FOLHA_TOTAL
        End If
    Else
        Livro = ""
        Folha = ""
        If Util.Confirma("O Livro atual já atingiu seu número total de páginas, deseja abrir um novo livro?") Then
        Dim sValores As String
        Dim sCampos As String
        
            Sql = "SELECT TAM_LIVRO, TAM_FOLHA_ATUAL, TAM_FOLHA_TOTAL" & _
                " FROM TAB_AFORAMENTO_MANUTENCAO" & _
                " WHERE TAM_LIVRO = (SELECT MAX(TAM_LIVRO) FROM TAB_AFORAMENTO_MANUTENCAO)"
            If Bdados.AbreTabela(Sql) Then
                sCampos = "TAM_LIVRO, TAM_FOLHA_ATUAL, TAM_FOLHA_TOTAL, TAM_STATUS, TAM_TUS_COD_USUARIO,TAM_DATA_ABERTURA"
                sValores = Bdados.PreparaValor(Bdados.Tabela(0) + 1, 0, Bdados.Tabela(2), 1, Aplicacoes.Usuario, Format(Date, "DD/MM/YYYY"))
                Bdados.InsereDados "TAB_AFORAMENTO_MANUTENCAO", sValores, sCampos
                ProximaFolha Livro, Folha, Total
            End If
        Else
            
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Public Function PreencherGrid(Grid As Object, Optional Ic As String, Optional IMCedente As String, Optional IMAdquirente As String, Optional Livro As String, Optional Ficha As String, Optional DataInicio As String, Optional DataFim As String, Optional Doc As String) As Boolean
     Dim Sql As String, Where As String
    
    Sql = ""
    Where = ""
    '1.
    If Trim$(Ic) <> "" Then Where = Where & " AND TAF_TIM_IC = '" & Ic & "'"
    If Trim$(IMCedente) <> "" Then Where = Where & " AND TAF_TCI_IM_CEDENTE = '" & IMCedente & "'"
    If Trim$(IMAdquirente) <> "" Then Where = Where & " AND TAF_TCI_IM_ADQUIRENTE = '" & IMAdquirente & "'"
    If Trim$(Livro) <> "" Then Where = Where & " AND TAF_LIVRO = " & CInt(Livro)
    If Trim$(Ficha) <> "" Then Where = Where & " AND TAF_FICHA = " & Ficha
    If Trim$(DataInicio) <> "" Then Where = Where & " AND TAF_DATA_AFORAMENTO >= " & Bdados.FormataValorCampo(DataInicio)
    If Trim$(DataFim) <> "" Then Where = Where & " AND TAF_DATA_AFORAMENTO <= " & Bdados.FormataValorCampo(DataFim)
    If Trim$(Doc) <> "" Then Where = Where & " AND TAF_TIPO_DOC = " & Bdados.Converte(Doc, tctexto)
    '2.
    'TIM_AFORAMENTO_FICHA
    Sql = "SELECT TAF_TIM_IC AS IC, " & _
            " TCI_NOME AS Foreiro," & _
            " TAF_NUM_ORDEM AS Ordem, " & _
            " TAF_LIVRO as Livro," & _
            " TAF_FOLHA as Folha,TAF_FICHA AS Ficha," & _
            " TAF_DATA_AFORAMENTO AS Data,TGE_NOME AS DOC " & _
        " FROM TAB_AFORAMENTO, TAB_CONTRIBUINTE,VIS_DOC" & _
        " WHERE TAF_TCI_IM_ADQUIRENTE = TCI_IM AND" & _
            " TGE_CODIGO = TAF_TIPO_DOC"
    Sql = Sql & Where
    Sql = Sql & " ORDER BY TAF_NUM_ORDEM"
    Grid.Preencher Bdados, Sql, (Grid.Width * 25 / 100), (Grid.Width * 40 / 100), (Grid.Width * 10 / 100), (Grid.Width * 10 / 100), (Grid.Width * 10 / 100), (Grid.Width * 10 / 100)
End Function

Public Function BuscaFicha(Ic As String) As String
    Dim Sql As String, Rs As VSRecordset
    Sql = "SELECT TIM_AFORAMENTO_FICHA FROM TAB_IMOVEL WHERE TIM_IC='" & Ic & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaFicha = Rs!TIM_AFORAMENTO_FICHA
    End If
End Function

Public Function BuscaDestinacao(Ic As String, RetDestinacao As String) As Boolean
Dim Sql As String, Rs As VSRecordset
Sql = "SELECT TDI_TCO_COD_COMPONENTE From TAB_DETALHE_IMOVEL WHERE TDI_TGC_COD_GRUPO = 11 AND TDI_TIM_IC = ' " & Ic & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        RetDestinacao = Rs!TDI_TCO_COD_COMPONENTE
        BuscaDestinacao = True
    End If
End Function

Public Function BuscaQuadraLote(Ic As String, retQuadra As String, retLote As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "SELECT TIM_QUADRA,TIM_LOTE FROM TAB_IMOVEL WHERE TIM_IC = '" & Ic & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        retQuadra = "" & Rs!tim_quadra
        retLote = "" & Rs!tim_Lote
    End If
End Function

Public Sub BuscarImovel(Ic As String, RetNumero As String, RetFicha As String, RetLivro As String, RetFolha As String, RetIM As String, RetTamFrente As String, _
                                    RetTamDireita As String, RetLimDireita As String, RetTamEsquerda As String, RetLimEsquerda As String, RetTamFundos As String, _
                                    RetLimFundos As String, RetIMCedente As String, RetTotal As String)
    Dim Sql As String
    Dim Rs As VSRecordset, rsAforamento As VSRecordset
    Dim RsAux As VSRecordset
    
    Dim PosVirgula As Integer
    Sql = "SELECT TIM_AFORAMENTO_NUMERO, " & _
                " TIM_AFORAMENTO_FICHA, " & _
                " TIM_AFORAMENTO_LIVRO, " & _
                " TIM_AFORAMENTO_FOLHA, " & _
                " TIM_TCI_IM" & _
        " FROM VIS_IMOVEL " & _
        " WHERE TIM_IC = '" & Ic & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        Sql = "SELECT TAF_TAM_FRENTE, TAF_LIM_FRENTE," & _
                " TAF_TAM_DIREITA, TAF_LIM_DIREITA," & _
                " TAF_TAM_ESQUERDA, TAF_LIM_ESQUERDA," & _
                " TAF_TAM_FUNDOS, TAF_LIM_FUNDOS" & _
            " FROM TAB_AFORAMENTO" & _
            " WHERE TAF_TIM_IC='" & Ic & "' AND" & _
                " TAF_NUM_ORDEM = (SELECT MAX(TAF_NUM_ORDEM) FROM TAB_AFORAMENTO WHERE TAF_TIM_IC='" & Ic & "')"
        If Bdados.AbreTabela(Sql, rsAforamento) Then
            RetTamFrente = "" & rsAforamento!TAF_TAM_FRENTE
            RetTamDireita = "" & rsAforamento!TAF_TAM_DIREITA: RetLimDireita = "" & rsAforamento!TAF_LIM_DIREITA
            RetTamEsquerda = "" & rsAforamento!TAF_TAM_ESQUERDA: RetLimEsquerda = "" & rsAforamento!TAF_LIM_ESQUERDA
            RetTamFundos = "" & rsAforamento!TAF_TAM_FUNDOS: RetLimFundos = "" & rsAforamento!TAF_LIM_FUNDOS
        Else
            RetTamFrente = BuscarComponenteImovel(Ic, cteTestadaPrincipal)
            RetTamDireita = "": RetLimDireita = ""
            RetTamEsquerda = "": RetLimEsquerda = ""
            RetTamFundos = "": RetLimFundos = ""
        End If
        Bdados.FechaTabela rsAforamento
        
        '2.Aforamento
        If "" & Rs!TIM_AFORAMENTO_FICHA = "" Then
            Sql = "select max(TIM_AFORAMENTO_FICHA) as num from tab_imovel"
            If Bdados.AbreTabela(Sql, RsAux) Then
                RetFicha = CDbl(Nvl(RsAux!num, 0)) + 1
            End If
            Bdados.FechaTabela RsAux
        Else
            RetFicha = Rs!TIM_AFORAMENTO_FICHA
            RetLivro = "" & Rs!TIM_AFORAMENTO_LIVRO
            'RetFolha = "" & rs!TIM_AFORAMENTO_FOLHA
            If Trim(RetLivro) <> "" Then RetTotal = "" & Bdados.BuscaCodigo("SELECT TAM_FOLHA_TOTAL FROM TAB_AFORAMENTO_MANUTENCAO WHERE TAM_LIVRO=" & RetLivro)
        End If
        ProximaFolha RetLivro, RetFolha, RetTotal
        
        '3. Cedente
        RetIMCedente = "" & Rs!tim_tci_im
        'txtcedete = Aforamento.BuscarContribuinte(txtIMCedente)
    End If
    Bdados.FechaTabela Rs
End Sub


Public Sub PreencherCboDestinacao(Cbo As Object)
    Dim Sql As String
    Sql = "SELECT TCO_DESCRICAO_COMPONENTE,TCO_COD_COMPONENTE From TAB_COMPONENTE_AVANCADO WHERE TCO_GRUPO = 29"
    Cbo.Preencher Bdados, Sql
End Sub

Public Function ConfirmaUltimo(Ic As String, NumOredem As Integer) As Boolean
    Dim Sql As String, Rs As VSRecordset
    Dim Ultimo As Integer
    Sql = "select max(taf_num_ordem) as ultimo from tab_aforamento where taf_tim_ic = '" & Ic & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        Ultimo = Rs!Ultimo
    Else
        Ultimo = 0
    End If
    If NumOrdem < Ultimo Then ConfirmaUltimo = False Else ConfirmaUltimo = True
End Function
