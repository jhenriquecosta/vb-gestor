VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGraficaAidf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Im As String
Private m_DataInicio As String
Private m_Validade As String
Private m_Status As Integer
Private m_Situacao As Integer
Private m_NumCredenciamento As Double
Private m_CodUsuario As String
Private m_DataDescredenciamento As String

Public Enum TipoAcaoGrafica
    eagCredenciar = 0
    eagDescredenciar = 1
End Enum

Public Enum TipoNumero
    Credenciamento = 22
    AIDF = 23
    AIDG = 73
End Enum

Public Property Get DataDescredenciamento() As String
    DataDescredenciamento = m_DataDescredenciamento
End Property

Public Property Let DataDescredenciamento(ByVal Value As String)
    m_DataDescredenciamento = Value
End Property

Public Property Get CodUsuario() As String
    CodUsuario = m_CodUsuario
End Property

Public Property Let CodUsuario(ByVal Value As String)
    m_CodUsuario = Value
End Property

Public Property Get NumCredenciamento() As Double
    NumCredenciamento = m_NumCredenciamento
End Property

Public Property Let NumCredenciamento(ByVal Value As Double)
    m_NumCredenciamento = Value
End Property

Public Property Get Situacao() As Integer
    Situacao = m_Situacao
End Property

Public Property Let Situacao(ByVal Value As Integer)
    m_Situacao = Value
End Property

Public Property Get Status() As Integer
    Status = m_Status
End Property

Public Property Let Status(ByVal Value As Integer)
    m_Status = Value
End Property

Public Property Get Validade() As String
    Validade = m_Validade
End Property

Public Property Let Validade(ByVal Value As String)
    m_Validade = Value
End Property

Public Property Get DataInicio() As String
    DataInicio = m_DataInicio
End Property

Public Property Let DataInicio(ByVal Value As String)
    m_DataInicio = Value
End Property

Public Property Get Im() As String
    Im = m_Im
End Property

Public Property Let Im(ByVal Value As String)
    m_Im = Value
End Property

Public Function Buscar(Optional NumCredenciamento As String, Optional Im As String) As Boolean
    Dim Sql As String, sqlaux As String
    Dim rs As VSRecordset
    Sql = "Select * from tab_Grafica_Credenciada "
    If NumCredenciamento <> "" Then
        sqlaux = sqlaux & " and tgc_num_credenciamento = '" & NumCredenciamento & "'"
    End If
    If Im <> "" Then
        sqlaux = sqlaux & " and tgc_tci_im = '" & Im & "' and tgc_status=0"
    End If
    If sqlaux <> "" Then Sql = Sql & " where " & Right(sqlaux, Len(sqlaux) - 4)
    If Bdados.AbreTabela(Sql, rs) Then
        m_Im = "" & rs!tgc_tci_im
        m_DataInicio = "" & rs!tgc_data_inicio
        m_Validade = "" & rs!tgc_validade
        m_Status = "" & rs!tgc_status
        m_Situacao = "" & rs!tgc_situacao
        m_NumCredenciamento = "" & rs!tgc_num_credenciamento
        m_CodUsuario = "" & rs!tgc_tus_cod_usuario
        m_DataDescredenciamento = "" & rs!tgc_DATA_DESCREDENCIAMENTO
        Buscar = True
    End If
End Function

Public Sub PreencherCboTipo(Combo As Object)
    Dim Sql As String
    Sql = "SELECT TGE_NOME, TGE_CODIGO FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =702 ORDER BY TGE_CODIGO ASC"
    Combo.Preencher Bdados, Sql
End Sub

Public Function PreencherGrid(Grid As Object, Optional NumCredenciamento As String, Optional Nome As String, Optional Im As String, Optional CGC As String) As Boolean
    Dim rs As VSRecordset
    Dim Sql As String
    Sql = "SELECT tgc_num_credenciamento as Número, " & _
            " tgc_tci_im AS IM , tci_nome as Nome, tci_cgc_cpf AS CGC_CPF" & _
            " From Tab_Contribuinte, tab_grafica_credenciada" & _
            " Where tgc_tci_im = tci_im" & _
            " and tgc_status = 0"
        If Trim(Im) <> "" Then
            Sql = Sql & " and tgc_tci_im = '" & Im & "'"
        End If
        If Trim(CGC) <> "" Then
            Sql = Sql & " and tci_cgc_cpf = '" & CGC & "'"
        End If
        If Trim(Nome) <> "" Then
            Sql = Sql & " and tci_nome like '%" & Nome & "%'"
        End If
        If Trim(NumCredenciamento) <> "" Then
            Sql = Sql & " and tgc_num_credenciamento='" & NumCredenciamento & "'"
        End If
        Sql = Sql & " order by tgc_num_credenciamento"
    If Grid.Preencher(Bdados, Sql) Then PreencherGrid = True
End Function

Public Function VerificaCredenciamentoGrafica(InscricaoMunicipal As String, Acao As TipoAcaoGrafica) As String
    Dim Sql As String
    Dim rs As VSRecordset

    Sql = " select tgc_validade from Tab_Grafica_Credenciada where tgc_tci_im='" & InscricaoMunicipal & _
        "' and tgc_status=0 and tgc_situacao=" & Acao
    If Bdados.AbreTabela(Sql, rs) Then
        VerificaCredenciamentoGrafica = "" & rs!tgc_validade
    Else
        VerificaCredenciamentoGrafica = ""
    End If
End Function

Public Function Descredenciar(InscricaoMunicipal As String) As Boolean
    Dim campos As String
    Dim valores As String
    Dim condicao As String
    campos = "tgc_situacao, tgc_status, tgc_tus_cod_usuario, TGC_DATA_DESCREDENCIAMENTO"
    valores = Bdados.PreparaValor(1, 0, AplicacoesVTFuncoes.Usuario, Bdados.Converte(Date, TCDataHora))
    condicao = "tgc_tci_im='" & InscricaoMunicipal & "' and tgc_status=0 and tgc_situacao = 0"
    Descredenciar = Bdados.GravaDados("Tab_Grafica_Credenciada", valores, campos, condicao)
    
    
'    Campos = "tgc_status"
'    Valores = Bdados.PreparaValor(1)
'    Bdados.AtualizaDados "Tab_Grafica_Credenciada", Valores, Campos, Condicao
End Function

Public Function Credenciar(InscricaoMunicipal As String, DataInicio As String, Validade As String) As Boolean
    Dim campos As String
    Dim valores As String
    Dim condicao As String
    campos = "tgc_num_credenciamento,tgc_tci_im,tgc_data_inicio,tgc_validade,tgc_situacao,tgc_status,tgc_tus_cod_usuario"
    valores = Bdados.PreparaValor(GeraNumero(Credenciamento), InscricaoMunicipal, Bdados.Converte(DataInicio, TCDataHora), Bdados.Converte(Validade, TCDataHora), 0, 0, AplicacoesVTFuncoes.Usuario)
    Credenciar = Bdados.InsereDados("Tab_Grafica_Credenciada", valores, campos)
    'desabilita anteriores
    campos = "tgc_status"
    valores = Bdados.PreparaValor(1)
    condicao = "tgc_tci_im='" & InscricaoMunicipal & "' and tgc_status = 0 and tgc_situacao = 1"
    Bdados.AtualizaDados "Tab_Grafica_Credenciada", valores, campos, condicao
End Function

Public Function GeraNumero(Operacao As TipoNumero) As Double
    Dim Sql As String
    Dim rs As VSRecordset

    Sql = "Update Tab_Num_Correlativo set tnc_sequencia =tnc_sequencia+1 where tnc_tipo_operacao = " & Operacao
        '& " And tnc_Ano = " & Right(Format(Date, "dd/mm/yyyy"), 1)
    Bdados.Executa Sql
    Sql = "SELECT tnc_sequencia from tab_num_correlativo where " _
        & " tnc_tipo_operacao = " & Operacao
        '& " And tnc_Ano = " & Right(Format(Date, "dd/mm/yyyy"), 1)
    If Bdados.AbreTabela(Sql, rs) Then
        GeraNumero = Operacao & Format(IIf(IsNull(rs(0)), 1, rs(0)), "000000")
    End If
    Bdados.FechaTabela rs
End Function

