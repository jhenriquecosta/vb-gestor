VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNotaAidf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_ImContribuinte              As String
Private m_ImGrafica                     As String
Private m_DataAutorizacao             As String
Private m_Validade                       As String
Private m_NotaInicial                    As Double
Private m_NotaFinal                      As Double
Private m_TotalBlocos                   As Integer
Private m_SituacaoAidf                 As Integer
Private m_NumAidf                      As Double
Private m_DataCancelamento        As String
Private m_CodUsuario                  As String
Private m_Serie                           As String
Private m_TipoAidf                      As Integer
Public Especie                             As String
Public NomeSerie                        As String
Public m_Motivo                          As String
Public m_Documento                          As String

Public Property Get Documento() As String
    Documento = m_Documento
End Property
Public Property Let Documento(ByVal Value As String)
    m_Documento = Value
End Property

Public Property Get Motivo() As String
    Motivo = m_Motivo
End Property

Public Property Let Motivo(ByVal Value As String)
    m_Motivo = Value
End Property

Public Property Get TipoAidf() As Integer
    TipoAidf = m_TipoAidf
End Property

Public Property Let TipoAidf(ByVal Value As Integer)
    m_TipoAidf = Value
End Property

Public Property Get Serie() As String
    Serie = m_Serie
End Property

Public Property Let Serie(ByVal Value As String)
    m_Serie = Value
End Property

Public Property Get CodUsuario() As String
    CodUsuario = m_CodUsuario
End Property

Public Property Let CodUsuario(ByVal Value As String)
    m_CodUsuario = Value
End Property

Public Property Get DataCancelamento() As String
    DataCancelamento = m_DataCancelamento
End Property

Public Property Let DataCancelamento(ByVal Value As String)
    m_DataCancelamento = Value
End Property

Public Property Get NumAidf() As Double
    NumAidf = m_NumAidf
End Property

Public Property Let NumAidf(ByVal Value As Double)
    m_NumAidf = Value
End Property

Public Property Get SituacaoAidf() As Integer
    SituacaoAidf = m_SituacaoAidf
End Property

Public Property Let SituacaoAidf(ByVal Value As Integer)
    m_SituacaoAidf = Value
End Property

Public Property Get TotalBlocos() As Integer
    TotalBlocos = m_TotalBlocos
End Property

Public Property Let TotalBlocos(ByVal Value As Integer)
    m_TotalBlocos = Value
End Property

Public Property Get NotaFinal() As Double
    NotaFinal = m_NotaFinal
End Property

Public Property Let NotaFinal(ByVal Value As Double)
    m_NotaFinal = Value
End Property

Public Property Get NotaInicial() As Double
    NotaInicial = m_NotaInicial
End Property

Public Property Let NotaInicial(ByVal Value As Double)
    m_NotaInicial = Value
End Property

Public Property Get Validade() As String
    Validade = DateAdd("yyyy", Nvl(Temp.PegaParametro(Bdados, "VALIDADE AIDF"), 1), m_DataAutorizacao)
End Property

Public Property Let Validade(ByVal Value As String)
    m_Validade = Value
End Property

Public Property Get DataAutorizacao() As String
    DataAutorizacao = m_DataAutorizacao
End Property

Public Property Let DataAutorizacao(ByVal Value As String)
    m_DataAutorizacao = Value
End Property

Public Property Get ImGrafica() As String
    ImGrafica = m_ImGrafica
End Property

Public Property Let ImGrafica(ByVal Value As String)
    m_ImGrafica = Value
End Property

Public Property Get ImContribuinte() As String
    ImContribuinte = m_ImContribuinte
End Property

Public Property Let ImContribuinte(ByVal Value As String)
    m_ImContribuinte = Value
End Property

Public Function Buscar(NumAidf As Double) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "select * from tab_aidf where tai_num_aidf = '" & NumAidf & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        m_ImContribuinte = "" & rs!tai_tci_im
        m_ImGrafica = "" & rs!tai_tgc_tci_im
        m_DataAutorizacao = "" & rs!tai_data_autorizacao
        m_Validade = "" & rs!tai_validade
        m_NotaInicial = "" & rs!tai_nota_inicial
        m_NotaFinal = "" & rs!tai_nota_final
        m_TotalBlocos = "" & rs!tai_total_blocos
        m_SituacaoAidf = "" & rs!tai_situacao_aidf
        m_NumAidf = "" & rs!tai_num_aidf
        m_DataCancelamento = "" & rs!tai_data_cancelamento
        m_CodUsuario = "" & rs!tai_tus_cod_usuario
        m_Serie = "" & rs!tai_serie
        m_TipoAidf = "" & rs!tai_tipo_aidf
        Buscar = True
    End If
End Function

Public Function GravarNota() As Boolean
    Dim campos As String
    Dim valores  As String
    valores = Bdados.PreparaValor(NumAidf, ImContribuinte, ImGrafica, Bdados.Converte(DataAutorizacao, TCDataHora), NotaInicial, NotaFinal, TotalBlocos, SituacaoAidf, AplicacoesVTFuncoes.Usuario, Serie, TipoAidf, Bdados.Converte(Validade, TCDataHora), m_Documento, Year(m_DataAutorizacao))
    campos = "tai_num_aidf,tai_tci_im,tai_tgc_tci_im,tai_data_autorizacao,tai_nota_inicial,tai_nota_final,tai_total_blocos,tai_situacao_aidf,tai_tus_cod_usuario,tai_serie,tai_tipo_aidf, tai_validade,TAI_TIPO_DOCUMENTO,TAI_ANO"
    GravarNota = Bdados.InsereDados("Tab_Aidf", valores, campos)
End Function

Public Sub PreencherCboEspecie(Combo As Object)
    Dim Sql As String
    Sql = "SELECT TGE_NOME, TGE_CODIGO FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =701 ORDER BY TGE_CODIGO ASC"
    Combo.Preencher Bdados, Sql
End Sub

Public Sub PreencherCboSerie(Combo As Object)
    Dim Sql As String
    Sql = "SELECT TGE_NOME, TGE_CODIGO FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =703 ORDER BY TGE_CODIGO ASC"
    Combo.Preencher Bdados, Sql
End Sub

Public Sub PreencherCboQtd(Combo As Object)
    Dim Sql As String
    Sql = "SELECT TGE_NOME, TGE_CODIGO FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =704 ORDER BY TGE_CODIGO ASC"
    Combo.Preencher Bdados, Sql
End Sub
'SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO >0 and TGE_TIPO =704 ORDER BY TGE_CODIGO ASC


Public Function UltimaNota(Im As String, Serie As String, Especie As String, Optional Doc As String) As Double
    Dim Sql As String
    Dim rs As VSRecordset
    ' and tai_situacao_aidf = 1
    Sql = "Select max(tai_nota_final) from  tab_aidf " & _
            " where tai_tci_im='" & Im & "'" & _
            "  and tai_serie='" & Serie & "' and tai_tipo_aidf=" & Especie & " and tai_situacao_aidf  = 1 "
        If Doc <> "" Then
            Sql = Sql & " and  TAI_TIPO_DOCUMENTO = '" & Trim(Doc) & "'"
        End If
            
        If Bdados.AbreTabela(Sql, rs) Then
            If Not IsNull(rs(0)) Then
                UltimaNota = CDbl(Nvl("" & rs(0), 0))
            Else
                UltimaNota = 0
            End If
        Else
            UltimaNota = 0
        End If
End Function

Public Function CancelarAidf() As Boolean
    Dim valores As String
    Dim campos As String
    Dim condicao As String
    campos = "tai_data_cancelamento,tai_situacao_aidf,tai_tus_cod_usuario,tai_Motivo"
    valores = Bdados.PreparaValor(m_DataCancelamento, m_SituacaoAidf, m_CodUsuario, m_Motivo)
    condicao = "tai_num_aidf = " & m_NumAidf
    If Bdados.AtualizaDados("TAB_AIDF", valores, campos, condicao) Then CancelarAidf = True
End Function

Public Function PreencherGrid(Grid As Object, Optional NumAidf As String, Optional DtAutorizacao As String, Optional InscricaoMunicipal As String, Optional Contribuinte As String, Optional Periodo1 As String, Optional Periodo2 As String, Optional NumNota As String, Optional Documento As String) As Boolean
    Dim Sql      As String
    Sql = " Select  tai_num_aidf AS AIDF, tai_tci_im AS Contribuinte, tci_nome AS Nome, " & _
        " tai_tgc_tci_im AS Gráfica, tai_nota_inicial AS Nota_Inicial, tai_nota_final AS Nota_Final," & _
        " tai_total_blocos AS Blocos, tai_serie AS Série, " & _
        " tai_data_autorizacao As Dt_Autorização, vis_especie_aidf.tge_nome as Espécie," & _
        " vis_status_aidf.tge_nome as Situacao,tai_motivo as Motivo , vis_tipo_impressão_doc.TGE_NOME AS Tipo from tab_aidf,vis_tipo_impressão_doc,tab_contribuinte, vis_especie_aidf,VIS_STATUS_AIDF " & _
        " where tai_tci_im = tci_im and vis_especie_aidf.tge_codigo = tai_tipo_aidf and vis_status_aidf.tge_codigo = tai_situacao_aidf and vis_tipo_impressão_doc.tge_codigo = tai_tipo_documento"
    If Trim(NumAidf) <> "" Then
        Sql = Sql & " and tai_num_aidf = " & Trim(NumAidf)
    End If
    
    If Documento <> "" Then
        Sql = Sql & " and taI_tipo_documento = '" & Documento & "'"
    End If
    
    If Trim(DtAutorizacao) <> "" Then
        Sql = Sql & " and tai_data_autorizacao = " & Bdados.Converte(DtAutorizacao, TCDataHora)
    End If
    If Trim(InscricaoMunicipal) <> "" Then
        Sql = Sql & " and tai_tci_im ='" & Trim(InscricaoMunicipal) & "'"
    End If
    If Trim(Contribuinte) <> "" Then
        Sql = Sql & " and tci_nome like '%" & Contribuinte & "%'"
    End If
    If Trim(Periodo1) <> "" And Trim(Periodo1) <> "" Then
        Sql = Sql & " and tai_data_autorizacao >= " & Bdados.Converte(Periodo1, TCDataHora) & " And tai_data_autorizacao <= " & Bdados.Converte(Periodo2, TCDataHora)
    End If
    If Trim(NumNota) <> "" Then
        Sql = Sql & " and  tai_nota_inicial <= '" & NumNota & "'  and tai_nota_final >= '" & NumNota & "'"
    End If
    If Grid.Preencher(Bdados, Sql) Then
        PreencherGrid = True
    Else
        PreencherGrid = False
    End If
End Function
Public Sub Imprimir(NumAidf As Double)
    Dim a As Byte
    If Me.Buscar(NumAidf) Then
        With Rpt
             If .DefinirArquivo(Bdados, App.Path + "\TAIDF.rpt") Then
                If UCase(AplicacoesVTFuncoes.Municipio) = "BARRA MANSA" Then
                    .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "GFF")
                Else
                    .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                End If
                .Formulas "Sequencia", Format(Me.NotaInicial, "0000") & " a " & Format(Me.NotaFinal, "0000")
                .Formulas "Blocos", Format(Me.TotalBlocos, "00")
                .Formulas "Cidade", AplicacoesVTFuncoes.Municipio
                .Formulas "UF", Temp.PegaParametro(Bdados, "UF")
                '.Formulas "VT_ESPECIE", Especie
                '.Formulas "VT_SERIE", Me.NomeSerie
                .Selecao = "{Tab_Aidf.tai_num_aidf} =" & NumAidf
'                .Formulas "NumAidf", CStr(NumAidf)
                .Titulo = "Autorização de Impressão de Documentos Fiscais"
                .Arvore = False
                .Visualizar
             End If
        End With
        Set Rpt = Nothing
    End If
End Sub

