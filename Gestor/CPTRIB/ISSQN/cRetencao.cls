VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRetencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================='
'               VsTIss.cRetencao: Usado em TRET101                   '
'                   Feito por Djeison  -  14/04/2003                            '
'=========================================='
Private m_Nota As cNota

Public Property Get Nota() As cNota
    Set Nota = m_Nota
End Property

Public Property Set Nota(ByVal Value As cNota)
    Set m_Nota = Value
End Property

Private Sub Class_Initialize()
    Set m_Nota = New cNota
End Sub

Private Sub Class_Terminate()
    Set m_Nota = Nothing
End Sub

Public Function Salvar(NovoRemetente As Boolean, NumIM As String, DataAtual As Date) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim Conta As New ContaCorrente
    Dim Incidencia As Integer
    
    Salvar = False
    If NovoRemetente Then m_Nota.GravaNovoRemetente
    m_Nota.Cod_Pagamento = Conta.GeraCodPagamento(m_Nota.Cod_Imposto)
    
        Valores = Bdados.PreparaValor(m_Nota.Nota_fiscal, 3, Bdados.Converte(m_Nota.Data_emissao, TCDataHora), Right(m_Nota.Periodo_Ref, 4) & Left(m_Nota.Periodo_Ref, 2), IIf(Trim(NumIM) <> "", NumIM, "11000000-00"), _
                        Bdados.Converte(m_Nota.Total_Nota, TCDuplo), Bdados.Converte(m_Nota.ISS_Devido, TCDuplo), m_Nota.Usuario, Bdados.Converte(m_Nota.Aliquota, TCDuplo), m_Nota.Cod_Pagamento, m_Nota.Cod_Imposto, Bdados.Converte(DataAtual, TCDataHora))
        Campos = "tna_numero_nota,tna_status_nota,tna_data_emissao,tna_periodo,tna_tca_identidade_remetente," & _
                        "tna_valor_nota," & _
                        "tna_valor_imposto,tna_tus_cod_usuario,tna_aliquota,TNA_COD_PAGAMENTO,TNA_TIP_COD_IMPOSTO,TNA_DATA_RECEPCAO"
    If Bdados.GravaDados("Tab_Nota_Avulsa", Valores, Campos, "tna_numero_nota=" & m_Nota.Nota_fiscal) Then
            Incidencia = Imposto.BuscaNumeroIncidencia(IIf(Len(NumIM) = 0, "11000000-00", NumIM), Right(m_Nota.Periodo_Ref, 4) & Left(m_Nota.Periodo_Ref, 2), m_Nota.Cod_Imposto)
            Conta.GeraPagamento IIf(Len(NumIM) = 0, "11000000-00", NumIM), "", m_Nota.Cod_Imposto, Right(m_Nota.Periodo_Ref, 4) & Left(m_Nota.Periodo_Ref, 2), CStr(DataAtual), m_Nota.ISS_Devido, 0, 0, CDbl(m_Nota.Cod_Pagamento), 0, 0, 0, "", , , Incidencia
            Salvar = True
    End If
End Function

Public Sub PreencheGrid(Grid As Object, Nome As String)
    Dim Sql As String
    Dim rs As VSRecordset
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sql = "Select tci_im as IM, tci_nome as Razao,tci_cgc_cpf as CPF_CGC from Tab_Contribuinte where tci_nome like '" & Nome & "%' or tci_nome like '%" & Nome & "%'"
    Sql = Sql & " and tci_tsc_cod_sit_cad =1 order by tci_nome"
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If Not Bdados.AbreTabela(Sql, rs) Then Call Util.Avisa("Nenhum contribuinte encontrado.")
    Bdados.FechaTabela rs
    Grid.Preencher Bdados, Sql, 15, 57, 25
End Sub

Public Sub BuscaAliquota(DataAtual As Date)
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Condicao As String
    
    Sql = "Select tpi_aliquota,tip_nome_imposto,tip_cod_imposto from tab_parametro_imposto,tab_imposto where tpi_tip_cod_imposto = '" & m_Nota.Cod_Imposto & "' and tpi_tip_cod_imposto = tip_cod_imposto"
    
    m_Nota.Cod_Imposto = BuscaCodigo("SELECT TIP_COD_IMPOSTO FROM TAB_IMPOSTO WHERE TIP_SIGLA_IMPOSTO = '" & Imposto.NomeTributo(ttr_ISSQNRET) & "'")
    Condicao = " and tpi_ano_imposto= '" & Imposto.BuscaAnoImposto(m_Nota.Cod_Imposto, (Right(Year(DataAtual), 4))) & "'"
    If Bdados.AbreTabela(Sql & Condicao, rs) Then
        m_Nota.Aliquota = rs(0)
        m_Nota.Cod_Imposto = rs(2)
        m_Nota.Nome_Imposto = rs(1)
    End If
    Bdados.FechaTabela rs
End Sub

