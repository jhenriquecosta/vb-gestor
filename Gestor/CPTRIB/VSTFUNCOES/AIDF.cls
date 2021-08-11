VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AIDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-------------Propiedades e funções para manutenção da tela TAID201 do módulo VStSEV--------------
Private a_Num_Aidf                            As String
Private a_IsncricaoMunicipal                As String
Private a_IsncricaoMunicipalGrafico      As String
Private a_DataAutorizacao                  As Date
Private a_NotaInicial                           As String
Private a_NotaFinal                             As String
Private a_TotalBloco                           As String
Private a_SituacaoAidf                        As String
Private a_Usuario                               As String
Private a_Serie                                  As String
Private a_TipoAidf                             As String
Private a_DataCancelamento             As Date
'-----Dados da busca do contribuinte-----------
Private a_NomeContribuinte              As String
Private a_LogradouroContribuinte      As String
Private a_NomeLogradouroContribuinte  As String
Private a_NumeroContribuinte               As String
Private a_ComplementoContribuinte       As String
Private a_BairroContribuinte                  As String
Private a_CidadeContribuinte                 As String
Private a_Uf_Contribuinte                      As String
Private a_CgcCpfContribuiinte                As String
Private a_DataCredenciada                    As Date
Private a_DataDescredenciada                As Date
Private a_Situacao                                  As String
Private a_Validade                                 As String
Private a_CepContribuinte                       As String
Public Enum TipoAcaoGrafica
    Credenciar = 0
    Descredenciar = 1
End Enum

Public Enum TipoNumero
    Credenciamento = 22
    AIDF = 23
End Enum
Public Enum TipoBuscarTnav101
        eContribuinteNormal = 1
        eContribuinteAvulsa = 2
End Enum
Public Enum TipoBuscas
    eBuscarContribuinte = 1
    eBuscarGrafica = 2
End Enum
Public Property Get Cep_Contribuinte() As String
    Cep_Contribuinte = a_CepContribuinte
End Property
Public Property Let Cep_Contribuinte(ByVal value As String)
    a_CepContribuinte = value
End Property
Public Property Get Validade() As String
    Validade = a_Validade
End Property
Public Property Let Validade(ByVal value As String)
    a_Validade = value
End Property
Public Property Get Situacao() As String
    Situacao = a_Situacao
End Property
Public Property Let Situacao(ByVal value As String)
    a_Situacao = value
End Property
Public Property Get Data_Descredenciada() As Date
    Data_Descredenciada = a_DataDescredenciada
End Property
Public Property Let Data_Descredenciada(ByVal value As Date)
    a_DataDescredenciada = value
End Property
Public Property Get Data_Credenciada() As Date
    Data_Credenciada = a_DataCredenciada
End Property
Public Property Let Data_Credenciada(ByVal value As Date)
    a_DataCredenciada = value
End Property
Public Property Get Data_Cancelamento() As Date
    Data_Cancelamento = a_DataCancelamento
End Property
Public Property Let Data_Cancelamento(ByVal value As Date)
    a_DataCancelamento = value
End Property

Public Property Get Documento_Contribuinte() As String
    Documento_Contribuinte = a_CgcCpfContribuiinte
End Property
Public Property Let Documento_Contribuinte(ByVal value As String)
    a_CgcCpfContribuiinte = value
End Property
Public Property Get Uf_Contribuinte() As String
    Uf_Contribuinte = a_Uf_Contribuinte
End Property
Public Property Let Uf_Contribuinte(ByVal value As String)
    a_Uf_Contribuinte = value
End Property
Public Property Get Cidade_Contribuinte() As String
    Cidade_Contribuinte = a_CidadeContribuinte
End Property
Public Property Let Cidade_Contribuinte(ByVal value As String)
    a_CidadeContribuinte = value
End Property
Public Property Get Bairro_Contribuinte() As String
    Bairro_Contribuinte = a_BairroContribuinte
End Property
Public Property Let Bairro_Contribuinte(ByVal value As String)
    a_BairroContribuinte = value
End Property
Public Property Get Complemento_Contribuinte() As String
    Complemento_Contribuinte = a_ComplementoContribuinte
End Property
Public Property Let Complemento_Contribuinte(ByVal value As String)
    a_ComplementoContribuinte = value
End Property
Public Property Get Numero_Contribuinte() As String
    Numero_Contribuinte = a_NumeroContribuinte
End Property
Public Property Let Numero_Contribuinte(ByVal value As String)
    a_NumeroContribuinte = value
End Property
Public Property Get Nome_Logrdouro_Contribuinte() As String
    Nome_Logrdouro_Contribuinte = a_NomeLogradouroContribuinte
End Property
Public Property Let Nome_Logrdouro_Contribuinte(ByVal value As String)
    a_NomeLogradouroContribuinte = value
End Property
Public Property Get Logrdouro_Contribuinte() As String
    Logrdouro_Contribuinte = a_LogradouroContribuinte
End Property
Public Property Let Logrdouro_Contribuinte(ByVal value As String)
    a_LogradouroContribuinte = value
End Property
Public Property Get Nome_Contribuinte() As String
    Nome_Contribuinte = a_NomeContribuinte
End Property
Public Property Let Nome_Contribuinte(ByVal value As String)
    a_NomeContribuinte = value
End Property
Public Property Get Numero_AIDF() As String
    Numero_AIDF = a_Num_Aidf
End Property
Public Property Let Numero_AIDF(ByVal value As String)
    a_Num_Aidf = value
End Property
Public Property Get Inscricao_Municipal() As String
    Inscricao_Municipal = a_IsncricaoMunicipal
End Property
Public Property Let Inscricao_Municipal(ByVal value As String)
    a_IsncricaoMunicipal = value
End Property
Public Property Get Inscricao_Municipal_Grafica() As String
    Inscricao_Municipal_Grafica = a_IsncricaoMunicipalGrafico
End Property
Public Property Let Inscricao_Municipal_Grafica(ByVal value As String)
    a_IsncricaoMunicipalGrafico = value
End Property
Public Property Get Data_Autorizacao() As Date
    Data_Autorizacao = a_DataAutorizacao
End Property
Public Property Let Data_Autorizacao(ByVal value As Date)
    a_DataAutorizacao = value
End Property
Public Property Get Nota_Inicial() As String
    Nota_Inicial = a_NotaInicial
End Property
Public Property Let Nota_Inicial(ByVal value As String)
    a_NotaInicial = value
End Property
Public Property Get Nota_Final() As String
    Nota_Final = a_NotaFinal
End Property
Public Property Let Nota_Final(ByVal value As String)
    a_NotaFinal = value
End Property
Public Property Get Total_Bloco() As String
    Total_Bloco = a_TotalBloco
End Property
Public Property Let Total_Bloco(ByVal value As String)
    a_TotalBloco = value
End Property
Public Property Get Situacao_Aidf() As String
    Situacao_Aidf = a_SituacaoAidf
End Property
Public Property Let Situacao_Aidf(ByVal value As String)
    a_SituacaoAidf = value
End Property
Public Property Get Usuario() As String
    Usuario = a_Usuario
End Property
Public Property Let Usuario(ByVal value As String)
    a_Usuario = value
End Property
Public Property Get Serie() As String
    Serie = a_Serie
End Property
Public Property Let Serie(ByVal value As String)
    a_Serie = value
End Property
Public Property Get Tipo_Aidf() As String
    Tipo_Aidf = a_TipoAidf
End Property
Public Property Let Tipo_Aidf(ByVal value As String)
    a_TipoAidf = value
End Property
Public Function GravarTaid201() As Boolean
    Dim Campos As String
    Dim Valores  As String
    Valores = Bdados.PreparaValor(a_Num_Aidf, a_IsncricaoMunicipal, a_IsncricaoMunicipalGrafico, a_DataAutorizacao, a_NotaInicial, a_NotaFinal, a_TotalBloco, a_SituacaoAidf, a_Usuario, a_Serie, a_TipoAidf)
    Campos = "tai_num_aidf,tai_tci_im,tai_tgc_tci_im,tai_data_autorizacao,tai_nota_inicial,tai_nota_final,tai_total_blocos,tai_situacao_aidf,tai_tus_cod_usuario,tai_serie,tai_tipo_aidf"
    GravarTaid201 = Bdados.InsereDados("Tab_Aidf", Valores, Campos)
End Function
Public Function BuscarTaid201(InscricaoMunicipal As String, TipoBusca As TipoBuscas) As Boolean
    Dim Sql        As String
    Dim rsConsulta As VSRecordset
    Select Case TipoBusca
        Case eBuscarContribuinte
             Sql = "Select * "
             Sql = Sql & " FROM tab_Contribuinte"
             Sql = Sql & " where tci_tsc_cod_sit_cad=1"
             Sql = Sql & " and tci_im = '" & InscricaoMunicipal & "'"
                
        Case eBuscarGrafica
            Sql = "Select tci_nome,tci_logradouro,tci_nome_logradouro,"
            Sql = Sql & " tci_numero,tci_complemento,tci_bairro,tci_cidade,"
            Sql = Sql & " tci_uf,tci_cgc_cpf "
            Sql = Sql & " FROM tab_Contribuinte,Tab_Grafica_Credenciada "
            Sql = Sql & " where tgc_tci_im=tci_im"
            Sql = Sql & " and tci_tsc_cod_sit_cad=1"
            Sql = Sql & " and tgc_situacao=0 "
            Sql = Sql & " and tgc_status =0"
            Sql = Sql & " and tci_im = '" & InscricaoMunicipal & "'"
    End Select
    If Bdados.AbreTabela(Sql, rsConsulta) Then
        BuscarTaid201 = True
        a_NomeContribuinte = "" & rsConsulta.Fields(2)
        a_NomeLogradouroContribuinte = "" & rsConsulta("tci_logradouro") & " " & rsConsulta("tci_nome_logradouro") & "," & rsConsulta("tci_numero") & " " & rsConsulta("tci_complemento") & " " & rsConsulta("tci_bairro")
        a_CidadeContribuinte = "" & rsConsulta("tci_cidade")
        a_Uf_Contribuinte = "" & rsConsulta("tci_uf")
        a_CgcCpfContribuiinte = rsConsulta!TCI_CGC_CPF
    End If
End Function
Public Function PreencherGridAidf(Grid As Object, Optional NumAidf As String, Optional DtAutorizacao As String, Optional InscricaoMunicipal As String, Optional Contribuinte As String, Optional Periodo1 As String, Optional Periodo2 As String) As Boolean
    Dim Sql      As String
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = " Select  tai_num_aidf AS [AIDF],tai_tci_im AS CONTRIBUINTE," & _
            " tai_tgc_tci_im AS [GRAFICA],tai_nota_inicial AS INICIAL,tai_nota_final AS FINAL," & _
            " tai_total_blocos AS [BLOCOS],tai_serie AS [SÉRIE], " & _
            " tai_data_autorizacao As [AUTORIZAÇÃO],tai_tipo_aidf AS TIPO," & _
            " tci_nome AS NOME ,tci_logradouro as TTL,tci_nome_logradouro AS LOGR,tci_numero AS NUM, " & _
            " tci_complemento AS COMPL,tci_bairro AS BAIRRO,tci_cidade AS CIDADE,tci_uf AS UF" & _
            " from tab_aidf, tab_contribuinte where tai_tci_im = tci_im"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = " Select  tai_num_aidf AS AIDF,tai_tci_im AS CONTRIBUINTE," & _
            " tai_tgc_tci_im AS GRAFICA,tai_nota_inicial AS INICIAL,tai_nota_final AS FINAL," & _
            " tai_total_blocos AS BLOCOS,tai_serie AS SÉRIE, " & _
            " tai_data_autorizacao As AUTORIZAÇÃO,tai_tipo_aidf AS TIPO," & _
            " tci_nome AS NOME ,tci_logradouro as TTL,tci_nome_logradouro AS LOGR,tci_numero AS NUM, " & _
            " tci_complemento AS COMPL,tci_bairro AS BAIRRO,tci_cidade AS CIDADE,tci_uf AS UF" & _
            " from tab_aidf, tab_contribuinte where tai_tci_im = tci_im"
    End If
    If Trim(NumAidf) <> "" Then
        Sql = Sql & " and tai_num_aidf = " & Trim(NumAidf)
    ElseIf Trim(DtAutorizacao) <> "" Then
        Sql = Sql & " and tai_data_autorizacao = " & Bdados.Converte(DtAutorizacao, TCDataHora)
    ElseIf Trim(InscricaoMunicipal) <> "" Then
        Sql = Sql & " and tai_tci_im ='" & Trim(InscricaoMunicipal) & "'"
    ElseIf Trim(Contribuinte) <> "" Then
        Sql = Sql & " and tci_nome like '%" & Contribuinte & "%'"
    ElseIf Trim(Periodo1) <> "" And Trim(Periodo1) <> "" Then
        Sql = Sql & " and tai_data_autorizacao >= " & Bdados.Converte(Periodo1, TCDataHora) & " And tai_data_autorizacao <= " & Bdados.Converte(Periodo2, TCDataHora)
    End If
    'Sql = Sql & " and tai_situacao_aidf =0"
    If Not Grid.Preencher(Bdados, Sql, 900, 1300, 1150, 800, 800, 800, 700, 1300, 700, 2050, 0, 0, 0, 0, 0, 0, 0) Then
        PreencherGridAidf = False
    End If

End Function
Public Function BuscarAidf(NumAidf As String, Optional Situacao As String) As Boolean
    Dim Sql              As String
    Dim SqlAux        As String
    Dim rsConsulta   As VSRecordset
    If NumAidf <> "" Then
        SqlAux = " and  tai_num_aidf = " & NumAidf
    End If
    If Situacao <> "" Then
        SqlAux = SqlAux & " and tai_situacao_aidf = " & Situacao
    End If
    Sql = "Select * from tab_aidf "
    If SqlAux <> "" Then Sql = Sql & " where " & Right(SqlAux, Len(SqlAux) - 4)
    If Bdados.AbreTabela(Sql, rsConsulta) Then
        BuscarAidf = True
        a_Num_Aidf = "" & rsConsulta.Fields(8)
        a_IsncricaoMunicipal = "" & rsConsulta.Fields("tai_tci_im")
        a_IsncricaoMunicipalGrafico = "" & rsConsulta.Fields("tai_tgc_tci_im")
        a_NotaInicial = "" & rsConsulta.Fields("tai_nota_inicial")
        a_NotaFinal = "" & rsConsulta.Fields("tai_nota_final")
        a_TotalBloco = "" & rsConsulta.Fields("tai_total_blocos")
        a_Serie = "" & rsConsulta.Fields("tai_serie")
        a_TipoAidf = "" & rsConsulta.Fields(12)
    End If
End Function
Public Function MontaGridTaid201A(Formulario As Object, Grid As Object, Nome As String)
    Dim Rs As VSRecordset
    Dim Sql As String
        If Formulario.Tag = "TAID201" Then
            Sql = "Select tci_im as IM, tci_nome as Razao from Tab_Contribuinte where tci_nome like '" & Nome & "%' or tci_nome like '% " & Nome & "%'"
            Sql = Sql & " and tci_tsc_cod_sit_cad =1"
        Else
            Sql = "SELECT Tab_Grafica_Credenciada.tgc_tci_im, Tab_Contribuinte.tci_nome" & _
                " FROM Tab_Contribuinte INNER JOIN Tab_Grafica_Credenciada " & _
                "ON Tab_Contribuinte.tci_im = Tab_Grafica_Credenciada.tgc_tci_im" & _
                " GROUP BY Tab_Grafica_Credenciada.tgc_tci_im, Tab_Contribuinte.tci_nome " & _
                "HAVING TCI_NOME LIKE '" & Nome & "%' or tci_nome like '% " & Nome & "%'"
        End If
        If Not Bdados.AbreTabela(Sql, Rs) Then
            Call Util.Avisa("Nenhum contribuinte encontrado.")
        End If
        Bdados.FechaTabela Rs
        MontaGrid Bdados, Grid, Sql, 1400
End Function
Public Function GravarTaid202() As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Condicao As String
    
    Campos = "tai_data_cancelamento,tai_situacao_aidf,tai_tus_cod_usuario"
    Valores = Bdados.PreparaValor(a_DataCancelamento, a_SituacaoAidf, a_Usuario)
    Condicao = "tai_num_aidf = " & a_Num_Aidf
    GravarTaid202 = Bdados.AtualizaDados(" tab_aidf ", Valores, Campos, Condicao)
End Function
Public Function BuscarTaid402(InsCricaoGrafica As String) As Boolean
        Dim Rs As VSRecordset
        Dim Sql As String

        Sql = "Select tci_nome,tci_logradouro,tci_nome_logradouro," & _
        " tci_numero,tci_complemento,tci_bairro,tci_cidade,tci_uf,tci_cgc_cpf,tgc_data_inicio,tgc_validade,tgc_situacao,TGC_DATA_DESCREDENCIAMENTO FROM tab_Contribuinte,Tab_Grafica_Credenciada where  tci_im='" & InsCricaoGrafica & _
        "' and tgc_tci_im=tci_im order by tgc_data_inicio desc,tgc_status asc"
        If Bdados.AbreTabela(Sql, Rs) Then
            BuscarTaid402 = True
            a_NomeContribuinte = "" & Rs!tci_nome
            a_NomeLogradouroContribuinte = "" & Rs("tci_logradouro") & " " & Rs("tci_nome_logradouro") & "," & Rs("tci_numero") & " " & Rs("tci_complemento") & " " & Rs("tci_bairro")
            a_CidadeContribuinte = "" & Rs("tci_cidade")
            a_Uf_Contribuinte = "" & Rs("tci_uf")
            a_CgcCpfContribuiinte = Rs!TCI_CGC_CPF
            a_Situacao = IIf(Rs!tgc_situacao = 0, "CREDENCIADA", "DESCREDENCIADA")
            a_DataCredenciada = "" & Rs!tgc_data_inicio
            a_Validade = "" & Rs!tgc_validade
            If Rs!TGC_DATA_DESCREDENCIAMENTO <> "null" Then
                a_DataDescredenciada = "" & Rs!TGC_DATA_DESCREDENCIAMENTO
            End If
        End If
        Bdados.FechaTabela Rs
    
End Function
Public Function PreencherGridTaid402(Formulario As Object, Grid As Object, Optional InscricaoMunicipalGrafica As String, Optional CgcGrafica As String, Optional Nome As String) As Boolean
    Dim Sql As String
    
    If Formulario.Tag = "TAID201" Then
        Sql = "Select tci_im as IM, tci_nome as Razao from Tab_Contribuinte where tci_nome like '" & Nome & "%' or tci_nome like '% " & Nome & "%'"
        Sql = Sql & " and tci_tsc_cod_sit_cad =1"
    Else
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "SELECT Tab_Grafica_Credenciada.tgc_tci_im AS IM, Tab_Contribuinte.tci_nome AS Razao, Tab_Contribuinte.tci_cgc_cpf AS [CGC/CPF]" & _
            " FROM Tab_Contribuinte INNER JOIN Tab_Grafica_Credenciada ON Tab_Contribuinte.tci_im = Tab_Grafica_Credenciada.tgc_tci_im" & _
            " WHERE  tgc_tci_im not in (select tgc_tci_im from tab_grafica_credenciada where tgc_status=0 and tgc_situacao=1) "
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "SELECT Tab_Grafica_Credenciada.tgc_tci_im AS IM, Tab_Contribuinte.tci_nome AS Razao, Tab_Contribuinte.tci_cgc_cpf AS CGC_CPF" & _
            " FROM Tab_Contribuinte INNER JOIN Tab_Grafica_Credenciada ON Tab_Contribuinte.tci_im = Tab_Grafica_Credenciada.tgc_tci_im" & _
            " WHERE  tgc_tci_im not in (select tgc_tci_im from tab_grafica_credenciada where tgc_status=0 and tgc_situacao=1) "
    End If
    
        
        Sql = Sql & " GROUP BY Tab_Grafica_Credenciada.tgc_tci_im, Tab_Contribuinte.tci_nome, " & _
        "Tab_Grafica_Credenciada.TGC_DATA_DESCREDENCIAMENTO," & _
        " Tab_Contribuinte.tci_cgc_cpf "
        
        If Trim(InscricaoMunicipalGrafica) <> "" Then
            Sql = Sql & " HAVING (dbo.Tab_Grafica_Credenciada.tgc_tci_im = '" & InscricaoMunicipalGrafica & "')"
        ElseIf Trim(CgcGrafica) <> "" Then
            Sql = Sql & " HAVING (Tab_Contribuinte.tci_cgc_cpf = '" & CgcGrafica & "')"
        ElseIf Trim(Nome) <> "" Then
            Sql = Sql & " HAVING (dbo.Tab_Contribuinte.tci_nome like '%" & Nome & "%')"
        End If
    End If
    
    Grid.Preencher Bdados, Sql, 1100, 3000, 2000
End Function
Public Sub HabilitaCamposTnav101(Formulario As Object, Status As Boolean, Tipo As String)
    If Tipo = "R" Then
        Formulario.txtEndereco_Rem.Enabled = Status
        Formulario.txtBairro_Rem.Enabled = Status
        Formulario.TxtCep_Rem.Enabled = Status
        Formulario.CboUF_Rem.Enabled = Status
        Formulario.txtMunicipio_Rem.Enabled = Status
    ElseIf Tipo = "D" Then
        Formulario.txtEndereco_Dest.Enabled = Status
        Formulario.txtBairro_Dest.Enabled = Status
        Formulario.TxtCep_Dest.Enabled = Status
        Formulario.CboUF_Dest.Enabled = Status
        Formulario.txtMunicipio_Dest.Enabled = Status
    End If
End Sub
Public Function MontaComboTnav101(Combo As Object)
    Dim Sql As String
    Dim Rs  As VSRecordset
    Sql = " SELECT tae_cae, Tae_nome  from tab_atividade_economica"
    If Bdados.AbreTabela(Sql, Rs) Then
        Do Until Rs.EOF
            Combo.AddItem Rs.Fields(0) & " - " & Rs.Fields(1)
            Rs.MoveNext
        Loop
    End If
End Function
Private Function RetornaIdex(Combo As Object, Texto As String)
    Dim Ilaco As Integer
    For Ilaco = 1 To Combo.ListCount - 1
        If Combo.List(Ilaco) = Texto Then
            RetornaIdex = Ilaco
            Exit For
        End If
    Next
End Function
Public Function BuscarDestinatarioTnav101(Formulario As Object, objGrid As Object, Optional InscricaoMunicipal_Cpf As String, Optional ByRef objComboAtividade As Object, Optional ByRef NovoDestino As Boolean, Optional ByRef objInscricaoDestino As Object, Optional ByRef objServico As Object, Optional ByRef objAliquota As Object, Optional Aliquotas As Double, Optional ByRef objNomeDestino As Object) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Imposto As New VSImposto
    Sql = "Select * FROM tab_Contribuinte where tci_im='" & InscricaoMunicipal_Cpf & "' or tci_cgc_cpf='" & InscricaoMunicipal_Cpf & "' and tci_tsc_cod_sit_cad =1"
    If Bdados.AbreTabela(Sql, Rs) Then
        Me.HabilitaCamposTnav101 Formulario, False, "D"
        BuscarDestinatarioTnav101 = True
        a_NomeContribuinte = Rs!tci_nome
        a_NomeLogradouroContribuinte = Rs!tci_logradouro & " " & Rs!tci_nome_logradouro & " " & Rs!tci_NUMERO & " " & Rs!tci_COMPLEMENTO
        a_BairroContribuinte = Rs!tci_BAIRRO
        a_CepContribuinte = Rs!tci_cep
        a_Uf_Contribuinte = Rs!tci_uf
        a_CidadeContribuinte = Rs!tci_cidade
        If Not IsNull(Rs!tci_tae_cae) Then
            Sql = "SELECT tae_nome  from tab_atividade_economica where tae_cae = " & Rs!tci_tae_cae
            If Bdados.AbreTabela(Sql, Rs) Then
                objComboAtividade.ListIndex = RetornaIdex(objComboAtividade, Rs(0))
            End If
        End If
        NovoDestino = False
        objInscricaoDestino.Enabled = False
    Else
        Sql = "Select * FROM tab_Contribuinte_Avulso where tca_identidade='" & InscricaoMunicipal_Cpf & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            BuscarDestinatarioTnav101 = True
            a_NomeContribuinte = Rs!tca_nome
            a_NomeLogradouroContribuinte = Rs!tca_endereco
            a_BairroContribuinte = Rs!tca_bairro
            a_CepContribuinte = Rs!tca_cep
            a_Uf_Contribuinte = Rs!tca_uf
            a_CidadeContribuinte = Rs!tca_cidade
            If Not IsNull(Rs!tca_tae_cae) Then
                Sql = "SELECT tae_nome  from tab_atividade_economica where tae_cae = " & Rs!tca_tae_cae
                If Bdados.AbreTabela(Sql, Rs) Then
                    objComboAtividade.ListIndex = ListIndexDe(objComboAtividade, Rs(0))
                End If
            End If
            NovoDestino = False
            objInscricaoDestino.Enabled = False
            objServico.SetFocus
        Else
            Me.HabilitaCamposTnav101 Formulario, True, "D"
            objInscricaoDestino.Tag = InscricaoMunicipal_Cpf
            objAliquota = Aliquotas
            objInscricaoDestino = objInscricaoDestino.Tag
            objNomeDestino.SetFocus
            NovoDestino = True
        End If
    End If
    Bdados.FechaTabela Rs
End Function
Public Function BuscarRemessaTnav101(Formulario As Object, objInscricao As Object, Inscricao As String, ObjUf As Object, NovoRemetentes As Boolean, ObjAliquotas As Object, Aliquota As Double, objNomeContribuinte As Object) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Imposto As New VSImposto
    Sql = "Select * FROM tab_Contribuinte where tci_im='" & Inscricao & "' or tci_cgc_cpf='" & Inscricao & "' and tci_tsc_cod_sit_cad =1"
    If Bdados.AbreTabela(Sql, Rs) Then
        Me.HabilitaCamposTnav101 Formulario, False, "R"
        BuscarRemessaTnav101 = True
        a_NomeContribuinte = Rs!tci_nome
        a_NomeLogradouroContribuinte = Rs!tci_logradouro & " " & Rs!tci_nome_logradouro & " " & Rs!tci_NUMERO & " " & Rs!tci_COMPLEMENTO
        a_BairroContribuinte = "" & Rs!tci_BAIRRO
        a_CepContribuinte = "" & Rs!tci_cep
        a_Uf_Contribuinte = Rs!tci_uf
        a_CidadeContribuinte = Rs!tci_cidade
        NovoRemetentes = False
        objInscricao.Enabled = False
    Else
        Sql = "Select * FROM tab_Contribuinte_Avulso where tca_identidade='" & Inscricao & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            BuscarRemessaTnav101 = True
            Me.HabilitaCamposTnav101 Formulario, False, "R"
            a_NomeContribuinte = Rs!tca_nome
            a_NomeLogradouroContribuinte = Rs!tca_endereco
            a_BairroContribuinte = Rs!tca_bairro
            a_CepContribuinte = Rs!tca_cep
            a_Uf_Contribuinte = Rs!tca_uf
            a_CidadeContribuinte = Rs!tca_cidade
            NovoRemetentes = False
            objInscricao.Enabled = False
        Else
            Me.HabilitaCamposTnav101 Formulario, True, "R"
            NovoRemetentes = True
            objInscricao.Tag = objInscricao
            ObjAliquotas = Aliquota
            objInscricao = objInscricao.Tag
            objNomeContribuinte.SetFocus
        End If
    End If
    Bdados.FechaTabela Rs
End Function

Public Function SalvarTnav101(Formulario As Object, NovoRemetente As Boolean, NovoDestino As Boolean, ByRef objLblNotaFiscal As Object, GridNotas As Object, ImDest As String, Nome_Dest As String, Endereco_Dest As String, Bairro_Dest As String, Municipio_Dest As String, UF_Dest As String, Cep_Dest As String, Usuario As String, Data As String, TotalNota As String, Iss As String, Aliquota As String, Periodo As String, Material As String, IRRF As String) As Boolean
    On Error GoTo error
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Valores As String
    Dim Campos As String
    Dim I As Byte
    If NovoRemetente Then
        Valores = Bdados.PreparaValor(a_IsncricaoMunicipal, a_NomeContribuinte, a_NomeLogradouroContribuinte, a_BairroContribuinte, a_CidadeContribuinte, a_Uf_Contribuinte, a_CepContribuinte, a_Usuario)
        Campos = "tca_identidade,tca_nome,tca_endereco,tca_bairro,tca_cidade,tca_uf,tca_cep,tca_tus_cod_usuario"
        Bdados.InsereDados "Tab_Contribuinte_Avulso", Valores, Campos
    End If
    If NovoDestino Then
        Valores = Bdados.PreparaValor(ImDest, Nome_Dest, Endereco_Dest, Bairro_Dest, Municipio_Dest, UF_Dest, Cep_Dest, Usuario)
        Campos = "tca_identidade,tca_nome,tca_endereco,tca_bairro,tca_cidade,tca_uf,tca_cep,tca_tus_cod_usuario"
        Bdados.InsereDados "Tab_Contribuinte_Avulso", Valores, Campos
    End If
    objLblNotaFiscal = Imposto.GeraNumNota(1, 65)
    Valores = Bdados.PreparaValor(objLblNotaFiscal, 0, Bdados.Converte(Format(Data, "dd/mm/yyyy"), TCDataHora), IIf(Len(a_IsncricaoMunicipal) = 12, a_IsncricaoMunicipal, Const_ImAvulso), ImDest, _
    Bdados.Converte(TotalNota, TCDuplo), Bdados.Converte(Iss, TCDuplo), Usuario, Bdados.Converte(Aliquota, TCDuplo), Right(Periodo, 4) & Left(Periodo, 2), Bdados.Converte(Nvl(Material, 0), TCDuplo), Bdados.Converte(Nvl(IRRF, 0), TCDuplo))
    Campos = "tna_numero_nota,tna_status_nota,tna_data_emissao,tna_tca_identidade_remetente," & _
        "tna_tca_identidade_destinatario,tna_valor_nota," & _
        "tna_valor_imposto,tna_tus_cod_usuario,tna_aliquota,tna_periodo,TNA_MATERIAL,TNA_IRRF"
    SalvarTnav101 = Bdados.InsereDados("Tab_Nota_Avulsa", Valores, Campos)
    Campos = "tin_tna_numero_nota,tin_descricao_servico,tin_valor,tin_quantidade," & _
        "tin_unidade"
    For I = 1 To GridNotas.ListItems.Count
        Valores = Bdados.PreparaValor(objLblNotaFiscal, GridNotas.ListItems.Item(I).Text, Bdados.Converte(GridNotas.ListItems.Item(I).SubItems(3), TCDuplo), GridNotas.ListItems.Item(I).SubItems(1), GridNotas.ListItems.Item(I).SubItems(2))
        Bdados.InsereDados "Tab_Item_Nota_Avulsa", Valores, Campos
    Next
    Exit Function
error:
    Util.Mensagem Err.Number & " - " & Err.Description
End Function

'--------------------------FIM DE PROPIEDADES--------------------------------





Public Function GeraNumero(Operacao As TipoNumero) As Double
    Dim Sql As String
    Dim Rs As VSRecordset

    Sql = "Update Tab_Num_Correlativo set tnc_sequencia =tnc_sequencia+1 where tnc_tipo_operacao = " & Operacao
        '& " And tnc_Ano = " & Right(Format(Date, "dd/mm/yyyy"), 1)
    Bdados.Executa Sql
    Sql = "SELECT tnc_sequencia from tab_num_correlativo where " _
        & " tnc_tipo_operacao = " & Operacao
        '& " And tnc_Ano = " & Right(Format(Date, "dd/mm/yyyy"), 1)
    If Bdados.AbreTabela(Sql, Rs) Then
        GeraNumero = Operacao & Format(IIf(IsNull(Rs(0)), 1, Rs(0)), "000000")
    End If
    Bdados.FechaTabela Rs
End Function

Public Function VerificaCredenciamentoGrafica(InscricaoMunicipal As String, Acao As TipoAcaoGrafica) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset

    Sql = " select tgc_validade from Tab_Grafica_Credenciada where tgc_tci_im='" & InscricaoMunicipal & _
        "' and tgc_status=0 and tgc_situacao=" & Acao
        VerificaCredenciamentoGrafica = False
    If Bdados.AbreTabela(Sql, Rs) Then
        If Acao = Credenciar Then
            Util.Informa "Gráfica com credenciamento ainda válido até " & Rs!tgc_validade & "."
        Else
            Util.Informa "Gráfica já descredenciada em " & Rs!tgc_validade & "."
        End If
        VerificaCredenciamentoGrafica = True
    End If
End Function


Public Function Gravar(InscricaoMunicipal As String, DataInicio As String, Validade As String, Acao As TipoAcaoGrafica) As Boolean
    Dim Campos As String
    Dim Valores As String
    Dim Condicao As String
    Campos = "tgc_num_credenciamento,tgc_tci_im,tgc_data_inicio,tgc_validade,tgc_situacao,tgc_status,tgc_tus_cod_usuario"
    Valores = Bdados.PreparaValor(GeraNumero(Credenciamento), InscricaoMunicipal, Bdados.Converte(DataInicio, TCDataHora), Bdados.Converte(Validade, TCDataHora), Acao, 0, AplicacoesVTFuncoes.Usuario)
    If Acao = Descredenciar Then
        Valores = Valores & Bdados.PreparaValor(Bdados.Converte(Date, TCDataHora))
        Campos = Campos & ",TGC_DATA_DESCREDENCIAMENTO"
    End If
    Gravar = Bdados.InsereDados("Tab_Grafica_Credenciada", Valores, Campos)
    'DESABILITA SITUACOES ANTERIORES
    Condicao = "tgc_tci_im='" & InscricaoMunicipal & "' and tgc_status=0 and tgc_situacao=" & IIf(Acao = 0, 1, 0)
    Campos = "tgc_status"
    Valores = Bdados.PreparaValor(1)
    Bdados.AtualizaDados "Tab_Grafica_Credenciada", Valores, Campos, Condicao
End Function
