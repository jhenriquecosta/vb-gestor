VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Atividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'--------------------------------------------------------------------------------
'    Componente  : Atividade
'    Projeto     : VTTrib
'    Descri��o   : Classe de administra��o de atividades econ�micas
'    Modificado  : 14/05/2003
'--------------------------------------------------------------------------------
Private m_Codigo As String
Private m_Nome As String
Private m_GrupoCodigo As Integer
Private m_Valor As Double
Private m_FatorCodigo As Integer
Private m_FatorDescricao As String
Private m_Estimativo As Integer
Private m_RamoCodigo As Integer
Private m_AliquotaTPPC As Double
Private m_AliquotaSPL As Double
Private m_AliquotaPJ As Double

Public Enum InformacaoAtividade
    iaRamo = 1
    iaGrupoAtividade = 2
End Enum

Public Property Get Codigo() As String
    Codigo = m_Codigo
End Property

Public Property Let Codigo(ByVal Value As String)
    m_Codigo = Value
End Property

Public Property Get Nome() As String
    Nome = m_Nome
End Property

Public Property Let Nome(ByVal Value As String)
    m_Nome = Value
End Property

Public Property Get GrupoCodigo() As Integer
    GrupoCodigo = m_GrupoCodigo
End Property

Public Property Let GrupoCodigo(ByVal Value As Integer)
    m_GrupoCodigo = Value
End Property

Public Property Get Valor() As Double
    Valor = m_Valor
End Property

Public Property Let Valor(ByVal Value As Double)
    m_Valor = Value
End Property

Public Property Get FatorCodigo() As Integer
    FatorCodigo = m_FatorCodigo
End Property

Public Property Let FatorCodigo(ByVal Value As Integer)
    m_FatorCodigo = Value
End Property

Public Property Get FatorDescricao() As String
    FatorDescricao = m_FatorDescricao
End Property

Public Property Let FatorDescricao(ByVal Value As String)
    m_FatorDescricao = Value
End Property

Public Property Get Estimativo() As Integer
    Estimativo = m_Estimativo
End Property

Public Property Let Estimativo(ByVal Value As Integer)
    m_Estimativo = Value
End Property

Public Property Get RamoCodigo() As Integer
    RamoCodigo = m_RamoCodigo
End Property

Public Property Let RamoCodigo(ByVal Value As Integer)
    m_RamoCodigo = Value
End Property

Public Property Get AliquotaTPPC() As Double
    AliquotaTPPC = m_AliquotaTPPC
End Property

Public Property Let AliquotaTPPC(ByVal Value As Double)
    m_AliquotaTPPC = Value
End Property

Public Property Get AliquotaSPL() As Double
    AliquotaSPL = m_AliquotaSPL
End Property

Public Property Let AliquotaSPL(ByVal Value As Double)
    m_AliquotaSPL = Value
End Property

Public Property Get AliquotaPJ() As Double
    AliquotaPJ = m_AliquotaPJ
End Property

Public Property Let AliquotaPJ(ByVal Value As Double)
    m_AliquotaPJ = Value
End Property

Public Function Buscar(CodigoAtividade As String, Optional Detalhes As Boolean = True) As Boolean
'--------------------------------------------------------------------------------
' Projeto      :  VTTrib
' Procedimento :  Buscar
' Par�metros   :  CodigoAtividade (String)
' Descri��o    :  Busca os dados da atividade econ�mica informada
' Criado por   :  �derson Robson Bragagnolo
' Maquina      :  Balsas - Planejamento02
' Data         :  14/05/2003
' Altera��o    :  14/05/2003
'--------------------------------------------------------------------------------
    Dim Sql As String
    Dim Rs2 As VSRecordset
    On Error Resume Next
        
    Sql = "SELECT " & IIf(Detalhes, "*", "TAE_CAE") & " FROM TAB_ATIVIDADE_ECONOMICA WHERE TAE_cae = '" & CodigoAtividade & "'"
    If Bdados.AbreTabela(LCase(Sql), Rs2) Then
        m_Codigo = Rs2(0)
        If Detalhes Then
            m_Nome = "" & Rs2!tae_nome
            m_GrupoCodigo = Nvl("" & Rs2!TAE_TGA_COD_GRUPO, 0)
            m_Valor = "" & Rs2!TAE_VALOR
            m_FatorCodigo = Nvl("" & Rs2!TAE_FATOR, 0)
            m_FatorDescricao = "" & Rs2!tae_desc_fator
            m_Estimativo = Nvl("" & Rs2!TAE_ESTIMATIVO, 0)
            m_RamoCodigo = Nvl("" & Rs2!TAE_TRA_COD_RAMO, 0)
            m_AliquotaTPPC = Nvl("" & Rs2!TAE_ALIQUOTA_TPPC, 0)
            m_AliquotaSPL = Nvl("" & Rs2!TAE_ALIQUOTA_SPL, 0)
            m_AliquotaPJ = Nvl("" & Rs2!TAE_ALIQUOTA_PJ, 0)
        End If
        Buscar = True
    End If
End Function
Public Function Gravar() As Boolean
'--------------------------------------------------------------------------------
' Projeto      :  VTTrib
' Procedimento :  Gravar
' Descri��o    :  Gravar os dados a respeito da atividade econ�mica
' Criado por   :  �derson Robson Bragagnolo
' Maquina      :  Balsas - Planejamento02
' Data         :  14/05/2003
' Altera��o    :  14/05/2003
'--------------------------------------------------------------------------------

    Dim Valores As String
    Dim Campos As String
    Dim condicao As String
        
    Campos = "TAE_CAE, TAE_NOME, TAE_TGA_COD_GRUPO, TAE_VALOR, TAE_FATOR, " & _
        " TAE_DESC_FATOR, TAE_ESTIMATIVO, TAE_TRA_COD_RAMO, " & _
        " TAE_ALIQUOTA_TPPC, TAE_ALIQUOTA_SPL, TAE_ALIQUOTA_PJ "
    Valores = Bdados.PreparaValor(m_Codigo, m_Nome, m_GrupoCodigo, m_Valor, _
        m_FatorCodigo, m_FatorDescricao, m_Estimativo, m_RamoCodigo, _
        m_AliquotaTPPC, m_AliquotaSPL, m_AliquotaPJ)
    condicao = "TAE_CAE = '" & m_Codigo & "'"
    If Bdados.GravaDados("TAB_ATIVIDADE_ECONOMICA", Valores, Campos, condicao) Then
        If m_Estimativo = 1 Then
            Bdados.DeletaDados "TAB_ATIVIDADE_ESTIMADA", "TAT_TAE_CAE = '" & m_Codigo & "'"
        End If
        Gravar = True
    End If
End Function

Public Function Excluir(CodigoAtividade As String) As Boolean
'--------------------------------------------------------------------------------
' Projeto      :  VTTrib
' Procedimento :  Exlcuir
' Par�metros   :  CodigoAtividade (String)
' Descri��o    :  Exclus�o da atividade econ�mica informada
' Criado por   :  �derson Robson Bragagnolo
' Maquina      :  Balsas - Planejamento02
' Data         :  14/05/2003
' Altera��o    :  14/05/2003
'--------------------------------------------------------------------------------
    If Bdados.DeletaDados("TAB_ATIVIDADE_ECONOMICA", "TAE_CAE='" & CodigoAtividade & "'") Then
        Bdados.DeletaDados "TAB_ATIVIDADE_ESTIMADA", "TAT_TAE_CAE = " & CodigoAtividade
        Excluir = True
    End If
End Function

Public Function PreencheGrid(Grid As Object, Optional GrupoAtividade As String, Optional CodAtividade As String, Optional Estimativo As String, Optional Nome As String) As Boolean
'--------------------------------------------------------------------------------
' Projeto      :  VTTrib
' Procedimento :  PreencheGrid
' Par�metros   :  Grid (Object)
' Descri��o    :  Monta uma grid com o conte�do de uma query especificada
' Criado por   :  �derson Robson Bragagnolo
' Maquina      :  Balsas - Planejamento02
' Data         :  14/05/2003
' Altera��o    :  14/05/2003
'--------------------------------------------------------------------------------
    Dim Sql As String
    Dim TamanhoColunas As String
    
   If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "SELECT TAE_CAE AS C�digo, TAE_NOME AS Atividade, " & _
                " TGA_NOME AS Grupo,  " & Bdados.Converte("TAE_VALOR", TCDuplo) & " AS [Valor(R$)], " & _
                " TAE_DESC_FATOR AS Fator, TGE_NOME AS Estimado, " & _
                " TAE_ALIQUOTA_TPPC AS TPPC, TAE_ALIQUOTA_SPL AS SPL, TAE_ALIQUOTA_PJ AS [PJ(%)]" & _
                " FROM TAB_ATIVIDADE_ECONOMICA, TAB_GRUPO_ATIVIDADE, VIS_SIM_NAO" & _
                " WHERE TAE_TGA_COD_GRUPO = TGA_COD_GRUPO AND " & _
                " TGE_CODIGO = TAE_ESTIMATIVO "
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
            Sql = "SELECT TAE_CAE AS C�digo, TAE_NOME AS Atividade, " & _
                " TGA_NOME AS Grupo,TO_CHAR(TAE_VALOR,'999999.99')  AS Valor, " & _
                " TAE_DESC_FATOR AS Fator, TGE_NOME AS Estimado, " & _
                " TAE_ALIQUOTA_TPPC AS TPPC, TAE_ALIQUOTA_SPL AS SPL, TAE_ALIQUOTA_PJ AS PJ" & _
                " FROM TAB_ATIVIDADE_ECONOMICA, TAB_GRUPO_ATIVIDADE, VIS_SIM_NAO" & _
                " WHERE TAE_TGA_COD_GRUPO = TGA_COD_GRUPO AND " & _
                " TGE_CODIGO = TAE_ESTIMATIVO "
    End If
    If GrupoAtividade <> "" Then
        Sql = Sql & " and TAE_TGA_COD_GRUPO = '" & GrupoAtividade & "'"
    End If
    If CodAtividade <> "" Then
        Sql = Sql & " and TAE_CAE = '" & CodAtividade & "'"
    End If
    If Estimativo <> "" Then
        Sql = Sql & " and TAE_ESTIMATIVO = '" & Estimativo & "'"
    End If
    If Nome <> "" Then
        Sql = Sql & " and (tae_nome like '" & Nome & "%' or tae_nome like '% " & Nome & "%')"
    End If
     Sql = Sql & " ORDER BY TAE_CAE"
    TamanhoColunas = "1200, 3100, 1800, 1000, 800, 900"
    If Grid.Preencher(Bdados, Sql) Then PreencheGrid = True
    
End Function

Public Function PreencheCombo(Combo As Object, Informacao As InformacaoAtividade) As Boolean
'--------------------------------------------------------------------------------
' Projeto      :  VTTrib
' Procedimento :  PreencheCombo
' Par�metros   :  Grid (Object)
' Descri��o    :  Monta um combobox com o conte�do do enum informado
' Criado por   :  �derson Robson Bragagnolo
' Maquina      :  Balsas - Planejamento02
' Data         :  15/05/2003
' Altera��o    :  15/05/2003
'--------------------------------------------------------------------------------
    Dim Sql As String
    Select Case Informacao
        Case iaRamo: Sql = "SELECT TRA_NOME_RAMO, TRA_COD_RAMO FROM TAB_RAMO_ATIVIDADE ORDER BY TRA_NOME_RAMO"
        Case iaGrupoAtividade: Sql = "SELECT TGA_NOME, TGA_COD_GRUPO FROM TAB_GRUPO_ATIVIDADE order by tga_cod_grupo asc"
    End Select
    
    If Combo.Preencher(Bdados, Sql) Then PreencheCombo = True
    
End Function

Public Function BuscaAliquotaAtividade(Bdados As Object, Im As String, Periodo As String, Optional ByRef ValorFixo As Double) As Double
'--------------------------------------------------------------------------------
' Projeto      :  VTTrib
' Procedimento :  BuscaAliquotaAtividade
' Par�metros   :  IM (String), ValorFixo (Double)
' Descri��o    :  Localiza e retorna o valor fixo ou al�quota correspondente a atividade do contribuinte
' Criado por   :  �derson Robson Bragagnolo
' Maquina      :  Balsas - Cobranca02
' Data         :  16/05/2003
' Altera��o    :  16/05/2003
'--------------------------------------------------------------------------------

Dim Sql As String
Dim rs As VSRecordset
    Periodo = Right(Periodo, 4) & Left(Periodo, 2)
    If CDbl(Nvl(Left(Periodo, 4), 0)) >= CDbl(Nvl(Temp.PegaParametro(Bdados, "ANO ISS NOVO"), 0)) Then
        Sql = "SELECT TCI_TAE_CAE,TCI_TIPO_RECOLHIMENTO_ISS + 1 as TCI_TIPO_RECOLHIMENTO_ISS  FROM TAB_CONTRIBUINTE WHERE TCI_IM = '" & Im & "' or tci_cgc_cpf ='" & Im & "'"
        
        If Bdados.AbreTabela(Sql, rs) Then
            Me.Buscar CStr(rs!tci_tae_cae & ""), True
            If rs!TCI_TIPO_RECOLHIMENTO_ISS = 2 Or rs!TCI_TIPO_RECOLHIMENTO_ISS = 6 Then
                Sql = "SELECT TAE_ALIQUOTA_TPPC,TAE_ALIQUOTA_PJ FROM TAB_ATIVIDADE_ECONOMICA WHERE TAE_CAE = " & rs!tci_tae_cae
                If Bdados.AbreTabela(Sql, rs) Then
                    ValorFixo = Nvl("" & rs!TAE_ALIQUOTA_TPPC, 0)
                    BuscaAliquotaAtividade = Nvl("" & rs!TAE_ALIQUOTA_PJ, 0) / 100
                End If
            ElseIf rs!TCI_TIPO_RECOLHIMENTO_ISS = 4 Then
                Sql = "SELECT TAE_ALIQUOTA_SPL,TAE_ALIQUOTA_PJ FROM TAB_ATIVIDADE_ECONOMICA WHERE TAE_CAE = " & rs!tci_tae_cae
                If Bdados.AbreTabela(Sql, rs) Then
                    ValorFixo = Nvl("" & rs!TAE_ALIQUOTA_SPL, 0)
                    BuscaAliquotaAtividade = Nvl("" & rs!TAE_ALIQUOTA_PJ, 0) / 100
                End If
            ElseIf rs!TCI_TIPO_RECOLHIMENTO_ISS = 1 Or rs!TCI_TIPO_RECOLHIMENTO_ISS = 3 Then
                Sql = "SELECT TAE_ALIQUOTA_PJ FROM TAB_ATIVIDADE_ECONOMICA WHERE TAE_CAE = " & rs!tci_tae_cae
                If Bdados.AbreTabela(Sql, rs) Then BuscaAliquotaAtividade = Nvl("" & rs!TAE_ALIQUOTA_PJ, 0) / 100
            End If
        End If
    Else
        BuscaAliquotaAtividade = Imposto.BuscaAliquota(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN)), _
                                                         Nvl(Left(Periodo, 4), 0))
    End If
End Function

Public Sub PreencherCboAtiv(Combo As Object, Optional Grupo As String)
    Dim Sql As String
    Sql = "Select distinct(tae_nome), TAE_CAE, TAE_TGA_COD_GRUPO From Tab_atividade_Economica"
    If Trim(Grupo) <> "" Then
        Sql = Sql & " where TAE_TGA_COD_GRUPO = '" & Grupo & "'"
    End If
    Sql = Sql & " order by tae_nome"
    Combo.Preencher Bdados, Sql
End Sub

Public Sub PreencherCboPoder(Combo As Object)
        Dim Sql As String
    Sql = "Select tap_poder, tap_cod_ativ_poder  From Tab_atividade_Poder order by tap_poder"
    Combo.Preencher Bdados, Sql
End Sub

Public Sub PreencherCboNturJur(Combo As Object)
        Dim Sql As String
    Sql = "SELECT TNJ_NATUREZA,TNJ_COD_NATUREZA FROM TAB_NATUREZA_JURIDICA order by TNJ_NATUREZA"
    Combo.Preencher Bdados, Sql
End Sub

Public Function BuscaFator(atividade As String, Optional RetFator As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "SELECT tae_desc_fator from tab_atividade_economica where tae_nome ='" & atividade & "' and tae_fator =1"
    If Bdados.AbreTabela(Sql, rs) Then
        RetFator = "" & rs!tae_desc_fator
        BuscaFator = True
    End If
End Function

Public Function VerificaSeTemCadastro(CodAtividade As String) As Boolean
    Dim Sql As String
    Dim rs As VSRecordset
    Sql = "select tci_im from tab_contribuinte where tci_tae_cae = '" & CodAtividade & "'"
    If Bdados.AbreTabela(Sql, rs) Then
        VerificaSeTemCadastro = True
    End If
End Function

