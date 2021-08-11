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
'    Descrição   : Classe de administração de atividades econômicas
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
' Parâmetros   :  CodigoAtividade (String)
' Descrição    :  Busca os dados da atividade econômica informada
' Criado por   :  Éderson Robson Bragagnolo
' Maquina      :  Balsas - Planejamento02
' Data         :  14/05/2003
' Alteração    :  14/05/2003
'--------------------------------------------------------------------------------
    Dim sql As String
    Dim rs As VSRecordset
    On Error Resume Next
        
    sql = "SELECT " & IIf(Detalhes, "*", "TAE_CAE") & " FROM TAB_ATIVIDADE_ECONOMICA WHERE TAE_CAE = '" & CodigoAtividade & "'"
    If Bdados.AbreTabela(sql, rs) Then
        m_Codigo = rs!TAE_CAE
        If Detalhes Then
            m_Nome = "" & rs!tae_nome
            m_GrupoCodigo = Nvl("" & rs!TAE_TGA_COD_GRUPO, 0)
            m_Valor = "" & rs!TAE_VALOR
            m_FatorCodigo = Nvl("" & rs!TAE_FATOR, 0)
            m_FatorDescricao = "" & rs!TAE_DESC_FATOR
            m_Estimativo = Nvl("" & rs!TAE_ESTIMATIVO, 0)
            m_RamoCodigo = Nvl("" & rs!TAE_TRA_COD_RAMO, 0)
            m_AliquotaTPPC = Nvl("" & rs!TAE_ALIQUOTA_TPPC, 0)
            m_AliquotaSPL = Nvl("" & rs!TAE_ALIQUOTA_SPL, 0)
            m_AliquotaPJ = Nvl("" & rs!TAE_ALIQUOTA_PJ, 0)
        End If
        Buscar = True
    End If
End Function
Public Function Gravar() As Boolean
'--------------------------------------------------------------------------------
' Projeto      :  VTTrib
' Procedimento :  Gravar
' Descrição    :  Gravar os dados a respeito da atividade econômica
' Criado por   :  Éderson Robson Bragagnolo
' Maquina      :  Balsas - Planejamento02
' Data         :  14/05/2003
' Alteração    :  14/05/2003
'--------------------------------------------------------------------------------

    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
        
    Campos = "TAE_CAE, TAE_NOME, TAE_TGA_COD_GRUPO, TAE_VALOR, TAE_FATOR, " & _
        " TAE_DESC_FATOR, TAE_ESTIMATIVO, TAE_TRA_COD_RAMO, " & _
        " TAE_ALIQUOTA_TPPC, TAE_ALIQUOTA_SPL, TAE_ALIQUOTA_PJ "
    Valores = Bdados.PreparaValor(m_Codigo, m_Nome, m_GrupoCodigo, m_Valor, _
        m_FatorCodigo, m_FatorDescricao, m_Estimativo, m_RamoCodigo, _
        m_AliquotaTPPC, m_AliquotaSPL, m_AliquotaPJ)
    Condicao = "TAE_CAE = '" & m_Codigo & "'"
    If Bdados.GravaDados("TAB_ATIVIDADE_ECONOMICA", Valores, Campos, Condicao) Then
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
' Parâmetros   :  CodigoAtividade (String)
' Descrição    :  Exclusão da atividade econômica informada
' Criado por   :  Éderson Robson Bragagnolo
' Maquina      :  Balsas - Planejamento02
' Data         :  14/05/2003
' Alteração    :  14/05/2003
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
' Parâmetros   :  Grid (Object)
' Descrição    :  Monta uma grid com o conteúdo de uma query especificada
' Criado por   :  Éderson Robson Bragagnolo
' Maquina      :  Balsas - Planejamento02
' Data         :  14/05/2003
' Alteração    :  14/05/2003
'--------------------------------------------------------------------------------
    Dim sql As String
    Dim TamanhoColunas As String
    
    sql = "SELECT TAE_CAE AS Código, TAE_NOME AS Atividade, " & _
            " TGA_NOME AS Grupo,  " & Bdados.Converte("TAE_VALOR", TCDuplo) & " AS [Valor(R$)], " & _
            " TAE_DESC_FATOR AS Fator, TGE_NOME AS Estimado, " & _
            " TAE_ALIQUOTA_TPPC AS TPPC, TAE_ALIQUOTA_SPL AS SPL, TAE_ALIQUOTA_PJ AS [PJ(%)]" & _
            " FROM TAB_ATIVIDADE_ECONOMICA, TAB_GRUPO_ATIVIDADE, VIS_SIM_NAO" & _
            " WHERE TAE_TGA_COD_GRUPO = TGA_COD_GRUPO AND " & _
            " TGE_CODIGO = TAE_ESTIMATIVO "
    If GrupoAtividade <> "" Then
        sql = sql & " and TAE_TGA_COD_GRUPO = '" & GrupoAtividade & "'"
    End If
    If CodAtividade <> "" Then
        sql = sql & " and TAE_CAE = '" & CodAtividade & "'"
    End If
    If Estimativo <> "" Then
        sql = sql & " and TAE_ESTIMATIVO = '" & Estimativo & "'"
    End If
    If Nome <> "" Then
        sql = sql & " and (tae_nome like '" & Nome & "%' or tae_nome like '% " & Nome & "%')"
    End If
     sql = sql & " ORDER BY TAE_CAE"
    TamanhoColunas = "1200, 3100, 1800, 1000, 800, 900"
    If Grid.Preencher(Bdados, sql, TamanhoColunas) Then PreencheGrid = True
    
End Function

Public Function PreencheCombo(Combo As Object, Informacao As InformacaoAtividade) As Boolean
'--------------------------------------------------------------------------------
' Projeto      :  VTTrib
' Procedimento :  PreencheCombo
' Parâmetros   :  Grid (Object)
' Descrição    :  Monta um combobox com o conteúdo do enum informado
' Criado por   :  Éderson Robson Bragagnolo
' Maquina      :  Balsas - Planejamento02
' Data         :  15/05/2003
' Alteração    :  15/05/2003
'--------------------------------------------------------------------------------
    Dim sql As String
    Select Case Informacao
        Case iaRamo: sql = "SELECT TRA_NOME_RAMO, TRA_COD_RAMO FROM TAB_RAMO_ATIVIDADE ORDER BY TRA_NOME_RAMO"
        Case iaGrupoAtividade: sql = "SELECT TGA_NOME, TGA_COD_GRUPO FROM TAB_GRUPO_ATIVIDADE"
    End Select
    
    If Combo.Preencher(Bdados, sql) Then PreencheCombo = True
    
End Function

Public Function BuscaAliquotaAtividade(Bdados As Object, Im As String, Optional ByRef ValorFixo As Double) As Double
'--------------------------------------------------------------------------------
' Projeto      :  VTTrib
' Procedimento :  BuscaAliquotaAtividade
' Parâmetros   :  IM (String), ValorFixo (Double)
' Descrição    :  Localiza e retorna o valor fixo ou alíquota correspondente a atividade do contribuinte
' Criado por   :  Éderson Robson Bragagnolo
' Maquina      :  Balsas - Cobranca02
' Data         :  16/05/2003
' Alteração    :  16/05/2003
'--------------------------------------------------------------------------------

Dim sql As String
Dim rs As VSRecordset
        
    sql = "SELECT TCI_TAE_CAE,TCI_TIPO_RECOLHIMENTO_ISS FROM TAB_CONTRIBUINTE WHERE TCI_IM = '" & Im & "'"
    If Bdados.AbreTabela(sql, rs) Then
        Me.Buscar rs!tci_tae_cae, True
        If rs!tci_tipo_recolhimento_iss = 2 Then
            sql = "SELECT TAE_ALIQUOTA_TPPC,TAE_ALIQUOTA_PJ FROM TAB_ATIVIDADE_ECONOMICA WHERE TAE_CAE = " & rs!tci_tae_cae
            If Bdados.AbreTabela(sql, rs) Then
                ValorFixo = Nvl("" & rs!TAE_ALIQUOTA_TPPC, 0)
                BuscaAliquotaAtividade = Nvl("" & rs!TAE_ALIQUOTA_PJ, 0) / 100
            End If
        ElseIf rs!tci_tipo_recolhimento_iss = 4 Then
            sql = "SELECT TAE_ALIQUOTA_SPL,TAE_ALIQUOTA_PJ FROM TAB_ATIVIDADE_ECONOMICA WHERE TAE_CAE = " & rs!tci_tae_cae
            If Bdados.AbreTabela(sql, rs) Then
                ValorFixo = Nvl("" & rs!TAE_ALIQUOTA_SPL, 0)
                BuscaAliquotaAtividade = Nvl("" & rs!TAE_ALIQUOTA_PJ, 0) / 100
            End If
        ElseIf rs!tci_tipo_recolhimento_iss = 1 Then
            sql = "SELECT TAE_ALIQUOTA_PJ FROM TAB_ATIVIDADE_ECONOMICA WHERE TAE_CAE = " & rs!tci_tae_cae
            If Bdados.AbreTabela(sql, rs) Then BuscaAliquotaAtividade = Nvl("" & rs!TAE_ALIQUOTA_PJ, 0) / 100
        End If
    End If
End Function

