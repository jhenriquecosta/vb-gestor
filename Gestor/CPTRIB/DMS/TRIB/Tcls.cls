VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public User As String
Public Bdados As New VSDados
Public Edita As New VSTexto
Public Rpt As New VSRelatorio
Public Util As New VSUtil
Public Instala As New VSInstala
Public Seguranca As New VSSeguranca
Public Temp As New VSTemp
Public Aplicacoes As Aplicacoes
'Public Imposto As VsTFuncoes.VSImposto
'Public CodPagamento As Double
Public MUN As String
Public CODMUN As String
'***********************Area de Instanciacoes dos Modulos*******************************
'Public AplicacoesVTFuncoes As New VsTFuncoes.VsTFuncAplicacoes
Public AplicacoesVsTCisp As New VsTCisp.Aplicacoes
'***********************---------------------------------*******************************
Public Const Const_Extrato As String = "EXTRATO"
Public Const Const_Notificacao As String = "NOTIFICA"
Public Const Const_Monetario As String = "#,##0.00"
Public Const Const_ImAvulso As String = "11000000-00"

Private Const cteEspacamentoColunas As Integer = 3
Private Const cteLinhasCabecalho As Integer = 11
Private Const cteLinhasRodape As Integer = 5

Public Sistema As String
Public Desc_Form As String
Public Cod_sis As String

Public Enum enuTipoCampo
    tipTexto = 5
    tipData = 1
    tipInteiro = 2
    tipMoeda = 3
    tipFloat = 4
End Enum
Public Enum enuAlinhamentoCampo
    aliEsquerda = 0
    aliCentro = 1
    aliDireita = 2
End Enum

Public Enum StatusLivroAforamento
    slaLivroAberto = 1
    slaLivroFechado = 0
End Enum

Public Enum SituacaoImovel
    siNaoAforado = 0
    siAforado = 1
    siTransferencia = 2
    siReimpressao = 3
End Enum

Sub Main()
    Set Aplicacoes = New Aplicacoes
'    Set Imposto = New VSImposto
End Sub

Public Function BuscaCodigo(Tabela As String) As Long
    Dim Rs As VSRecordset
    Dim ConsegAbrir As Boolean
    
    BuscaCodigo = 0
    If Bdados.AbreTabela(Tabela, Rs) Then
        BuscaCodigo = "" & Rs(0)
    End If
    Bdados.FechaTabela Rs
    
End Function

Public Function BuscaIndiceCombo(Combo As ComboBox, Tabela As String, CampoCodigo As String, CampoNome As String, Indice As Integer) As Integer
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim i As Integer
    
    Sql = "SELECT " & CampoNome & " from " & Tabela & " where " & CampoCodigo & "=" & Indice
    If Bdados.AbreTabela(Sql, Rs) Then
        For i = 0 To Combo.ListCount - 1
            Combo.ListIndex = i
            If Combo.Text = Rs(0) Then
                BuscaIndiceCombo = i
                Bdados.FechaTabela Rs
                Exit Function
            End If
        Next
    End If
    BuscaIndiceCombo = -1
    Bdados.FechaTabela Rs
End Function

Public Function UltimoDiaDoMes(Data As Date) As Date
    UltimoDiaDoMes = DateAdd("d", -1, "01/" & Mid(DateAdd("m", 1, Data), 4))
End Function

Function PreencheEspaco(Texto, Tamanho As Byte) As String
        PreencheEspaco = Texto & Space(Tamanho - Len(Trim(Texto)))
End Function

Public Function CepCliente() As String
    CepCliente = Temp.PegaParametro(Bdados, "CEP CLIENTE") & "-" & Temp.PegaParametro(Bdados, "COMPLEMENTO CEP CLIENTE")
End Function

Public Function FuncaoReal(Campo As String) As String
    FuncaoReal = "cast(" & Campo & " as decimal(14,2))"
End Function

Public Sub AtualizaUF(Combo As ComboBox)
   Combo.Clear
   Combo.AddItem "MA"
   Combo.AddItem "AC"
   Combo.AddItem "AM"
   Combo.AddItem "AP"
   Combo.AddItem "AL"
   Combo.AddItem "BA"
   Combo.AddItem "CE"
   Combo.AddItem "DF"
   Combo.AddItem "ES"
   Combo.AddItem "GO"
   Combo.AddItem "MG"
   Combo.AddItem "MS"
   Combo.AddItem "MT"
   Combo.AddItem "PA"
   Combo.AddItem "PB"
   Combo.AddItem "PE"
   Combo.AddItem "PI"
   Combo.AddItem "PR"
   Combo.AddItem "SC"
   Combo.AddItem "SE"
   Combo.AddItem "SP"
   Combo.AddItem "RJ"
   Combo.AddItem "RN"
   Combo.AddItem "RO"
   Combo.AddItem "RR"
   Combo.AddItem "RS"
   Combo.AddItem "TO"
End Sub

Public Sub AtualizaCabecalho(Lista As Object, Optional Titulo As String)
    With Lista
        .CabecalhoCliente = Temp.PegaParametro(Bdados, "CLIENTE")
        .CabecalhoDepartamento = Temp.PegaParametro(Bdados, "SETOR")
        .CabecalhoEstado = Temp.PegaParametro(Bdados, "ESTADO")
        .CabecalhoSecretaria = Temp.PegaParametro(Bdados, "SEMFAZ")
        .CabecalhoTitulo = Titulo
        .RodapeUsuario = Aplicacoes.Usuario
    End With
End Sub

Public Function ListIndexDe(Combo As ComboBox, Texto As String) As Integer
    Dim i As Integer
    For i = 0 To Combo.ListCount
        If Combo.List(i) = Texto Then
            ListIndexDe = i
            Exit Function
        End If
    Next
    ListIndexDe = -1
End Function

Public Function NomeDe(Oque As Byte, Codigo As String) As String
    Dim Sql As String
    Sql = "SELECT "
    Select Case Oque
        Case 0
            Sql = Sql & "tco_descricao_componente FROM Tab_Componente WHERE tco_tmu_cod_municipio =" & Aplicacoes.Codigo_Municipio & " and tco_cod_componente = "
            
        Case 1
            Sql = Sql & "tba_nome FROM tab_bairro WHERE tba_tmu_cod_municipio =" & Aplicacoes.Codigo_Municipio & " and TBA_COD_BAIRRO = "
        Case 2
            Sql = Sql & "ttl_nome from tab_tipo_logr where ttl_cod_tip_logr="
    End Select
    
    Sql = Sql & Codigo & IIf(Oque = 1, " and tba_tmu_cod_municipio=" & Aplicacoes.Codigo_Municipio, "")
    
    Dim Rs As VSRecordset
    If Bdados.AbreTabela(Sql, Rs) Then
        NomeDe = Rs(0)
    End If
    Bdados.FechaTabela Rs
    
End Function


