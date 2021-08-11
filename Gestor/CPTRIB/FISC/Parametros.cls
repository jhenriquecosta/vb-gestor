VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public vCodigo As String
Public vParametro As String
Public vDescricao As String

Public Function BuscaCodigo(Parametro As String) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "Select TFT_CODIGO from TAB_FISCALIZACAO_TEXTO_PADRAO where TFT_PARAMETRO ='" & Parametro & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaCodigo = Rs!TFT_CODIGO
    End If
End Function

Public Function Gravar() As Boolean
    On Error GoTo Trata
    Dim Valores As String
    Dim Campos As String
    Dim Conta As New ContaCorrente
    If vCodigo = "" Then
        vCodigo = Conta.GeraCodPagamento("43")
    End If
    Campos = "TFT_CODIGO,TFT_PARAMETRO,TFT_DESCRICAO"
    Valores = Bdados.PreparaValor(vCodigo, vParametro, vDescricao)
    
    Gravar = Bdados.GravaDados("TAB_FISCALIZACAO_TEXTO_PADRAO", Valores, Campos, _
    "TFT_CODIGO = " & vCodigo)

    Exit Function
Trata:
    If Err.Number <> 0 Then
        Util.Avisa "Erro: " & Err.Number & " - " & Err.Description & "."
        Screen.MousePointer = 0
    End If
End Function

Public Function PreencheGrid(Grd As Object) As Boolean
    PreencheGrid = Grd.Preencher(Bdados, "SELECT TFT_CODIGO,TFT_PARAMETRO AS Parmâmetro, TFT_DESCRICAO AS Descrição FROM TAB_FISCALIZACAO_TEXTO_PADRAO", 0, 2500, 4300)
End Function

Public Function PreencheCombo(Cbo As Object) As Boolean
    PreencheCombo = Cbo.Preencher(Bdados, "SELECT TFT_CODIGO,TFT_PARAMETRO FROM TAB_FISCALIZACAO_TEXTO_PADRAO", 1)
End Function

Public Function Deletar(Codigo As String) As Boolean
    Deletar = Bdados.DeletaDados("TAB_FISCALIZACAO_TEXTO_PADRAO", "TFT_CODIGO=" & Codigo)
End Function

Public Function CarregaDadosParametro(CodParametro As String) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    LimpaDadosParametro
    Sql = "SELECT * FROM TAB_FISCALIZACAO_TEXTO_PADRAO WHERE TFT_CODIGO =" & CodParametro
    If Bdados.AbreTabela(Sql, Rs) Then
        vCodigo = Nvl("" & Rs!TFT_CODIGO, 0)
        vDescricao = "" & Rs!TFT_DESCRICAO
        vParametro = "" & Rs!TFT_PARAMETRO
    End If
End Function
Public Sub LimpaDadosParametro()
    vCodigo = ""
    vDescricao = ""
    vParametro = ""
End Sub
