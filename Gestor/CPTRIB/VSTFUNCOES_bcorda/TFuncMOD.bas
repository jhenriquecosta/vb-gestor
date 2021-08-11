Attribute VB_Name = "TFuncMOD"
Option Explicit
Public User As String
Public Bdados As New VSDados
Public Edita As New VSTexto
Public Rpt As New VSRelatorio
Public Util As New VSUtil
Public Instala As New VSInstala
Public Seguranca As New VSSeguranca
Public Temp As New VSTemp
'Public Aplicacoes As Aplicacoes
Public Imposto As VSImposto
Public Conta As ContaCorrente
Public CalculoIptu As New VSIptu
'Public CodPagamento As Double
Public MUN As String
Public CODMUN As String
Public Sis As String
Public Desc_Form As String
Public AplicacoesVTFuncoes As New VsTFuncoes.VsTFuncAplicacoes
Public Const Const_Monetario As String = "#,##0.00"
Public Const Const_ImAvulso As String = "11000000-00"
Public Const Const_Obrig As String = "63"
Public Const Const_NaoParcelaveis As String = "3,4,6,7,11"
Public Const Const_NaoPagos As String = "2,5,9,10"
Public Const Const_NaoPagosTodos As String = "2,4,5,6,8,9,10"
Public Const Const_Extrato As String = "EXTRATO"
Public Const Const_Notificacao As String = "NOTIFICA"
Public Const Const_AutoInfracao As String = "AUTO"

Public NovaDataVencimento As String
Public SenhaLiberacao As String

Sub Main()
    Set AplicacoesVTFuncoes = New VsTFuncoes.VsTFuncAplicacoes
    Set Imposto = New VSImposto
    Set Conta = New ContaCorrente
End Sub

Public Function BuscaCodigo(Tabela As String) As Long
    Dim Rs As VSRecordset
    Dim ConsegAbrir As Boolean
    
    If Bdados.AbreTabela(Tabela, Rs) Then
        BuscaCodigo = Nvl("" & Rs(0), 0)
    End If
    Bdados.FechaTabela Rs
End Function

Public Function UltimoDiaDoMes(Data As Date) As Date
    UltimoDiaDoMes = DateAdd("d", -1, "01/" & Mid(DateAdd("m", 1, Data), 4))
End Function

Public Function BuscaNaGeral(Tabela As String, Registro As Integer) As String
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "select TGE_NOME from tab_geral where TGE_TIPO = (SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME ='" & Tabela & "' ) and TGE_CODIGO =" & Registro
    BuscaNaGeral = ""
    If Bdados.AbreTabela(Sql, Rs) Then
        BuscaNaGeral = Rs(0)
    End If
    Rs.Fechar
End Function
