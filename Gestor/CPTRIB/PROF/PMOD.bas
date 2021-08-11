Attribute VB_Name = "PMOD"
Option Explicit

Public BDados As VSDados
Public Edita As VSTexto
Public Util As VSUtil
Public Instala As VSInstala
Public Temp As VSTemp
Public Relatorio As VSRelatorio
Public Seguranca As VSSeguranca

Public Lotacoes As String

'Public BDados As Object
'Public Edita As Object
'Public Util As Object
'Public Instala As Object
'Public Temp As Object
'Public Seguranca As Object

Public Aplicacoes As Aplicacoes
Public User As String

Public Sistema As String
Public Desc_Form As String
Public Cod_sis As String

Public Enum TipoTabela
    Municipio
    Gerencia
End Enum

Sub Main()
    Set Aplicacoes = New Aplicacoes
End Sub

Public Function CodigoDe(Oque As TipoTabela, Nome As String) As String
    Dim Sql As String
    Sql = "SELECT "
    Select Case Oque
        Case Municipio
            Sql = Sql & "TMU_COD_MUNICIPIO FROM TAB_MUNICIPIO WHERE TMU_NOME = '"
        Case Gerencia
    Sql = Sql & "TGR_COD_GERENCIA FROM TAB_GERENCIA WHERE TGR_NOME = '"
    End Select
    Sql = Sql & Nome & "'"

    Dim RS As Object
    If BDados.AbreTabela(Sql, RS) Then CodigoDe = RS(0)
    BDados.FechaTabela RS

End Function
