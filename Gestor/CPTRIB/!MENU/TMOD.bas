Attribute VB_Name = "TMOD"
Option Explicit
Public User As String
Public Bdados As Object
Public Temp As Object
Public Util As Object
Public Aplicacoes As Aplicacoes
Public MUN As String
Public CODMUN As String

Public Sistema As String
Public Desc_Form As String
Public Cod_sis As String

Sub Main()
    Set Aplicacoes = New Aplicacoes
    Set Bdados = CreateObject("VSClass.VSDados")
    Set Temp = CreateObject("VSClass.VSTemp")
    Set Util = CreateObject("VSClass.VSUtil")
End Sub
