VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eImovel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function BuscarImovel(Ic As String, Optional ByRef TipoLogr As Object, Optional ByRef Logr As Object, Optional ByRef Numero As Object, _
                        Optional ByRef Comp As Object, Optional ByRef Bairro As Object, Optional ByRef Cep As Object, Optional ByRef Cidade As Object, _
                        Optional ByRef Uf As Object) As Boolean
    Dim Sql As String
    Dim Rs As VSRecordset
    Sql = "Select * from Vis_Imovel Where tim_Ic = '" & Ic & "'" & _
            " AND TBA_TMU_COD_MUNICIPIO = " & Aplicacoes.Codigo_Municipio & " AND tlg_tmu_cod_municipio = " & Aplicacoes.Codigo_Municipio
    If Bdados.AbreTabela(Sql, Rs) Then
        TipoLogr = "" & Rs!ttl_nome
        Logr = "" & Rs!tlg_nome
        Numero = "" & Rs!tim_numero
        Comp = "" & Rs!tim_complemento
        Bairro = "" & Rs!tba_nome
        Cep = "" & Rs!tim_cep
        Cidade = Aplicacoes.Municipio
        Uf = "MA"
        BuscarImovel = True
    End If
End Function
