VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VsAidf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum TipoNumero
    Credenciamento = 22
    Aidf = 23
End Enum
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

