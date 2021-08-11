VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eFuncionarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Atributos As New eFuncionario

Public Function CarregaFuncionarios() As Boolean
    Dim Condicao As String
    Dim sql As String
    
    sql = "SELECT * FROM TAB_FUNCIONARIO "
    
    If (Atributos.Matricula <> "") Then
        If (Condicao = "") Then
            Condicao = Condicao & " TFU_MATRICULA = '" & Atributos.Matricula & "'"
        Else
            Condicao = Condicao & " AND TFU_MATRICULA = '" & Atributos.Matricula & "'"
        End If
    End If
    
    If (Atributos.Cargo <> 0) Then
        If (Condicao = "") Then
            Condicao = Condicao & " TFU_COD_CARGO = " & Atributos.Cargo
        Else
            Condicao = Condicao & " AND TFU_COD_CARGO = " & Atributos.Cargo
        End If
    End If
   
    If (Atributos.Lotacao <> 0) Then
        If (Condicao = "") Then
            Condicao = Condicao & " TFU_TLO_COD_LOTACAO = " & Atributos.Lotacao
        Else
            Condicao = Condicao & " AND TFU_TLO_COD_LOTACAO = " & Atributos.Lotacao
        End If
    End If
     
    If Condicao <> "" Then
        sql = sql & " WHERE " & Condicao
    End If
End Function
 
   
