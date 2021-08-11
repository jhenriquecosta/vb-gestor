VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Funcionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub PreencheComboFuncionario(Combo As Object)
    Combo.Preencher Bdados, "SELECT TFU_COD_FUNCIONARIO,TFU_NOME FROM TAB_FUNCIONARIO " & _
        " ORDER BY TFU_NOME", 1
End Sub

