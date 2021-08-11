VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Function Alvara(Bdados As Object, Im As String, Restricoes As String) As String
    With TCOB102
        .Tag = "EXTERNO"
        .txtIm = Im
        .PreencheTela "tci_im = '" & Im & "'", Bdados
        .txtRestricao = Restricoes
        .Show
    End With
End Function
