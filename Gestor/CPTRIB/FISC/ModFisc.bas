Attribute VB_Name = "ModFisc"
Option Explicit

Public Type RedeFiscalizacao
    CodEtapa As Double
    Descricao As String
    CodEtapaPai As Double
    Ordem As Integer
    CodFuncionario As Integer
    CaminhoRpt As String
    TipoEtapa As Integer
    Prazo As Integer
    DataCadastro As String
    Usuario As String
End Type
Public EtapaRede As RedeFiscalizacao
