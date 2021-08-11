VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSIptuVerde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub CalculaValorVenalTerreno()
    'VVT = valor m² terreno X área do terreno X fator situação X fator topografia X
    'fator pedologia
    CalculaAreaTributavel
    ValorBasicoUnitario = CalculaVBU(Logr, Secao) 'Valor Basico Unitario
    Lote.ValorVenalTerreno = ValorBasicoUnitario * AreaTributavelTerreno * Lote.FatorSituacao * _
                            Lote.FatorTopografia * Lote.FatorPedologia
                
                
    DadosCalculo.ValorM2Terreno = ValorBasicoUnitario
    DadosCalculo.AreaTributavelTerreno = AreaTributavelTerreno
    DadosCalculo.FatorTopografia = Lote.FatorTopografia
    DadosCalculo.FatorPedologia = Lote.FatorPedologia
    DadosCalculo.FatorSituacao = Lote.FatorSituacao
    DadosCalculo.ValorTerreno = Lote.ValorVenalTerreno
    DadosCalculo.AreaLote = Lote.AreaLote
    DadosCalculo.TestadaPrincipal = Lote.TestadaLote
End Sub


