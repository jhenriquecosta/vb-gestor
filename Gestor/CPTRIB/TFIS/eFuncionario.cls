VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    
  Public Enum TipoCargo
    tcCargoFiscal = 1
    tcCargo_Auditor = 2
  End Enum
  
  Public Enum TipoLotacao
    tlLotacaoFazenda = 1
    tlLotacaoAdministracao = 2
  End Enum
  
  Private f_Matricula As String
  Private f_Nome As String
  Private f_CPF As String
  Private f_Lotacao As Integer
  Private f_Cargo As Integer


Public Property Get Matricula() As String
    Matricula = f_Matricula
End Property

Public Property Let Matricula(ByVal Value As String)
    f_Matricula = Value
End Property

Public Property Get Nome() As String
    Nome = f_Nome
End Property

Public Property Let CPF(ByVal Value As String)
    f_CPF = Value
End Property



Public Property Get Lotacao() As String
    Lotacao = f_Lotacao
End Property

Public Property Let Lotacao(ByVal Value As String)
    f_Lotacao = Value
End Property

Public Property Get Cargo() As String
    Cargo = f_Cargo
End Property

Public Property Let Cargo(ByVal Value As String)
    f_Cargo = Value
End Property
