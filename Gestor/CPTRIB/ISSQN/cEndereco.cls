VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEndereco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Endereco As String
Private m_Bairro As String
Private m_Municipio As String
Private m_UF As String
Private m_CEP As String

Public Property Get CEP() As String
    CEP = m_CEP
End Property

Public Property Let CEP(ByVal Value As String)
    m_CEP = Value
End Property

Public Property Get UF() As String
    UF = m_UF
End Property

Public Property Let UF(ByVal Value As String)
    m_UF = Value
End Property

Public Property Get Municipio() As String
    Municipio = m_Municipio
End Property

Public Property Let Municipio(ByVal Value As String)
    m_Municipio = Value
End Property

Public Property Get Bairro() As String
    Bairro = m_Bairro
End Property

Public Property Let Bairro(ByVal Value As String)
    m_Bairro = Value
End Property

Public Property Get Endereco() As String
    Endereco = m_Endereco
End Property

Public Property Let Endereco(ByVal Value As String)
    m_Endereco = Value
End Property
