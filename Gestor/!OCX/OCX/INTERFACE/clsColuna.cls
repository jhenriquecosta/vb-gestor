VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColuna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private strNome As String
Private enuTipo As enuTipoCampo
Private varMin As Double
Private varMax As Double
Private dblMedia As Double
Private dblSoma As Double
Private m_Tamanho As Integer
Private m_Width As Long

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Let Width(ByVal Value As Long)
    m_Width = Value
End Property

Public Property Get Tamanho() As Integer
    Tamanho = m_Tamanho
End Property

Public Property Let Tamanho(ByVal Value As Integer)
    If Value > m_Tamanho Then m_Tamanho = Value
End Property

Public Property Get Nome() As String
    Nome = strNome
End Property

Friend Property Let Nome(ByVal vnewvalue As String)
    strNome = vnewvalue
End Property

Public Property Get Tipo() As enuTipoCampo
    Tipo = enuTipo
End Property

Public Property Let Tipo(ByVal vnewvalue As enuTipoCampo)
    enuTipo = vnewvalue
End Property

Public Property Get Min() As Double
    Min = varMin
End Property

Friend Property Let Min(ByVal vnewvalue As Double)
    If vnewvalue < varMin Or varMin = 0 Then varMin = vnewvalue
End Property

Public Property Get Max() As Double
    Max = varMax
End Property

Friend Property Let Max(ByVal vnewvalue As Double)
    If vnewvalue > varMax Then varMax = vnewvalue
End Property

Public Property Get Media() As Double
    Media = dblMedia
End Property

Friend Property Let Media(Valor As Double)
    dblMedia = Valor
End Property
Public Property Get Soma() As Double
    Soma = dblSoma
End Property

Friend Property Let Soma(ByVal vnewvalue As Double)
    dblSoma = dblSoma + vnewvalue
End Property

Private Sub Class_Initialize()
    m_Tamanho = 0
End Sub

