VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Aplicacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Usuario() As String
    Usuario = User
End Property

Public Property Let Usuario(ByVal vNewValue As String)
    User = vNewValue
End Property

Public Property Get Banco() As Object
    Set Banco = BDados.Conexao
End Property

Public Property Set Banco(NovoBanco As Object)
    Set BDados.Conexao = NovoBanco
End Property

Private Sub Class_Initialize()
    Set BDados = New VSClass.VSDados
    Set Edita = New VSClass.VSTexto
    Set Util = New VSClass.VSUtil
    Set Instala = New VSClass.VSInstala
    Set Temp = New VSClass.VSTemp
    Set Seguranca = New VSClass.VSSeguranca

'    Set Bdados = CreateObject("VSClass.VSDados")
'    Set Edita = CreateObject("VSClass.VSTexto")
'    Set Util = CreateObject("VSClass.VSUtil")
'    Set Instala = CreateObject("VSClass.VSInstala")
'    Set Temp = CreateObject("VSClass.VSTemp")
'    Set Seguranca = CreateObject("VSClass.VSSeguranca")
End Sub

Public Function Abre_Aplicacao(Formulario As String, Modo As Byte, Cod_s As String, S As String, Desc_F As String) As Boolean
    Abre_Aplicacao = True
    Cod_sis = Cod_s: Sistema = S & "  " & App.Major & "." & App.Minor & "." & App.Revision: Desc_Form = Desc_F
    Select Case UCase(Formulario)
'--------------------------------------------------
        Case "PCAU101": PCAU101.Show
        Case "PCAU102": PCAU102.Show
        Case "PCAU301":  PCAU301.Show
'--------------------------------------------------
        Case "PMUS101": PMUS101.Show
        Case "PMUS201": PMUS201.Show
'--------------------------------------------------
        Case "PTBS101": PTBS101.Show
        Case "PTBS201": PTBS201.Show
        Case "PTBS301": PTBS301.Show
        Case "PTBS401": PTBS401.Show
        Case "PTBS501": PTBS501.Show
        Case "PTBS502": PTBS502.Show
        Case "PTBS601": PTBS601.Show
        
        Case "PTBS701": PTBS701.Show
'--------------------------------------------------
        Case "PATS101": PATS101.Show
'--------------------------------------------------
        Case "PDEF101": PDEF101.Show

'--------------------------------------------------
        Case "GMEN101": GMEN101.Show
        'Case "GMEN102": GMEN102.Show
        
        Case Else
            Abre_Aplicacao = False
            Util.Informa "Acesso Negado."
    End Select
End Function
