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

Private Sub Class_Initialize()
    Set Bdados = CreateObject("VSClass.VSDados")
    Set Edita = CreateObject("VSClass.VSTexto")
    Set Util = CreateObject("VSClass.VSUtil")
    Set Seguranca = CreateObject("VSClass.VSSeguranca")
    Set Instala = CreateObject("VSClass.VSInstala")
    Set Temp = CreateObject("VSClass.VSTemp")
    Set Relatorio = CreateObject("VSClass.VSRelatorio")
End Sub
Public Property Get Banco() As Object
    Set Banco = Bdados.Conexao
End Property

Public Property Set Banco(NovoBanco As Object)
    Set Bdados.Conexao = NovoBanco
End Property

Public Function Abre_Aplicacao(Formulario As String, Modo As Byte, Cod_s As String, S As String, Desc_F As String) As Boolean
    Abre_Aplicacao = True
    Cod_sis = Cod_s: Sistema = S & "  " & App.Major & "." & App.Minor & "." & App.Revision: Desc_Form = Desc_F
    Select Case UCase(Formulario)
'--------------------------------------------------
        Case "CSEG101": CSEG101.Show
'--------------------------------------------------
        Case "CPAR101": CPAR101.Show
        Case "CPAR201"
            With CPAR201
                .cabVISUAL1.Exibir Bdados, Formulario, App.Path
                .rodVISUAL1.Exibir Bdados, Formulario, App.Major, App.Minor, App.Revision
                .Show
            End With
'--------------------------------------------------
        Case "CAPL101": CAPL101.Show
        Case "CAPL102"
            With CAPL102
                .cabVISUAL1.Exibir Bdados, Formulario, App.Path
                .rodVISUAL1.Exibir Bdados, Formulario, App.Major, App.Minor, App.Revision
                .Show
            End With
        Case "CAPL201": CAPL201.Show
        Case "CAPL301": CAPL301.Show
'--------------------------------------------------
        Case "CTRN101":
            With CTRN101
                .cabVISUAL1.Exibir Bdados, Formulario, App.Path
                .rodVISUAL1.Exibir Bdados, Formulario, App.Major, App.Minor, App.Revision
                .Show
            End With
        Case "CTRN102"
            With CTRN102
                .cabVISUAL1.Exibir Bdados, Formulario, App.Path
                .rodVISUAL1.Exibir Bdados, Formulario, App.Major, App.Minor, App.Revision
                .Show
            End With
        Case "CTRN401"
            With CTRN401
                .cabVISUAL1.Exibir Bdados, Formulario, App.Path
                .rodVISUAL1.Exibir Bdados, Formulario, App.Major, App.Minor, App.Revision
                .Show
            End With

        Case Else
            Abre_Aplicacao = False
            Util.informa "Acesso Negado."
    End Select

End Function
