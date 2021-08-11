Attribute VB_Name = "Api"
Option Explicit

Public bRegistrado As Boolean 'Indica se o controle está registrado

Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As _
    Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal _
    hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

' this variable is shared between the two routines below
Dim m_ClientIsInterpreted As Boolean

' return True if the client application of this DLL
' is an interpreted Visual Basic program running in the IDE
'
' NOTE: this code is meant to be inserted in a BAS module
'       inside an ActiveX DLL project

Public Function Programando() As Boolean
    EnumThreadWindows App.ThreadID, AddressOf EnumThreadWindows_CBK, 0
    Programando = m_ClientIsInterpreted
End Function

' this is a callback function that is executed for each
' window in the same thead as the DLL

Private Function EnumThreadWindows_CBK(ByVal hWnd As Long, _
    ByVal lParam As Long) As Boolean
    Dim buffer As String * 512
    Dim length As Long
    Dim windowClass As String

    ' get the class name of this window
    length = GetClassName(hWnd, buffer, Len(buffer))
    windowClass = Left$(buffer, length)
    
    If windowClass = "IDEOwner" Then
        ' this is the main VB IDE window, therefore
        ' the client application is interpreted
        m_ClientIsInterpreted = True
        ' return False to stop evaluation
        EnumThreadWindows_CBK = False
    Else
        ' return True to continue enumeration
        EnumThreadWindows_CBK = True
    End If
End Function

Public Sub ValidaComponente(Componente As String)

    bRegistrado = True
    Exit Sub
    
    On Error GoTo Trata
    
    'Faz validacao de registro se o componente estiver sendo
    'utilizado para desenvolvimento
    If Programando Then
        Dim Reg As Object
        Set Reg = CreateObject("VTRegistro.VTReg")
        bRegistrado = Reg.Verifica(Componente)
    Else
        bRegistrado = True
    End If
    Exit Sub
Trata:
    bRegistrado = False
End Sub
