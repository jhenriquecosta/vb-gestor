Attribute VB_Name = "MMOD"
Option Explicit

Public Usuario As String
Public BdSis As Object
'Public BdSeg As Object
Public Edita As Object
Public Instala As Object
Public Seguranca As Object
'Public Temp As Object
Public VSis As Object
Private Tipo As Long
Private Dsn As String
Private Pass As String
Private Cat As String
Private ArqBanco As String
Public Lotacoes As String

Public User As String
Public Bdados As Object
Public Temp As Object
Public Util As Object
Public Aplicacoes As Aplicacoes
Public MUN As String
Public CODMUN As String

Public Sistema As String
Public Desc_Form As String
Public Cod_sis As String


Public Sub ConectaAplicacao()
    
End Sub

Sub Main()
    Dim SQL As String
    Dim RStemp As Object
    Set BdSis = CreateObject("VSClass.VSDados")
    Set BdSis = CreateObject("VSClass.VSDados")
    Set Edita = CreateObject("VSClass.VSTexto")
    Set Util = CreateObject("VSClass.VSUtil")
    Set Seguranca = CreateObject("VSClass.VSSeguranca")
    Set Instala = CreateObject("VSClass.VSInstala")
    Set Temp = CreateObject("VSClass.VSTemp")
    Set Aplicacoes = New Aplicacoes
    If Not BdSis.AbreBanco(0, App.Path & "\" & "CTDMS.MDB", "admin") Then
        Call Util.Erro("O Sistema não pode ser iniciado. Problemas com o Banco de Dados do Sistema.")
'        End
    End If
    
    
    Set VSis = CreateObject("VSTRIB.Aplicacoes")
    
    Set VSis.Banco = Nothing
    Set VSis.Banco = BdSis.Conexao
    VSis.Usuario = User
    VSis.Codigo_Municipio = Temp.PegaParametro(BdSis, "MUNICIPIO")
'------------------------------------------------------------------
    SQL = "SELECT TMU_NOME FROM TAB_MUNICIPIO WHERE TMU_COD_MUNICIPIO=" & VSis.Codigo_Municipio
    If BdSis.AbreTabela(SQL, RStemp) Then
        VSis.Municipio = RStemp(0)
    End If
    BdSis.FechaTabela RStemp
    
    ChamaAplicacao "GMNU101", "", "", ""
    DoEvents
End Sub

Public Sub ChamaAplicacao(Formulario As String, Cod_sis As String, Sis As String, Desc As String)
    On Error GoTo Trata
    
    Screen.MousePointer = 11
    If Trim(Cod_sis) <> "" Then Set VSis = CreateObject("VS" & Cod_sis & ".Aplicacoes")
    Select Case UCase(Formulario)
        Case "GMNU101": GMNU101.Show
        Case Else
            Set VSis.Banco = Nothing
            Set VSis.Banco = BdSis.Conexao
            VSis.Abre_Aplicacao Formulario, 0, Cod_sis, Sis, Desc
    End Select
    Screen.MousePointer = 0
    DoEvents
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        If Err.Number = 91 Then
            Util.Informa "Módulo " & Sistema & " não encontrado."
        Else
            Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        End If
        Screen.MousePointer = 0
    End If
End Sub

Public Sub Conecta(ByVal User As String, Sis As String)
    On Error GoTo Trata
    Dim RS As Object
    Dim RStemp As Object
    Dim CodSis As String
    Dim SQL As String
    
'IMAGEM DO RELATORIO-----------------------------------------------
    Dim Arq As String
    Arq = App.Path & "\imagens\" & Sis & ".bmp"
    
    If Dir(Arq) <> "" Then
        FileCopy Arq, "C:\LogoSis.bmp"
    End If
'------------------------------------------------------------------
POSERRO:
'------------------------------------------------------------------
    If Instala.PegaConfig(ArqBanco, Sis, 0) <> "" Then
        Tipo = Instala.PegaConfig(ArqBanco, Sis, 0)
        Dsn = Instala.PegaConfig(ArqBanco, Sis, 1)
        User = Instala.PegaConfig(ArqBanco, Sis, 2)
        Cat = Instala.PegaConfig(ArqBanco, Sis, 3)
        Pass = Instala.PegaConfig(ArqBanco, Sis, 4)
        If Not BdSis.AbreBanco(Tipo, Dsn, User, Pass, Cat) Then
            Call Util.Erro("O Sistema não pode ser iniciado. Problemas com o Banco de Dados do Sistema.")
'            Unload Me
        End If
        Set VSis.Banco = Nothing
        Set VSis.Banco = BdSis.Conexao
    End If
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        If Err.Number = 429 Then
            Util.Informa "Módulo " & Sistema & " não encontrado."
        ElseIf Err.Number = 438 Then
            Resume POSERRO
        Else
            Util.AVISA "Erro: " & Err.Number & " - " & Err.Description & "."
        End If
        Screen.MousePointer = 0
    End If
End Sub
