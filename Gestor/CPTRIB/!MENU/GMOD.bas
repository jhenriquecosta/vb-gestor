Attribute VB_Name = "GMOD"
Option Explicit

Public Usuario As String
Public Sistema As String
Public BdSis As Object
Public BdSeg As Object
Public Edita As Object
Public Util As Object
Public Instala As Object
Public Seguranca As Object
Public Temp As Object
Public VSis As Object
Private Tipo As Long
Private Dsn As String
Private User As String
Private Pass As String
Private Cat As String
Private ArqBanco As String
Public Lotacoes As String

Sub Main()
    
    GMNU000.Show
    DoEvents
    
    Set BdSis = CreateObject("VSClass.VSDados")
    Set BdSeg = CreateObject("VSClass.VSDados")
    Set Edita = CreateObject("VSClass.VSTexto")
    Set Util = CreateObject("VSClass.VSUtil")
    Set Seguranca = CreateObject("VSClass.VSSeguranca")
    Set Instala = CreateObject("VSClass.VSInstala")
    Set Temp = CreateObject("VSClass.VSTemp")
    
    ArqBanco = App.Path & "\Banco.cfg"
    
    If Instala.PegaConfig(ArqBanco, "SEG", 0) <> "" Then
        Tipo = Instala.PegaConfig(ArqBanco, "SEG", 0) 'tcTipo = 0
        Dsn = Instala.PegaConfig(ArqBanco, "SEG", 1) 'tcDsn = 1
        User = Instala.PegaConfig(ArqBanco, "SEG", 2) 'tcUser = 2
        Cat = Instala.PegaConfig(ArqBanco, "SEG", 3) 'tcCatalog = 3
        Pass = Instala.PegaConfig(ArqBanco, "SEG", 4) 'tcPassword = 4
        
        DoEvents
        If Not BdSeg.AbreBanco(Tipo, Dsn, User, Pass, Cat) Then
            Call Util.Erro("Banco de Dados de Segurança não foi encontrado.")
            End
        Else
            AcertaData
            If UCase("" & Temp.PegaParametro(BdSeg, "CONTA")) <> "LIVRE" Then
                If Instala.Expirou(BdSeg) Then
                    If Not (Instala.AchouAtualizacao(BdSeg, "C:") Or Instala.AchouAtualizacao(BdSeg, "A:")) Then
                        Util.AVISA "Não foi possível iniciar o sistema. Chame o suporte técnico."
                        End
                    Else
                        ChamaAplicacao "GMNU101", "", "", ""
                    End If
                Else
                    ChamaAplicacao "GMNU101", "", "", ""
                End If
            Else
                ChamaAplicacao "GMNU101", "", "", ""
            End If
        End If
        
        Unload GMNU000
    
    Else
        Call Util.Erro("Arquivo de configuração não foi encontrado.")
        End
    End If
    

End Sub

Public Sub ChamaAplicacao(Formulario As String, Cod_Sis As String, Sis As String, Desc As String)
    On Error GoTo Trata
    
    Screen.MousePointer = 11
    
    Select Case UCase(Formulario)
'--------------------------------------------------
        Case "GMNU101": GMNU101.Show
'--------------------------------------------------
        Case Else: VSis.Abre_Aplicacao Formulario, 0, Cod_Sis, Sis, Desc
        
        
    End Select
    
    Screen.MousePointer = 0
    
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
    Sistema = Sis
    Set VSis = Nothing
    Set VSis = CreateObject("VS" & Sis & ".Aplicacoes")
    
    Set VSis.Banco = Nothing
    Set VSis.Banco = BdSeg.Conexao
    VSis.Usuario = User
    VSis.CODIGO_MUNICIPIO = Temp.PegaParametro(BdSeg, "MUNICIPIO")
'------------------------------------------------------------------
    SQL = "SELECT TMU_NOME FROM TAB_MUNICIPIO WHERE TMU_COD_MUNICIPIO=" & VSis.CODIGO_MUNICIPIO
    If BdSeg.AbreTabela(SQL, RStemp) Then
        VSis.MUNICIPIO = RStemp(0)
    End If
    BdSeg.FechaTabela RStemp
'------------------------------------------------------------------
    VSis.Lotacoes = VSis.AcessosUsuario(VSis.Usuario)
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
            End
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

Private Sub AcertaData()
    Dim RSDATA As Object, strConsulta As String
    
    
    Select Case BdSeg.Conexao.FormatoBanco
        Case 0 'access
        Case 1 'sql server
            strConsulta = "select day(getdate()) as DIA , month(getdate()) AS MES , year(getdate()) AS ANO, DATEPART(hour, GETDATE()) AS HORA, DATEPART(minute, GETDATE()) AS MINUTO, DATEPART(second, GETDATE()) AS SEGUNDO"
        Case 2 'oracle
            strConsulta = "select to_char(sysdate,'dd') as DIA , to_char(sysdate,'mm') AS MES , to_char(sysdate,'yyyy') AS ANO, to_char(sysdate,'hh24') AS HORA, to_char(sysdate,'mi') AS MINUTO, to_char(sysdate,'ss') AS SEGUNDO from DUAL"
        Case 3 'interbase
    End Select
    If strConsulta <> "" Then
        If BdSeg.AbreTabela(strConsulta, RSDATA) Then
            Date = RSDATA(0) & "/" & RSDATA(1) & "/" & RSDATA(2)
            Time = RSDATA(3) & ":" & RSDATA(4) & ":" & RSDATA(5)
            DoEvents
        End If
    End If
End Sub
