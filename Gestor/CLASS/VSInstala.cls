VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSInstala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const ArqAtualizacao As String = "\Winft.fta"
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Util As New VSUtil
Private Seguranca As New VSSeguranca
Private Temp As New VSTemp
Public Enum TipoConfig
    tcTipo = 5
    tcDsn = 6
    tcUser = 7
    tcCatalog = 8
    tcPassword = 9
End Enum

Public Function Expirou(Bdados As Object) As Boolean
    If bRegistrado Then
        Dim Rs2 As VSRecordset
        Dim Sql As String
        Sql = "SELECT * FROM TAB_ATUALIZACAO"
        If Bdados.AbreTabela(Sql, Rs2) Then
            If ControleOK(Bdados, IIf(Rs2(0), 1, 0), Rs2(1), Rs2(2), Rs2(3)) Then
                If Rs2(0) = False Then
                    If CLng(DateDiff("d", CDate(Rs2(1)), Date)) < CLng(Rs2(2)) Then
                        Expirou = False
                    Else
                        Call AtualizaExpiro(Bdados, "1", Rs2(1), Rs2(2), PegaSerialDisco(Bdados))
                        Expirou = True
                    End If
                Else
                    Expirou = True
                End If
            Else
                Expirou = True
            End If
        Else
            Expirou = True
        End If
        Bdados.FechaTabela Rs2
    End If
End Function

Public Sub AtualizaExpiro(Bdados As Object, StatusDoExpiro As String, Data As String, Duracao As String, Serial As String)
    If bRegistrado Then
        Dim Crc As String
        Dim RSTEMP As VSRecordset
        Dim Sql As String
        Crc = GeraControle(StatusDoExpiro, Data, Duracao, Serial)
        Sql = "SELECT * FROM TAB_ATUALIZACAO"
        If Bdados.AbreTabela(Sql, RSTEMP) Then
            Call Bdados.AtualizaDados("TAB_ATUALIZACAO", Bdados.PreparaValor(StatusDoExpiro, Bdados.Converte(Data, tctexto), Duracao, Crc, Bdados.Converte(Serial, tctexto)), "TAT_EXPIROU, TAT_DAT_ATUALIZADO, TAT_DURACAO, TAT_CONTROLE, TAT_SERIAL")
        Else
            Call Bdados.InsereDados("TAB_ATUALIZACAO", Bdados.PreparaValor(StatusDoExpiro, Bdados.Converte(Data, tctexto), Duracao, Crc, Bdados.Converte(Serial, tctexto)), "TAT_EXPIROU, TAT_DAT_ATUALIZADO, TAT_DURACAO, TAT_CONTROLE, TAT_SERIAL")
        End If
        Bdados.FechaTabela RSTEMP
    End If
End Sub

Private Function ControleOK(Bdados As Object, Expirou As String, Data As String, Duracao As String, Controle As String) As Boolean
    If bRegistrado Then ControleOK = (GeraControle(Expirou, Data, Duracao, PegaSerialDisco(Bdados)) = Controle)
End Function

Public Function GeraControle(Expirou As String, Data As String, Duracao As String, Serial As String) As String
    If bRegistrado Then
        Dim Resp As String
        Dim A As String
        Dim B As String
        Dim C As String
        Dim D As String
        Dim E As String
        Dim F As String
        Dim X1 As String
        Dim X2 As String
        Dim X3 As String
        Dim X4 As String
        A = Expirou
        B = Left(Data, 2)
        C = Mid(Data, 4, 2)
        D = Right(Data, 4)
        E = Duracao
        F = Serial
        X1 = CStr((2 * CInt(A)) + CLng(B) + CLng(C) + CLng(D) + CLng(E) + CLng(F))
        X2 = CStr((CInt(B) * CInt(B)) + CLng(X1))
        X3 = CStr((CInt(C) * (CInt(C) - CInt(A)) * (CInt(C) - (2 * CInt(A)))) + CLng(X2))
        X4 = CStr(CLng(X3) + CLng(X1) - CInt(A) - CLng(F))
        Resp = CStr(Hex(CLng(X1)) & Hex(CLng(X2)) & Hex(CLng(X3)) & Hex(CLng(X4)))
        GeraControle = Resp
    End If
End Function

Public Function AchouAtualizacao(Bdados As Object, Caminho As String) As Boolean
    On Error GoTo Trata
    If bRegistrado Then
        Dim dat As String
        Dim Dur As String
        Dim Serial As String
        If Dir(Caminho & ArqAtualizacao) = "" Then
            AchouAtualizacao = False
            Exit Function
        Else
            If PegaDadosDeAtualizacao(Bdados, Caminho, ArqAtualizacao, dat, Dur, Serial) Then
                AtualizaExpiro Bdados, "0", dat, Dur, Serial
                Util.ApagarArquivo Caminho & ArqAtualizacao
                AchouAtualizacao = True
            Else
                AchouAtualizacao = False
            End If
        End If
    End If
Trata:
    If Err.Number = 68 Then
        Resume Next
    ElseIf Err.Number = 52 Then
        Resume Next
    ElseIf Err.Number <> 0 Then
        Util.Erro Err.Description
        Screen.MousePointer = 0
    End If
End Function

Private Function PegaDadosDeAtualizacao(Bdados As Object, Cam As String, ByVal Arquivo As String, ByRef DataDeAtualizacao As String, ByRef Duracao As String, ByRef Serial As String) As Boolean
    If bRegistrado Then
        Dim fileFile As Integer
        Dim Linha As String
        Dim OK As Boolean
        Dim i As Integer
        
        fileFile = FreeFile(0)
        If Dir(Cam & Arquivo) <> "" Then
            Open Cam & Arquivo For Input As fileFile
        Else
            Exit Function
        End If
        
        Line Input #fileFile, Linha
        If Linha = "FUTURO" Then
            Line Input #fileFile, Linha
            Linha = Seguranca.DesCriptografa(Linha)
            If IsNumeric(Linha) Then
                If ConfereSerialArq(Bdados, Linha) Then
                    Serial = Linha
                    Line Input #fileFile, Linha
                    Linha = Seguranca.DesCriptografa(Linha)
                    If IsDate(Linha) Then
                        DataDeAtualizacao = Linha
                        Line Input #fileFile, Linha
                        Linha = Seguranca.DesCriptografa(Linha)
                        If IsNumeric(Linha) Then
                            Duracao = Linha
                            PegaDadosDeAtualizacao = True
                        Else
                            PegaDadosDeAtualizacao = False
                        End If
                    Else
                        PegaDadosDeAtualizacao = False
                    End If
                Else
                    PegaDadosDeAtualizacao = False
                End If
            Else
                PegaDadosDeAtualizacao = False
            End If
        Else
            PegaDadosDeAtualizacao = False
        End If
        Close fileFile
    End If
End Function

Public Function GeraArquivoAtualizador(DataDeAtualizacao As String, Duracao As String, Serial As String, Mun As String, Cam As String) As Boolean
    ' RETIRAR PARAMETRO MUN, MAS QUEBRARA COMPATIBILIDADE...
    
    On Error GoTo Trata
    If bRegistrado Then
        Dim fileFile As Integer
        Dim Linha As String
        Dim OK As Boolean
        Dim i As Integer
        Dim J As Integer
        Dim K As Integer
        
        GeraArquivoAtualizador = False
        
        fileFile = FreeFile(0)
    
        Open Cam & ArqAtualizacao For Output As fileFile
    
        Linha = "FUTURO"
        Print #fileFile, Linha
        
        Linha = Format(Serial, "00000")
        Linha = Seguranca.Criptografa(Linha)
        Print #fileFile, Linha
        
        Linha = DataDeAtualizacao
        Linha = Seguranca.Criptografa(Linha)
        Print #fileFile, Linha
        
        Linha = Format(Duracao, "000")
        Linha = Seguranca.Criptografa(Linha)
        Print #fileFile, Linha

    
        For i = 1 To 10
            Linha = ""
            K = Rnd * 10
            For J = 1 To K
                Linha = Linha & Chr(CInt(250 * Rnd))
            Next
            Print #fileFile, Linha
        Next
        Linha = "FUTURO."
        Print #fileFile, Linha
        
        Close fileFile
        GeraArquivoAtualizador = True
    End If
    Exit Function
Trata:
    If Err.Number = 71 Then
        Util.Avisa "Disco não encontrado."
        GeraArquivoAtualizador = False
        Close fileFile
    ElseIf Err.Number <> 0 Then
        Util.Erro Err.Description
        Screen.MousePointer = 0
        GeraArquivoAtualizador = False
        Close fileFile
    End If
End Function

Public Function PegaSerialDisco(Bdados As Object) As String
    If bRegistrado Then
        Dim Sql As String
        Dim RsAux As VSRecordset
        Sql = "SELECT * FROM TAB_ATUALIZACAO"
        If Bdados.AbreTabela(Sql, RsAux) Then
            PegaSerialDisco = RsAux(4)
        Else
            PegaSerialDisco = 0
        End If
        
        Bdados.FechaTabela RsAux
    End If
End Function

Private Function ConfereSerialArq(Bdados As Object, Serial As String) As Boolean
    If bRegistrado Then
        If Serial > Format(PegaSerialDisco(Bdados), "00000") Then
            ConfereSerialArq = True
        Else
            ConfereSerialArq = False
        End If
    End If
End Function

Public Sub NovoPerfil(ByRef Formulario As Object, ByRef Cabecalho As Object, Cod_Sistema As String, Sistema As String, Desc_Formulario As String, Caminho As String)
    If bRegistrado Then
        Formulario.Caption = Formulario.Name
        Cabecalho.EXIBE Sistema, Desc_Formulario, Caminho & Cod_Sistema & ".gif"
    End If
End Sub

Public Sub Perfil(Formulario As Object, lblForm As Object, lblModulo As Object, Usuario As String, Sistema As String, Desc_Formulario As String)
    If bRegistrado Then
        Formulario.Caption = "Futuro Tecnologia - " & Sistema
        Formulario.lblForm.Caption = Desc_Formulario
        Formulario.lblModulo = Formulario.Name
        Formulario.lblForm.ToolTipText = "Usuário: " & Usuario
    End If
End Sub

Public Function PegaConfig(ArqivoConfig As String, Banco As String, Parametro As TipoConfig) As String
    If bRegistrado Then
        Dim Val As String
        Dim B As String
        'TipoBancot
        'dsnd
        'useru
        'catalogoc
        'senhap
        If Dir(ArqivoConfig) <> "" Then
            Open ArqivoConfig For Input As #1
                Do Until Eof(1)
                    Line Input #1, Val
                    Select Case Right(Val, 1)
                        Case "t" 'APLICACAO
                            B = Seguranca.DesCriptografa(Mid(Val, 2, Len(Val) - 2))
                            If Banco = B Then
                                PegaConfig = Left(Val, 1)
                                If Parametro = tcTipo Then Close #1: Exit Function
                            End If
                        Case "d" 'SERVIDOR
                            If Parametro = tcDsn Then
                                If Banco = B Then
                                    PegaConfig = Left(Val, Len(Val) - 1)
                                    Close #1
                                    Exit Function
                                End If
                            End If
                        Case "u" 'USUARIO
                            If Parametro = tcUser Then
                                If Banco = B Then
                                    PegaConfig = Seguranca.DesCriptografa(Left(Val, Len(Val) - 1))
                                    Close #1
                                    Exit Function
                                End If
                            End If
                        Case "c" 'SENHA
                            If Parametro = tcCatalog Then
                                If Banco = B Then
                                    PegaConfig = Seguranca.DesCriptografa(Left(Val, Len(Val) - 1))
                                    Close #1
                                    Exit Function
                                End If
                            End If
                        Case "p" 'BANCO
                            If Parametro = tcPassword Then
                                If Banco = B Then
                                    PegaConfig = Seguranca.DesCriptografa(Left(Val, Len(Val) - 1))
                                    Close #1
                                    Exit Function
                                End If
                            End If
                    End Select
                Loop
            Close #1
        Else
            Util.Avisa "Arquivo '" & ArqivoConfig & "' não encontrado."
        End If
        
        PegaConfig = ""
    End If
End Function

Friend Function PegaConfiguracao(Sistema As String, Parametro As TipoConfig) As String
    If bRegistrado Then
        Dim Linha As String
        Dim LinhaSistema As String
        'TipoBancot
        'dsnd
        'useru
        'catalogoc
        'senhap
        Dim ArqivoConfiguracao As String
        ArqivoConfiguracao = App.Path & "\Conexao.dbc"
        If Dir(ArqivoConfiguracao) <> "" Then
            Open ArqivoConfiguracao For Input As #1
                Do Until Eof(1)
                    Line Input #1, Linha
                    Select Case Right(Linha, 1)
                        Case "A" 'SISTEMA(APLICACAO)
                            LinhaSistema = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Left(Linha, Len(Linha) - 1))))
                            'If Sistema <> LinhaSistema Then Close #1: Exit Function
                        Case "D" 'SGBD
                            If Parametro = tcTipo Then
                                If Sistema = LinhaSistema Then
                                    PegaConfiguracao = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Left(Linha, Len(Linha) - 1))))
                                    If Parametro = tcTipo Then Close #1: Exit Function
                                End If
                            End If
                        Case "S" 'SERVIDOR
                            If Parametro = tcDsn Then
                                If Sistema = LinhaSistema Then
                                    PegaConfiguracao = Left(Linha, Len(Linha) - 1)
                                    Close #1
                                    Exit Function
                                End If
                            End If
                        Case "B" 'BANCO
'                            If Sistema = LinhaSistema Then
'                                If Parametro = tcCatalog Then
'                                    PegaConfiguracao = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Left(Linha, Len(Linha) - 1))))
'                                    Close #1
'                                    Exit Function
'                                End If
'                            End If
                        
                        Case "U" 'USUARIO
                            If Parametro = tcUser Then
                                If Sistema = LinhaSistema Then
                                    PegaConfiguracao = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Left(Linha, Len(Linha) - 1))))
                                    Close #1
                                    Exit Function
                                End If
                            End If
                        Case "P" 'SENHA
                            If Parametro = tcPassword Then
                                If Sistema = LinhaSistema Then
                                    PegaConfiguracao = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Left(Linha, Len(Linha) - 1))))
                                    Close #1
                                    Exit Function
                                End If
                            End If
                        Case "C" 'CATALOGO
                            If Parametro = tcCatalog Then
                                If Sistema = LinhaSistema Then
                                    PegaConfiguracao = Seguranca.DesCriptografa(Seguranca.DesCriptografa(Seguranca.DesCriptografa(Left(Linha, Len(Linha) - 1))))
                                    Close #1
                                    Exit Function
                                End If
                            End If
                    End Select
                Loop
            Close #1
        Else
            Util.Avisa "Arquivo de Configuração não encontrado."
        End If
        
        PegaConfiguracao = ""
    End If
End Function

Public Function GeraConfig(Tipo As String, Sistema As String, Dsn As String, User As String, Password As String, Catalog As String) As String
    If bRegistrado Then
        'TipoBancot
        'dsnd
        'useru
        'catalogoc
        'senhap
        GeraConfig = Tipo & Seguranca.Criptografa(Sistema) & "t"
        GeraConfig = GeraConfig & vbCrLf & Dsn & "d"
        If User <> "" Then GeraConfig = GeraConfig & vbCrLf & Seguranca.Criptografa(User) & "u"
        If Password <> "" Then GeraConfig = GeraConfig & vbCrLf & Seguranca.Criptografa(Password) & "p"
        If Catalog <> "" Then GeraConfig = GeraConfig & vbCrLf & Seguranca.Criptografa(Catalog) & "c"
    End If
End Function

Private Sub Class_Initialize()
    ValidaComponente "CLASS"
End Sub
