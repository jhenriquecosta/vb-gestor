VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSUtil"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function apiSysDir Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function LogonUser Lib "kernel32" Alias " LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Private Declare Function apiUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Private Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Temp As New VSTemp

Public Type Retorno
    OpcaoBotao As Byte
    Resposta As String
End Type

Public Enum DesligaComo
    Logoff = 0
    Desligar = 1
    Reiniciar = 2
End Enum

Public Sub Pausa(Tempo_Milisegundos As Long)
    If bRegistrado Then Call Sleep(Tempo_Milisegundos)
End Sub
    
Public Function Confirma(Mensagem As String, Optional Titulo As String = "Confirmação") As Boolean
    If bRegistrado Then
        'Cria msgbox para o usuario confirmar uma operacao
        T.OpcaoBotao = 0
        With VSMensagem
            .imgExclama.Visible = True
            .lblTitulo = Titulo
            .lblMsg = Mensagem
            Load .cmdBotao(7)
            .cmdBotao(7).Caption = "&Não"
            .cmdBotao(7).Left = .lblMsg.Left + .lblMsg.Width - .cmdBotao(7).Width
            .cmdBotao(7).Top = .cmdBotao(6).Top
            .cmdBotao(7).Visible = True
            .cmdBotao(6).Caption = "&Sim"
            .cmdBotao(6).Left = .cmdBotao(7).Left - .cmdBotao(6).Width - 30
            .cmdBotao(6).Top = .cmdBotao(7).Top
            .cmdBotao(7).Cancel = True
           .Show 1
            Confirma = IIf(T.OpcaoBotao = 6, True, False)
        End With
    End If
End Function

Public Sub Erro(Mensagem As String)
    If bRegistrado Then
        'Cria msgbox para o usuario informando a existencia de um erro em uma operacao
        VSMensagem.imgErro.Visible = True
        MostraMensagem Mensagem, "Erro"
    End If
End Sub

Public Sub Informa(Mensagem As String)
    If bRegistrado Then
        'Cria uma msgbox contendo uma informacao de Informa ao usuario
        VSMensagem.imgInforma.Visible = True
        MostraMensagem Mensagem, "Aviso"
    End If
End Sub

Public Sub Mensagem(Mensagem As String)
    If bRegistrado Then
        'Cria uma msgbox contendo uma informacao de Informa ao usuario
        VSMensagem.imgMsg.Visible = True
        MostraMensagem Mensagem, "Aviso"
    End If
End Sub

Public Sub Avisa(Mensagem As String)
    If bRegistrado Then
        '    Cria uma msgbox contendo uma informacao de Informa ao usuario
        VSMensagem.imgAvisa.Visible = True
        MostraMensagem Mensagem, "Atenção"
    End If
End Sub

Public Function Entrada(Pergunta As String, Titulo As String, Optional ValorPadrao As String = "") As String
    If bRegistrado Then
        '    Cria uma msgbox contendo uma informacao de Informa ao usuario
        T.Resposta = ""
        With VSMensagem
            .lblMsg.Visible = False
            .txtEntrada.Visible = True
            .imgEntrada.Visible = True
            .lblTitulo = Pergunta
            .Caption = Titulo
            .lblMsg = Pergunta
            .Show 1
        End With
        Entrada = IIf(T.Resposta <> "", T.Resposta, ValorPadrao)
    End If
End Function

Private Sub MostraMensagem(Mensagem As String, Titulo As String)
    If bRegistrado Then
        '   Mostra o conteúdo das janelas de informações ao usuário
        With VSMensagem
            .lblTitulo = Titulo
            .lblMsg = Mensagem
            .Caption = Titulo
            .Show 1
        End With
    End If
End Sub

Public Function Nvl(Conteudo As String, Valor_para_Retorno) As Variant
    If bRegistrado Then Nvl = IIf(Trim(Conteudo) <> "", Conteudo, Valor_para_Retorno)
End Function

Public Function Nnl(Conteudo, Valor_para_Retorno) As Variant
    If bRegistrado Then Nnl = IIf(Not Conteudo Is Null, Conteudo, Valor_para_Retorno)
End Function

Public Sub Negrito(Formulario As Object, Caixa_Inicial As Byte, Caixa_Final As Byte)
    If bRegistrado Then
        Dim i As Integer
        For i = Caixa_Inicial To Caixa_Final
           Formulario.txt(i).FontBold = True
        Next
    End If
End Sub

Public Function ValidaCartao(Num_Cartao As String) As Boolean
    If bRegistrado Then
        'Retorna True se numero do cartao de credito for valido
        Dim Cartao(1 To 16) As Byte 'Recebe os digitos do nº do cartao
        Dim Mult_Dig As Integer ' Valor da Multip. por 2 dos digitos das posicoes impares do nº
        Dim Soma_Total As Integer 'Acumula total da soma de todos os digitos do nº do cartao
        Dim i As Integer
        
        Soma_Total = 0
        Mult_Dig = 0
        If Len(Num_Cartao) < 16 Then 'Verifica se foi digitado 16 numeros
            ValidaCartao = False
            Exit Function
        End If
        For i = 1 To 16 'Carrega o vetor com o nº do cartao
            Cartao(i) = Val(Mid(Num_Cartao, i, 1))
        Next
        For i = 1 To 16
            Select Case i 'Verifica a posicao do digito no numero do cartao
                Case 1, 3, 5, 7, 9, 11, 13, 15
                    Mult_Dig = Cartao(i) * 2
                    If Mult_Dig > 9 Then
                        Mult_Dig = Mult_Dig - 9
                    End If
                    Soma_Total = Soma_Total + Mult_Dig
                Case Else 'Posicoes pares do numero do cartao
                    Soma_Total = Soma_Total + Cartao(i)
            End Select
        Next
        'Soma tem que ser menor que 150 e  multipla de 10
        If Soma_Total < 150 Then
            ValidaCartao = IIf(Soma_Total Mod 10 = 0, True, False)
        Else 'Numero de cartao invalido
            ValidaCartao = False
        End If
    End If
End Function

Public Function ValidaCgc(Cgc As String) As Boolean
    On Error GoTo Err_CGC
    If bRegistrado Then
        Dim strCgc As String 'armazena a parte do CGC entre /0001- ou seja 0001
        Dim Pega_Cgc As String 'armazena do CPF que será utilizada para o cálculo
        Dim Inverte_Cgc As String 'armazena os digitos do CPF da direita para a esquerda
        Dim Dig_Cgc As Integer 'armazena o digito separado para cálculo (uma a um)
        Dim Dig_Cgc_Mult As Integer 'armazena o digito específico multiplicado pela sua base
        Dim Soma_Dig_Cgc_Mult As Long 'armazena a soma dos digitos multiplicados pela sua base(Dig_Cgc_Mult)
        Dim Soma_Dig_Cgc_Mult1 As Long 'armazena a soma dos 8 primeiros digitos multiplicados pela sua base(Dig_Cgc_Mult)
        Dim Soma_Dig_Cgc_Mult2 As Long ''armazena a soma dos 4 ultimos digitos multiplicados pela sua base(Dig_Cgc_Mult)
        Dim Div_Dig_Cgc_Mult As Double 'armazena a divisão dos digitos*base por 11
        Dim Result_Div As Long 'armazena inteiro da divisão
        Dim Resto_Div As Integer 'armazena o resto
        Dim Dig_Ver1 As Integer 'armazena o 1º digito verificador
        Dim Dig_Ver2 As Integer 'armazena o 2º digito verificador
        Dim Dig_Ver As String 'armazena o digito verificador
        Dim i As Integer
        Dim Edita As VSTexto
        
        Soma_Dig_Cgc_Mult = 0
        Soma_Dig_Cgc_Mult1 = 0
        Soma_Dig_Cgc_Mult2 = 0
        Dig_Cgc = 0
        Dig_Cgc_Mult = 0
        'Inicia cálculos do 1º dígito
        'Separa os dígitos do CGC que serão multiplicados de 2 a 9.
        'Retira a "/" da máscara de entrada.
        strCgc = Right(Cgc, 7)
        strCgc = Left(strCgc, 4)
        Pega_Cgc = Left(Cgc, 8)
        Pega_Cgc = Right(Pega_Cgc, 4) & strCgc
        For i = 2 To 9
            Inverte_Cgc = Right(Pega_Cgc, i - 1)
            Dig_Cgc = Left(Inverte_Cgc, 1)
            Dig_Cgc_Mult = Dig_Cgc * i
            Soma_Dig_Cgc_Mult1 = Soma_Dig_Cgc_Mult1 + Dig_Cgc_Mult
        Next i
        'Separa os 4 primeiros dígitos do CGC
        Pega_Cgc = Left(Cgc, 4)
        For i = 2 To 5
            Inverte_Cgc = Right(Pega_Cgc, i - 1)
            Dig_Cgc = Left(Inverte_Cgc, 1)
            Dig_Cgc_Mult = Dig_Cgc * i
            Soma_Dig_Cgc_Mult2 = Soma_Dig_Cgc_Mult2 + Dig_Cgc_Mult
        Next i
        Soma_Dig_Cgc_Mult = Soma_Dig_Cgc_Mult1 + Soma_Dig_Cgc_Mult2
        Div_Dig_Cgc_Mult = Soma_Dig_Cgc_Mult / 11
        Result_Div = Int(Div_Dig_Cgc_Mult) * 11
        Resto_Div = Soma_Dig_Cgc_Mult - Result_Div
        If Resto_Div = 0 Or Resto_Div = 1 Then
            Dig_Ver1 = 0
        Else
            Dig_Ver1 = 11 - Resto_Div
        End If
        Soma_Dig_Cgc_Mult = 0
        Soma_Dig_Cgc_Mult1 = 0
        Soma_Dig_Cgc_Mult2 = 0
        Dig_Cgc = 0
        Dig_Cgc_Mult = 0
        'Inicia cálculos do 2º dígito
        strCgc = Right(Cgc, 7)
        strCgc = Left(strCgc, 4)
        Pega_Cgc = Left(Cgc, 8)
        Pega_Cgc = Right(Pega_Cgc, 3) & strCgc & Dig_Ver1
        For i = 2 To 9
            Inverte_Cgc = Right(Pega_Cgc, i - 1)
            Dig_Cgc = Left(Inverte_Cgc, 1)
            Dig_Cgc_Mult = Dig_Cgc * i
            Soma_Dig_Cgc_Mult1 = Soma_Dig_Cgc_Mult1 + Dig_Cgc_Mult
        Next i
        Pega_Cgc = Left(Cgc, 5)
        For i = 2 To 6
            Inverte_Cgc = Right(Pega_Cgc, i - 1)
            Dig_Cgc = Left(Inverte_Cgc, 1)
            Dig_Cgc_Mult = Dig_Cgc * i
            Soma_Dig_Cgc_Mult2 = Soma_Dig_Cgc_Mult2 + Dig_Cgc_Mult
        Next i
        Soma_Dig_Cgc_Mult = Soma_Dig_Cgc_Mult1 + Soma_Dig_Cgc_Mult2
        Div_Dig_Cgc_Mult = Soma_Dig_Cgc_Mult / 11
        Result_Div = Int(Div_Dig_Cgc_Mult) * 11
        Resto_Div = Soma_Dig_Cgc_Mult - Result_Div
        If Resto_Div = 0 Or Resto_Div = 1 Then
            Dig_Ver2 = 0
        Else
            Dig_Ver2 = 11 - Resto_Div
        End If
        Dig_Ver = Dig_Ver1 & Dig_Ver2
        'Caso o CGC esteja errado dispara a mensagem
        If Dig_Ver <> Right(Cgc, 2) Then
            ValidaCgc = False
            Avisa "CNPJ inválido."
        Else
            ValidaCgc = True
        End If
    End If
    Exit Function
Exit_CGC:
        Exit Function
Err_CGC:
        MsgBox Error$
        Resume Exit_CGC
End Function

Public Function ValidaCpf(Cpf As String) As Boolean
    If bRegistrado Then
        Dim Pega_Cpf As String 'armazena do CPF que será utilizada para o cálculo
        Dim Inverte_Cpf As String 'armazena os digitos do CPF da direita para a esquerda
        Dim Dig_Cpf As Integer 'armazena o digito separado para cálculo (uma a um)
        Dim Dig_Cpf_Mult As Integer 'armazena o digito específico multiplicado pela sua base
        Dim Soma_Dig_Cpf_Mult As Long 'armazena a soma dos digitos multiplicados pela sua base(Dig_Cpf_Mult)
        Dim Div_Dig_Cpf_Mult As Double 'armazena a divisão dos digitos*base por 11
        Dim Result_Div As Long 'armazena inteiro da divisão
        Dim Resto_Div As Integer 'armazena o resto
        Dim Dig_Ver1 As Integer 'armazena o 1º digito verificador
        Dim Dig_Ver2 As Integer 'armazena o 2º digito verificador
        Dim Dig_Ver As String 'armazena o digito verificador
        Dim i As Integer
        On Error GoTo Err_CPF
        
        Soma_Dig_Cpf_Mult = 0
        Dig_Cpf = 0
        Dig_Cpf_Mult = 0
        Pega_Cpf = Left(Cpf, 9)
        
        'Inicia cálculos do 1º dígito
        For i = 2 To 10
            Inverte_Cpf = Right(Pega_Cpf, i - 1)
            Dig_Cpf = Left(Inverte_Cpf, 1)
            Dig_Cpf_Mult = Dig_Cpf * i
            Soma_Dig_Cpf_Mult = Soma_Dig_Cpf_Mult + Dig_Cpf_Mult
        Next i
        Div_Dig_Cpf_Mult = Soma_Dig_Cpf_Mult / 11
        
        Result_Div = Int(Div_Dig_Cpf_Mult) * 11
        Resto_Div = Soma_Dig_Cpf_Mult - Result_Div
        If Resto_Div = 0 Or Resto_Div = 1 Then
            Dig_Ver1 = 0
        Else
            Dig_Ver1 = 11 - Resto_Div
        End If
        
        Pega_Cpf = Pega_Cpf & Dig_Ver1 'concatena o CPF com o primeiro digito verificador
        Soma_Dig_Cpf_Mult = 0
        Dig_Cpf = 0
        Dig_Cpf_Mult = 0
        'Inicia cálculos do 2º dígito
        For i = 2 To 11
            Inverte_Cpf = Right(Pega_Cpf, i - 1)
            Dig_Cpf = Left(Inverte_Cpf, 1)
            Dig_Cpf_Mult = Dig_Cpf * i
            Soma_Dig_Cpf_Mult = Soma_Dig_Cpf_Mult + Dig_Cpf_Mult
        Next i
        Div_Dig_Cpf_Mult = Soma_Dig_Cpf_Mult / 11
        Result_Div = Int(Div_Dig_Cpf_Mult) * 11
        Resto_Div = Soma_Dig_Cpf_Mult - Result_Div
        If Resto_Div = 0 Or Resto_Div = 1 Then
            Dig_Ver2 = 0
        Else
            Dig_Ver2 = 11 - Resto_Div
        End If
        Dig_Ver = Dig_Ver1 & Dig_Ver2
        'Caso o CPF esteja errado dispara a mensagem
        If Dig_Ver <> Right(Cpf, 2) Then
            ValidaCpf = False
            Avisa "CPF inválido."
        Else
            ValidaCpf = True
        End If
    End If
    Exit Function
Exit_CPF:
        Exit Function
Err_CPF:
        MsgBox Error$
        Resume Exit_CPF
End Function

Public Function SystemDir() As String

    If bRegistrado Then
        SystemDir = Space(50)
        If apiSysDir(SystemDir, 50) > 0 Then
            SystemDir = RTrim(SystemDir)
            If InStr(SystemDir, Chr(0)) > 0 Then
                SystemDir = Left(SystemDir, InStr(SystemDir, Chr(0)) - 1)
            End If
            SystemDir = IIf(Right(SystemDir, 1) <> "\", SystemDir & "\", SystemDir)
        Else
            SystemDir = ""
        End If
    End If
End Function

Public Sub DesligaComputador(Tipo As DesligaComo)
    If bRegistrado Then Call ExitWindowsEx(Tipo, 0)
End Sub

Public Function ProcuraArquivo(Unidade As String, Arquivo As String) As String
    If bRegistrado Then
        Dim dir_names() As String
        Dim num_dirs As Integer
        Dim i As Integer
        Dim FileName As String
        Dim new_files As String
        Dim attr As Integer
        On Error Resume Next
        
            
        FileName = Dir(Unidade & "\" & Arquivo, vbNormal)
        Do While FileName <> ""
            new_files = new_files & Unidade & "\" & FileName & vbCrLf
            FileName = Dir()
        Loop
        ProcuraArquivo = ProcuraArquivo & new_files
        
        FileName = Dir(Unidade & "\*.*", vbDirectory)
        Do While FileName <> ""
            attr = 0
            attr = GetAttr(Unidade & "\" & FileName)
            
            If FileName <> "." And FileName <> ".." And (attr And vbDirectory) <> 0 Then
                num_dirs = num_dirs + 1
                ReDim Preserve dir_names(1 To num_dirs)
                dir_names(num_dirs) = FileName
            End If
            FileName = Dir()
        Loop
        
        For i = 1 To num_dirs
            ProcuraArquivo = ProcuraArquivo & ProcuraArquivo(Unidade & "\" & dir_names(i), Arquivo)
        Next i
    End If
End Function

Public Function MapUnidade(Caminho As String, Senha As String, Unidade As String) As Boolean
    If bRegistrado Then
        If InStr(Unidade, ":") = 0 Then Unidade = Unidade & ":"
        MapUnidade = IIf(WNetAddConnection(Caminho, Senha, Unidade) > 0, False, True)
    End If
End Function

Public Sub MudaNomeMaquina(Nome As String)
    If bRegistrado Then Call SetComputerName(Nome)
End Sub

Function PegaIdentidadeComputador() As String
    If bRegistrado Then
        Dim NomeComputador As String
        NomeComputador = Space(256)
        Call GetComputerName(NomeComputador, 256)
        PegaIdentidadeComputador = Left(Trim(NomeComputador), Len(Trim(NomeComputador)) - 1)
    End If
End Function

Public Sub CentralizaForm(Form As Object, Mdi As Object)
    If bRegistrado Then
        Form.Left = (((Mdi.ScaleWidth - Form.Width) \ 2) + Mdi.Left)
        Form.Top = (((Mdi.ScaleHeight - Form.Height) \ 2) + Mdi.Top)
    End If
End Sub
Public Sub HabilitaForm(Form As Object, Valor As Boolean)
    If bRegistrado Then
        Form.Enabled = Valor
        If Not Valor Then
            Screen.MousePointer = vbArrowHourglass
        Else
            Screen.MousePointer = vbNormal
        End If
    End If
End Sub
Public Function ApagarArquivo(Arquivo As String) As Boolean
    If bRegistrado Then
        Dim A As String
        A = Dir(Arquivo)
        If A <> "" Then
            Kill Arquivo
            ApagarArquivo = True
        Else
            ApagarArquivo = False
        End If
    End If
End Function

Public Sub OrdenaGrid(Grid As Object, Coluna As Object)
    If bRegistrado Then
        Grid.Sorted = True
        If Grid.SortKey = Coluna.Index - 1 Then
            Grid.SortOrder = Abs(Grid.SortOrder - 1)
        Else
            Grid.SortOrder = 0
            Grid.SortKey = Coluna.Index - 1
        End If
    End If
End Sub
Public Function MontaGrid(Bdados As Object, Grid As Object, Sql As String, ParamArray Tamanho_Colunas()) As Boolean
    If bRegistrado Then
        Dim RS As VSRecordset
        Dim ItmX As Object
        Dim i As Integer
    
        Grid.AllowColumnReorder = True
        Grid.Arrange = 2 'lvwAutoTop
        Grid.Gridlines = True
        Grid.LabelEdit = 1 'lvwManual
        Grid.View = 3 'lvwReport
        Grid.FullRowSelect = True
        Grid.HotTracking = True
        Grid.FlatScrollBar = False
        Grid.HideSelection = False
        Grid.LabelWrap = False
        Grid.ListItems.Clear
        Grid.ColumnHeaders.Clear
    
        If Trim(Sql) = "" Then
            Grid.ListItems.Clear
        Else
            If Bdados.AbreTabela(Sql, RS) Then
                For i = 0 To RS.Fields.Count - 1
                    If i <= UBound(Tamanho_Colunas()) Then
                        Grid.ColumnHeaders.Add , , RS.Fields(i).Name, Tamanho_Colunas(i)
                    Else
                        Grid.ColumnHeaders.Add , , RS.Fields(i).Name, (Grid.Width / RS.Fields.Count)
                    End If
                Next
                Do Until RS.Eof
                    Set ItmX = Grid.ListItems.Add(, , CStr("" & RS(0)))
                    For i = 1 To RS.Fields.Count - 1
                        If Not IsNull(RS(i)) Then
                            ItmX.SubItems(i) = CStr("" & RS(i))
                        End If
                    Next
                    RS.MoveNext
                Loop
                MontaGrid = True
            End If
        End If
        Bdados.FechaTabela RS
    End If
End Function

Public Sub Marcar(Marc As Object, Fixo As Object, Valor As Boolean)
    On Error GoTo Trata
    If bRegistrado Then
        
        Marc.Checked = Valor
        If Marc.Children > 0 Then
            Marcar Marc.Child, Marc, Valor
        End If
        If Not Marc.Parent Is Nothing Then
            If Marc.Parent.Children > 1 And Not Marc.Next Is Nothing Then
                If Marc <> Fixo Then Marcar Marc.Next, Marc, Valor
            End If
        End If
    End If
Trata:
    If Err.Number <> 0 Then
        Avisa Err.Description
        Screen.MousePointer = 0
    End If
End Sub

Public Sub Integridade(No As Object, Optional Subindo As Boolean)
    On Error GoTo Trata
    
    If bRegistrado Then
        If No.Children > 0 And Not Subindo Then
            Integridade No.Child
        Else
            If No.Next Is Nothing Then
                If Not No.Parent Is Nothing Then
                    No.Parent.Checked = MarcarPai(No.Parent.Child)
                    If Not No.Parent.Next Is Nothing Then
                        Integridade No.Parent.Next
                    Else
                        Integridade No.Parent, True
                    End If
                End If
            Else
                Integridade No.Next
            End If
        End If
    End If
Trata:
    If Err.Number <> 0 Then
        Avisa Err.Description
        Screen.MousePointer = 0
    End If
End Sub

Public Function MarcarPai(No As Object) As Boolean
    On Error GoTo Trata
    
    If bRegistrado Then
        If No.Checked Then
            MarcarPai = True
            Exit Function
        Else
            If Not No.Next Is Nothing Then
                MarcarPai = MarcarPai(No.Next)
            End If
        End If
    End If
Trata:
    If Err.Number <> 0 Then
        Avisa Err.Description
        Screen.MousePointer = 0
    End If
End Function

Public Sub CarregaFig(Img As Object, Cam As String, TipoMin As Byte, TipoMax As Byte, Tam As Byte)
    If bRegistrado Then
        On Error Resume Next
        'Tipo = 4-7 p/ Menu
        Dim Arq As String
        Dim T As String
        
        Img.ListImages.Clear
        Img.ImageWidth = Tam
        Img.ImageHeight = Tam
        
        Arq = Dir(Cam & "\Imagens\*.ico", vbNormal)
        Do While Arq <> ""
            T = Mid(Arq, 1, Len(Arq) - 4)
            If Len(T) >= TipoMin And Len(T) <= TipoMax Then
                Img.ListImages.Add , T, LoadPicture(Cam & "\Imagens\" & Arq)
            End If
            Arq = Dir()
        Loop
        If Img.ListImages.Count = 0 Then
            Img.ListImages.Add , "NAOTEM", LoadPicture(Cam & "\NAOTEM.ico")
        End If
    End If
End Sub

Public Function UltimoDiaDoMes(Data As Date) As Date
    If bRegistrado Then UltimoDiaDoMes = DateAdd("d", -1, "01/" & Mid(DateAdd("m", 1, Data), 4))
End Function

Public Function PrimeiroDiaDoMes(Data As Date) As Date
    If bRegistrado Then PrimeiroDiaDoMes = CDate("01/" & Month(Data) & "/" & Year(Data))
End Function

Public Sub AtualizaCount(Label As Object, Grid As Object)
    If bRegistrado Then
        Select Case Grid.ListItems.Count
            Case 0
                Label.Caption = "Nenhum Registro"
            Case 1
                Label.Caption = "1 Registro"
            Case Else
                Label.Caption = Grid.ListItems.Count & " Registros"
        End Select
    End If
End Sub

Public Sub HabilitarGuias(ActiveTab As Object, ParamArray NumerosGuias())
    If bRegistrado Then
        Dim numguia As Integer
            For numguia = 1 To ActiveTab.Tabs.Count
                ActiveTab.Tabs(numguia).Enabled = Contem(numguia, NumerosGuias())
            Next numguia
            ActiveTab.SelectedTab = NumerosGuias(LBound(NumerosGuias))
    End If
End Sub

Public Sub ExibirGuias(ActiveTab As Object, ParamArray NumerosGuias())
    If bRegistrado Then
        Dim numguia As Integer
            For numguia = 1 To ActiveTab.Tabs.Count
                ActiveTab.Tabs(numguia).Visible = Contem(numguia, NumerosGuias())
            Next numguia
    End If
End Sub

Public Function Contem(Elemento, ParamArray Vetor()) As Boolean
    If bRegistrado Then
        Dim item As Integer
        For item = LBound(Vetor(0)) To UBound(Vetor(0))
            If Vetor(0)(item) = Elemento Then
                Contem = True
                Exit For
            End If
        Next
    End If
End Function

Public Function ParseString(ByVal vsString As String, ByVal vsDelimiter As String, ByVal viNumber As Integer)
' VSClass.VSUtil.Function ParseString
'================================================================================
' Queiroz em VTDES_01
' 24/05/2002-11:24:25
'
' Descricao  : Considerando uma string com delimitadores como um vetor, ParseString
'               retorna a substring que esta na posicao viNumber
'
' Parametros : vsString (String) - String com delimitadores
'              vsDelimiter (String) - Delimitador da string
'              viNumber (Integer) - Posicao no "vetor". Comeca em 1
'
' Ex: ParseString("Maranhao#Piaui#Ceara#Para", "#", 3) = "Ceara"
'--------------------------------------------------------------------------------
    On Error GoTo ErroParsing
    If bRegistrado Then
        Dim PosBusca As Long, PosAnterior As Long
        Dim Elemento As Integer
        Dim Valor As String
        
        If Trim(vsString) = "" Then Exit Function
        'If Trim(vsDelimiter) = "" Then Exit Function
        If viNumber <= 0 Then Exit Function
        
        PosBusca = -1 * Len(vsDelimiter) + 1
        Valor = ""
        Elemento = 0
        
        Do
            PosAnterior = PosBusca + Len(vsDelimiter)
            PosBusca = InStr(PosAnterior, vsString, vsDelimiter)
            Elemento = Elemento + 1
        Loop Until (Elemento = viNumber) Or (PosBusca = 0)
        
        If Elemento = viNumber Then
            If PosBusca = 0 Then
                Valor = Mid(vsString, PosAnterior)
            Else
                Valor = Mid(vsString, PosAnterior, PosBusca - 1)
                PosBusca = InStr(1, Valor, vsDelimiter)
                If PosBusca > 0 Then Valor = Mid(Valor, 1, PosBusca - 1)
            End If
        End If
        
        ParseString = Valor
    End If
    Exit Function
ErroParsing:
    Erro Err.Description
End Function

Private Sub Class_Initialize()
    ValidaComponente "CLASS"
End Sub

Public Function PosicaoLista(Valor As String, Lista As String, Delimitador As String) As Integer
    Dim i As Integer, item As String
    
    i = 1
    Do
        item = ParseString(Lista, Delimitador, i)
        If item = Valor Then
            PosicaoLista = i
            Exit Do
        End If
        i = i + 1
    Loop While item <> ""
End Function

