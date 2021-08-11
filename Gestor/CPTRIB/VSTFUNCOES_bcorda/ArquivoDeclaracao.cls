VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArquivoDeclaracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Reg As String
Private DrvSelec  As String
Private diretorio As String
Private Versao As String
Private Registro As String
Private soma_hasharq As Long
Private num_inscricao As String
Private periodo_inicial As String
Private tipo_decla As Integer
Private data_geracao As String
Private HeaderArquivo As String
Private TraillerArquivo As String
Public LinhaAtual As String
Private NumLinhas As Double
Private nArquivo As Integer
Private Const FormatoValor As String = "00000000000000"

Public Municipio As String

Public DataArquivo As String
Public TotalDeclaracoes As Double
Public TotalAceitas As Double
Public TotalRejeitadas As Double

Public Function AbreArquivo(CaminhoArquivo As String)
    nArquivo = FreeFile(0)
    CaminhoArquivo = IIf(Trim(CaminhoArquivo) = "", "C:", CaminhoArquivo)
    Open CaminhoArquivo & "\" & "DMS.DEC" For Output As nArquivo
    NumLinhas = 0
    AbreArquivo = GravaLinha(nArquivo, CriaHeader)
End Function

Public Function CarregaDeclaracao(CaminhoArquivo As String, Optional Lista As Object) As Boolean
    On Error GoTo trata
    Dim Dec As New cDeclaracao
    Dim Declaracao As Object
    Dim GravaDeclaracao As Boolean
    Dim ItemApuracao As cItemDeclaracao
    Dim NotaFiscal As NotaFiscal
    Dim TempData As String
    Dim QtDec As Double
    Dim CgcCpf As Boolean
    Dim Obrig As New Obrigacao
    Dim RsInscricao As VSRecordset
    
    nArquivo = FreeFile(0)
    TotalAceitas = 0
    TotalDeclaracoes = 0
    TotalRejeitadas = 0
    Close
    GravaDeclaracao = False
    CarregaDeclaracao = False
    QtDec = 0
    If Not ValidaArquivo(CaminhoArquivo) Then
        Avisa "Arquivo de declaracão com problema. Verifique o arquivo de log C:\LogValida.log."
        Exit Function
    End If
    Open CaminhoArquivo For Input As nArquivo
    If Not Lista Is Nothing Then Lista.ListItems.Clear
    Do While Not EOF(nArquivo)
        Line Input #nArquivo, LinhaAtual
        Select Case Left(LinhaAtual, 1)
            Case 0
                Municipio = Mid(LinhaAtual, 18, Len(Mid(LinhaAtual, 18)) - 4)
                DataArquivo = Mid(LinhaAtual, 6, 8)
                DataArquivo = Right(DataArquivo, 2) & "/" & Mid(DataArquivo, 3, 2) & "/" & Left(DataArquivo, 4)
            Case 3
                If GravaDeclaracao Then GravaDec Dec, Declaracao
                QtDec = QtDec + 1
                GravaDeclaracao = True
                '******************SECAO DE DEFINICAO DA INSCRICAO******************
                Dec.IM = CDbl(Mid(LinhaAtual, 7, 15))
                CgcCpf = False
                If Len(Dec.IM) = 14 Then
                    Dec.IM = Left(Dec.IM, 2) & "." & Mid(Dec.IM, 3, 3) & "." & _
                            Mid(Dec.IM, 6, 3) & "/" & Mid(Dec.IM, 9, 4) & "-" & Right(Dec.IM, 2)
                    CgcCpf = True
                End If
                If Len(Dec.IM) = 11 Then
                    Dec.IM = Left(Dec.IM, 3) & "." & Mid(Dec.IM, 4, 3) & "." & _
                            Mid(Dec.IM, 7, 3) & "-" & Right(Dec.IM, 2)
                    CgcCpf = True
                End If
                If Len(Dec.IM) = 10 Then Dec.IM = Imposto.FormataInscricao(Dec.IM, InscContrib)
                Set RsInscricao = Obrig.SelecionaContribuinte(Dec.IM, InscContrib)
                If Not RsInscricao.EOF Then Dec.IM = RsInscricao(0)
                '******************FIM DA SECAO******************
                Dec.Periodo = Mid(LinhaAtual, 22, 6)
                Dec.Periodo = Right(Dec.Periodo, 2) & Left(Dec.Periodo, 4)
                TempData = Mid(LinhaAtual, 30, 8)
                Dec.Data = Right(TempData, 2) & "/" & Mid(TempData, 3, 2) & "/" & Left(TempData, 4)
                Dec.Origem = orgInternet
                Dec.Recepcao = Format(Date, "DD/MM/YYYY")
                Dec.Tipo = CDbl(Mid(LinhaAtual, 28, 2))
                Dec.Versao = CDbl(Mid(LinhaAtual, 5, 2))
                Dec.Status = decFinalizada
                DoEvents
                If Not Lista Is Nothing Then
                    Set Declaracao = Lista.ListItems.Add(, , Dec.IM)
                    Declaracao.SubItems(1) = Dec.Periodo
                    Declaracao.SubItems(2) = Dec.Tipo
                    Declaracao.SubItems(3) = Dec.Data
                    Declaracao.SubItems(4) = Dec.Versao
                End If
                DoEvents
            Case 4
                Set ItemApuracao = New cItemDeclaracao
                ItemApuracao.Numero = CDbl(Mid(LinhaAtual, 2, 3))
                ItemApuracao.Valor = CDbl(Mid(LinhaAtual, 5, 15)) / 100
                Dec.Itens.Adicionar ItemApuracao
            Case 5
                Set NotaFiscal = New NotaFiscal
                NotaFiscal.Numero = CDbl(Mid(LinhaAtual, 19, 6))
                NotaFiscal.Destinatario = CDbl(Mid(LinhaAtual, 2, 15))
                NotaFiscal.TipoOperacao = CDbl(Mid(LinhaAtual, 17, 2))
                TempData = Mid(LinhaAtual, 25, 8)
                NotaFiscal.Data = Right(TempData, 2) & "/" & Mid(TempData, 3, 2) & "/" & Left(TempData, 4)
                NotaFiscal.Status = CDbl(Mid(LinhaAtual, 33, 2))
                
                NotaFiscal.ValorTotal = CDbl(Mid(LinhaAtual, 35, Len(FormatoValor))) / 100
                NotaFiscal.BaseCalculo = CDbl(Mid(LinhaAtual, 35 + Len(FormatoValor), Len(FormatoValor))) / 100
                NotaFiscal.ValorMaterialICMS = CDbl(Mid(LinhaAtual, 35 + (Len(FormatoValor) * 2), Len(FormatoValor))) / 100
                NotaFiscal.ImpostoDevido = CDbl(Mid(LinhaAtual, 35 + (Len(FormatoValor) * 3), Len(FormatoValor))) / 100
                NotaFiscal.ImpostoRetido = CDbl(Mid(LinhaAtual, 35 + (Len(FormatoValor) * 4), Len(FormatoValor))) / 100
                Dec.Notas.Adicionar NotaFiscal
            Case 6
                If GravaDeclaracao Then GravaDec Dec, Declaracao
                GravaDeclaracao = False
            Case 9
                If GravaDeclaracao Then GravaDec Dec, Declaracao
                CarregaDeclaracao = True
        End Select
    Loop
    TotalDeclaracoes = QtDec
    Close
    Exit Function
trata:
    Avisa Err.Description
    Resume Next
End Function

Public Function GravaDetalhes(NumDeclaracao As String) As Boolean
    On Error GoTo errogeracao
    GravaDetalhes = CriaDetalhes(nArquivo, NumDeclaracao)
Exit Function
errogeracao:
  Close
  MsgBox error & " - Não foi possível gerar o arquivo."
End Function

Public Function FechaArquivo()
    FechaArquivo = GravaLinha(nArquivo, CriaTrailler)
    Close
End Function

Private Function CalculaCRC(Linha As String) As String
    Dim Soma As Double
    Dim I As Integer
    
    Soma = 0
    For I = 1 To Len(Linha)
        Soma = Soma + (Asc(Mid(Linha, I, 1)) * 2 ^ (I Mod 8))
    Next
    CalculaCRC = Format$(Soma Mod 10000, "0000")
End Function

Private Function GravaLinha(nArquivo As Integer, Optional Linha As String) As Boolean
    On Error GoTo ErroLinha
    Print #nArquivo, IIf(Len(Trim(Linha)) = 0, LinhaAtual, Linha)
    GravaLinha = True
    NumLinhas = NumLinhas + 1
    Exit Function
ErroLinha:
    Avisa "Erro ao gravar linha no arquivo aberto."
    Close
    GravaLinha = False
End Function

Private Function CriaHeader() As String
    
    HeaderArquivo = "0" & "CIAP" & Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00")
    HeaderArquivo = HeaderArquivo & Format(AplicacoesVTFuncoes.Codigo_Municipio, "0000") & _
    UCase(Mid(AplicacoesVTFuncoes.Municipio, 1, IIf(Len(AplicacoesVTFuncoes.Municipio) > 20, 20, Len(AplicacoesVTFuncoes.Municipio))))
    HeaderArquivo = HeaderArquivo & CalculaCRC(HeaderArquivo)
    LinhaAtual = HeaderArquivo
    CriaHeader = HeaderArquivo
End Function

Private Function CriaTrailler() As String
    TraillerArquivo = "9" & Format(NumLinhas + 1, "000000")
    TraillerArquivo = TraillerArquivo & CalculaCRC(TraillerArquivo)
    LinhaAtual = TraillerArquivo
    CriaTrailler = TraillerArquivo
End Function

Private Function CriaDetalhes(nArquivo As Integer, NumDeclaracao As String) As Boolean
    On Error GoTo TrataErro
    Dim Declaracao As New cDeclaracao
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim RsContrib As VSRecordset
    Dim I As Integer
    
    '3. Linha de Declaracao
    '4. Linha de Item de Declaracao
    '5. Linha de Cadastro de Contribuinte
    Declaracao.Numero = NumDeclaracao
    Sql = "SELECT TDC_TCI_IM,TDC_PERIODO,TDC_TIPO_DEC,TDC_VERSAO,TDC_DATA_DECLARACAO FROM TAB_DECLARACAO_CONTRIBUINTE WHERE TDC_NUM_DECLARACAO =" & NumDeclaracao
    If Not Bdados.AbreTabela(Sql, Rs) Then
        Exit Function
    Else
        Sql = "SELECT TCI_NOME FROM TAB_CONTRIBUINTE WHERE TCI_IM  ='" & Rs!TDC_TCI_IM & "' OR TCI_CGC_CPF ='" & Rs!TDC_TCI_IM & "'"
        If Not Bdados.AbreTabela(Sql, RsContrib) Then Exit Function
    End If
    '3.
    LinhaAtual = "3" & "DMS" & Format(Trim(Rs!TDC_VERSAO), "00") & Format(TiraTudo(Rs!TDC_TCI_IM), "000000000000000") & Rs!TDC_PERIODO & _
                Format(Rs!TDC_TIPO_DEC, "00") & Format(Rs!TDC_DATA_DECLARACAO, "YYYYMMDD") & RsContrib!tci_nome
    LinhaAtual = LinhaAtual & CalculaCRC(LinhaAtual)
    GravaLinha nArquivo, LinhaAtual
    '4.
    Declaracao.PreencherItens
    For I = 1 To Declaracao.Itens.Qtd
        LinhaAtual = "4" & Format(Declaracao.Itens(I).Numero, "000") & Format(CDbl(Declaracao.Itens(I).Valor) * _
                            100, "000000000000000")
        LinhaAtual = LinhaAtual & CalculaCRC(LinhaAtual)
        GravaLinha nArquivo
    Next
    Declaracao.PreencheNotas
    For I = 1 To Declaracao.Notas.Qtd
        LinhaAtual = "5" & Format(TiraTudo(Declaracao.Notas.Nota(I).Destinatario), "000000000000000") & _
                            Format(Declaracao.Notas.Nota(I).TipoOperacao, "00") & _
                            Format(Declaracao.Notas.Nota(I).Numero, "000000") & _
                            Format(Declaracao.Notas.Nota(I).Data, "YYYYMMDD") & _
                            Format(Declaracao.Notas.Nota(I).Status, "00") & _
                            Format(Declaracao.Notas.Nota(I).ValorTotal, FormatoValor) & _
                            Format(Declaracao.Notas.Nota(I).BaseCalculo, FormatoValor) & _
                            Format(Declaracao.Notas.Nota(I).ValorMaterialICMS, FormatoValor) & _
                            Format(Declaracao.Notas.Nota(I).ImpostoDevido, FormatoValor) & _
                            Format(Declaracao.Notas.Nota(I).ImpostoRetido, FormatoValor)
                            
        LinhaAtual = LinhaAtual & CalculaCRC(LinhaAtual)
        GravaLinha nArquivo
    Next
    CriaDetalhes = True
    Exit Function
TrataErro:
    CriaDetalhes = False
    Avisa Err.Description
End Function


Private Sub GravaDec(ObjDec As cDeclaracao, Optional Lista As Object)
    
    If ObjDec.Gravar Then
        ObjDec.Finalizar (False)
        Lista.SubItems(5) = "OK"
        TotalAceitas = TotalAceitas + 1
    Else
        TotalRejeitadas = TotalRejeitadas + 1
        If Not Lista Is Nothing Then Lista.SubItems(5) = "ERRO"
    End If
    DoEvents
End Sub

Public Function ValidaArquivo(Caminho As String) As Boolean
    On Error GoTo trata
    Dim CodMunicipio As String
    Dim TotLinhas As Double
    Dim NovoCRC As Double
    Dim nValida As Integer
    Dim ErroValidacao() As String
    Dim I As Double
    nArquivo = FreeFile(0)
    If Trim(Dir(Caminho)) <> "" Then
        Open Caminho For Input As nArquivo
    Else
        Avisa "Caminho inválido."
        ValidaArquivo = False
        Exit Function
    End If
    TotLinhas = 0
    ValidaArquivo = True
    ReDim ErroValidacao(1 To 1) As String
    Do While Not EOF(nArquivo)
        Line Input #nArquivo, LinhaAtual
        TotLinhas = TotLinhas + 1
        Select Case Left(LinhaAtual, 1)
            Case 0
                CodMunicipio = Mid(LinhaAtual, 14, 4)
                DataArquivo = Mid(LinhaAtual, 6, 8)
                DataArquivo = Right(DataArquivo, 2) & "/" & Mid(DataArquivo, 3, 2) & "/" & Left(DataArquivo, 4)
                If CDbl(CodMunicipio) <> CDbl(AplicacoesVTFuncoes.Codigo_Municipio) Then
                    ReDim Preserve ErroValidacao(1 To TotLinhas) As String
                    ErroValidacao(TotLinhas) = "LINHA " & Format(TotLinhas, "0000") & ": Arquivo de declaracão com código de município inválido. Arquivo válido para " & Mid(LinhaAtual, 18, Len(Mid(LinhaAtual, 18)) - 4) & "."
                End If
            Case 9
                If TotLinhas <> CDbl(Mid(LinhaAtual, 2, 6)) Then
                    ReDim Preserve ErroValidacao(1 To TotLinhas) As String
                    ErroValidacao(TotLinhas) = "LINHA " & Format(TotLinhas, "0000") & ": Arquivo com número de registros diferente do informado."
                End If
        End Select
        NovoCRC = CDbl(CalculaCRC(Left(LinhaAtual, Len(LinhaAtual) - 4)))
        If NovoCRC <> CDbl(Right(LinhaAtual, 4)) Then
            ReDim Preserve ErroValidacao(1 To TotLinhas) As String
            ErroValidacao(TotLinhas) = "LINHA " & Format(TotLinhas, "0000") & ": Erro de CRC."
        End If
    Loop
    Close
    nValida = FreeFile(0)
    Open "C:\LogValida.log" For Output As nValida
    For I = 1 To UBound(ErroValidacao)
        If Trim(ErroValidacao(I)) <> "" Then
            Print #nValida, ErroValidacao(I)
            ValidaArquivo = False
        End If
    Next
    If UBound(ErroValidacao) = 1 And ValidaArquivo = True Then
        Print #nValida, "Arquivo Validado com sucesso!"
    End If
    Close
    Exit Function
trata:
    Avisa Err.Description
    Resume Next
End Function
