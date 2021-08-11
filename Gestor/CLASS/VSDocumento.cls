VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Aplicacao As Word.Application 'Object '
Private Documento As Word.Document 'Object '
Private Faixa As Word.Range 'Object '

Private Sub Class_Initialize()
    Set Aplicacao = CreateObject("Word.Application")
End Sub

Private Sub Class_Terminate()
    Set Aplicacao = Nothing
    Set Documento = Nothing
    Set Faixa = Nothing
End Sub

Public Function Novo(Optional strModelo As String) As Boolean
    If Not Aplicacao Is Nothing Then
        Aplicacao.Visible = True
        Set Documento = Aplicacao.Documents.Add(Template:=strModelo & IIf(InStr(1, strModelo, ".dot") > 0, "", ".dot"), Visible:=True)
        Novo = Not Documento Is Nothing
    End If
End Function

Public Function Selecionar(Optional intInicio As Integer = 0, Optional intFim As Integer = -1) As Boolean
    If Not Documento Is Nothing Then
        If intFim < 0 Then intFim = Documento.Characters.Count
        Set Faixa = Documento.Range(Start:=intInicio, End:=intFim)
        Selecionar = Not Faixa Is Nothing
    End If
End Function

Public Sub Substituir(strLocalizar As String, strSubstituir As String)
    If Faixa Is Nothing Then Selecionar
    If Not Faixa Is Nothing Then
        With Faixa.Find
            .Execute FindText:=strLocalizar, ReplaceWith:=strSubstituir, Replace:=wdReplaceAll
        End With
    End If
End Sub

Public Sub salvarComo(strArquivo As String)
    If Not Documento Is Nothing Then
        Documento.SaveAs FileName:=strArquivo & IIf(InStr(1, strArquivo, ".doc") > 0, "", ".doc")
    End If
End Sub

Public Sub Celula(intTabela As Integer, intLinha As Integer, intColuna As Integer, strValor As String)
    'If Faixa Is Nothing Then Selecionar
    If intTabela <= Documento.Tables.Count Then
        If intColuna <= Documento.Tables(intTabela).Columns.Count Then
            If intLinha > Documento.Tables(intTabela).Rows.Count Then
                inserirLinhas intTabela, intQuantidade:=intLinha - Faixa.Information(15)
            End If
            Documento.Tables(intTabela).Cell(intLinha, intColuna).Range.Text = strValor
        End If
    End If
End Sub

Public Sub inserirLinhas(intTabela As Integer, Optional intAntesLinha As Integer, Optional intQuantidade As Integer = 1)
    Dim i As Integer
    Dim Linha As Row
    
    Set Linha = Nothing
    If intTabela <= Documento.Tables.Count Then
        If intAntesLinha > 0 Then
            If intAntesLinha <= Documento.Tables(intTabela).Rows.Count Then
                Set Linha = Documento.Tables(intTabela).Rows(intAntesLinha)
            End If
        End If
        Do
            If Not (Linha Is Nothing) Then
                Documento.Tables(intTabela).Rows.Add BeforeRow:=Linha
            Else
                Documento.Tables(intTabela).Rows.Add
            End If
            i = i + 1
        Loop While i < intQuantidade
    End If
End Sub

Public Sub Ativar()
    Documento.Activate
    Aplicacao.Activate
End Sub

Public Sub Cabecalho(Estado As String, Prefeitura As String, Secretaria As String, Departamento As String)
    Dim Secao As Object
    
    Set Secao = Documento.Sections(1)
    If Secao.Headers(1).Exists = True Then '1= wdHeaderFooterPrimary
        Set Faixa = Secao.Headers(1).Range
        Substituir "@Estado", Estado
        Substituir "@Prefeitura", Prefeitura
        Substituir "@Secretaria", Secretaria
        Substituir "@Departamento", Departamento
    End If
End Sub

Public Sub textoObjeto(strLocalizar As String, strSubstituir As String)
    Dim i As Integer, J As Integer
    Dim qtdShapes As Integer
    Dim qtdItens As Integer
    
    qtdShapes = Documento.Shapes.Count
    For i = 1 To qtdShapes
        qtdItens = Documento.Shapes(i).GroupItems.Count
        For J = 1 To qtdItens
            With Documento.Shapes(i).GroupItems(J).TextFrame
                If .HasText Then
                    If Left(.TextRange.Text, Len(.TextRange.Text) - 1) = strLocalizar Then .TextRange.Text = strSubstituir
                End If
            End With
        Next J
    Next i
End Sub
