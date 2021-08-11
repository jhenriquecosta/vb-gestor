VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TipoOBJ
    xGrid
    XTree
End Enum

Private Util As New VSUtil

Public Function PegaParametro(Bdados As Object, Tipo As String) As String
    If bRegistrado Then
        Dim Sql As String
        Dim RS As VSRecordset
        
        Sql = "SELECT TPR_DESCRICAO FROM TAB_PARAMETRO WHERE TPR_TIPO = '" & Tipo & "'"
        If Bdados.AbreTabela(Sql, RS) Then
            PegaParametro = "" & RS(0)
        End If
        Bdados.FechaTabela RS
    End If
End Function


Public Sub CarregaTree(Bdados As Object, oTree As Object, Tipo As String)
    On Error GoTo Trata
    If bRegistrado Then
        Dim Sql As String
        Dim RS1 As VSRecordset
        Dim Rs2 As VSRecordset
        Dim REQ As String
        oTree.NodesCollection.Clear
    
        Sql = "SELECT * FROM TAB_GRUPO_COMPONENTE where (TGC_MODELO=1 OR TGC_MODELO=2) AND " & _
        " TGC_TTC_COD_TIPO=" & Tipo & " ORDER BY TGC_COD_GRUPO"
    
        If Bdados.AbreTabela(Sql, RS1) Then
            RS1.MoveFirst
    
            Do Until RS1.Eof
    
                REQ = IIf(RS1(4) = 1, "R", "G") & RS1(0)
    
                oTree.AddFolder REQ, , RS1(1), 1, , True
    
                If RS1(2) = 1 Then oTree.AddOption "N" & RS1(0), oTree.Nodes(REQ), "NDA", False
    
                Sql = "SELECT * FROM TAB_COMPONENTE WHERE TCO_TGC_COD_GRUPO=" & RS1(0) & " ORDER BY TCO_COD_COMPONENTE"
                If Bdados.AbreTabela(Sql, Rs2) Then
    
                    Do Until Rs2.Eof
                        If RS1(2) <> 1 Then
                            oTree.AddCheck "C" & Rs2(0), oTree.Nodes(REQ), CStr(Rs2(2)), 0, False
                            oTree.Nodes("C" & Rs2(0)).Tag = "CHECK"
                        Else
                            oTree.AddOption "C" & Rs2(0), oTree.Nodes(REQ), CStr(Rs2(2)), False
                        End If
                        Rs2.MoveNext
                    Loop
    
                Else
                    Util.Erro "Componentes do Grupo " & RS1(1) & " não encontrado."
                End If
                Bdados.FechaTabela Rs2
                
                RS1.MoveNext
    
            Loop
            
        End If
        Bdados.FechaTabela RS1
    End If
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Erro "Erro: " & Err.Number & " - " & Err.Description & "."
    End If
End Sub

Public Sub CarregaValoresTree(Bdados As Object, oTree As Object, Tabela As String, Condicao As String)
    On Error GoTo Trata
    If bRegistrado Then
        Dim Sql As String
        Dim RS As VSRecordset
        Dim i As Integer
        
        For i = 1 To oTree.NodesCollection.Count
            If oTree.Nodes(i).Tag = "CHECK" Then
                oTree.Value(i) = 0
            ElseIf oTree.Nodes(i).Children > 0 Then
                If oTree.Nodes(i + 1).Tag = "" Then oTree.Value(i + 1) = 1
            End If
        Next
        
        Sql = "SELECT * FROM " & Tabela & " where " & Condicao
        If Bdados.AbreTabela(Sql, RS) Then
            RS.MoveFirst
            
            Do Until RS.Eof
                If "" & RS(2) = "" Then
                    oTree.Value("C" & RS(1)) = 1
                End If
                RS.MoveNext
            Loop
        End If
        Bdados.FechaTabela RS
    End If
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Erro "Erro: " & Err.Number & " - " & Err.Description & "."
    End If
End Sub

Public Sub CarregaGrid(Bdados As Object, oGrid As Object, Tipo As String)
    On Error GoTo Trata
    If bRegistrado Then
        Dim Sql As String
        Dim RS As VSRecordset
        Dim ItmX As Object
        Dim i As Byte
        Dim REQ As String
        
        Sql = "SELECT tco_cod_componente ,tgc_nome ,tco_nome, tgc_modelo,TGC_REQUERIDO FROM " & _
        "TAB_GRUPO_COMPONENTE,TAB_COMPONENTE where TGC_MODELO>=3  AND " & _
        "TCO_TGC_COD_GRUPO=TGC_COD_GRUPO  AND " & _
        " TGC_TTC_COD_TIPO=" & Tipo & " ORDER BY TGC_COD_GRUPO,TCO_COD_COMPONENTE "
        
        oGrid.AllowColumnReorder = True
        oGrid.Arrange = 2 'lvwAutoTop
        oGrid.GridLines = True
        oGrid.View = 3 'lvwReport
        oGrid.FullRowSelect = True
        oGrid.ListItems.Clear
        oGrid.ColumnHeaders.Clear
    
        
        If Bdados.AbreTabela(Sql, RS) Then
            
            For i = 1 To 3
                oGrid.ColumnHeaders.Add , , , (oGrid.Width / 3) - 100
            Next
            
            Do Until RS.Eof
                
                REQ = IIf(RS(4) = 1, "R", "G") & RS(0)
                
                Set ItmX = oGrid.ListItems.Add(, REQ, CStr(RS(1)))
                ItmX.SubItems(1) = CStr(RS(2))
                ItmX.SubItems(2) = ""
                
                oGrid.ListItems(REQ).Tag = RS(3)
                
                RS.MoveNext
            Loop
        End If
        Bdados.FechaTabela RS
    End If
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Erro "Erro: " & Err.Number & " - " & Err.Description & "."
    End If
End Sub

Public Sub CarregaValoresGrid(Bdados As Object, oGrid As Object, Tabela As String, Condicao As String)
    On Error GoTo Trata
    If bRegistrado Then
        Dim Sql As String
        Dim RS As VSRecordset
        Dim i As Integer
        Dim K As String
        Dim Nao As Boolean
        
        For i = 1 To oGrid.ListItems.Count
            oGrid.ListItems(i).ListSubItems.item(2).Text = ""
        Next
        
        
        Sql = "SELECT * FROM " & Tabela & " where " & Condicao
        
        If Bdados.AbreTabela(Sql, RS) Then
            RS.MoveFirst
            
            Do Until RS.Eof
                If RS(2) <> "" Then
                    Nao = False
                    K = "R" & RS(1)
                    oGrid.ListItems(K).ListSubItems.item(2).Text = RS(2)
                End If
                RS.MoveNext
            Loop
        End If
        Bdados.FechaTabela RS
    End If
    Exit Sub
Trata:
    If Err.Number = 35601 Then
        If Not Nao Then
            Nao = True
            K = "G" & RS(1)
            Resume
        Else
            Util.Erro "Componentes '" & RS(1) & "' com erro."
        End If
    Else
        Util.Erro "Erro: " & Err.Number & " - " & Err.Description & "."
    End If
End Sub

Public Sub PegaValor(oGrid As Object)
    On Error GoTo Trata
    If bRegistrado Then
        If oGrid.ListItems.Count = 0 Then Exit Sub
        
        oGrid.SelectedItem.ListSubItems.item(2).Text = _
        Trim(UCase(InputBox("Digite o novo valor para '" & oGrid.SelectedItem.ListSubItems.item(1).Text & _
        "':", oGrid.SelectedItem.Text, oGrid.SelectedItem.ListSubItems.item(2).Text)))
    End If
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Util.Erro "Erro: " & Err.Number & " - " & Err.Description & "."
    End If
End Sub

Public Function CamposGridOK(oGrid As Object) As Boolean
    On Error GoTo Trata
    If bRegistrado Then
        Dim i As Integer
        
        For i = 1 To oGrid.ListItems.Count
            If ((oGrid.ListItems(i).Tag >= 4) And (oGrid.ListItems(i).ListSubItems.item(2).Text <> "")) Then
                If Not IsNumeric(oGrid.ListItems(i).ListSubItems.item(2).Text) Then
                    Util.Avisa "Valor numérico no campo " & oGrid.ListItems(i).Text & " é inválido."
                    Exit Function
                End If
            End If
            If Mid(oGrid.ListItems(i).Key, 1, 1) = "R" And _
            Trim(oGrid.ListItems(i).ListSubItems.item(2).Text) = "" Then
                Util.Avisa "Campo " & oGrid.ListItems(i).Text & " é requerido."
                Exit Function
            End If
        Next
        CamposGridOK = True
    End If
    Exit Function
Trata:
    If Err.Number <> 0 Then
        Util.Erro "Erro: " & Err.Number & " - " & Err.Description & "."
    End If
End Function

Public Function CamposTreeOK(oTree As Object) As Boolean
    On Error GoTo Trata
    If bRegistrado Then
        Dim i As Integer
        Dim J As Integer
        Dim K As Integer
        Dim OK As Boolean
        
        For i = 1 To oTree.NodesCollection.Count
            If Mid(oTree.Nodes(i).Key, 1, 1) = "R" Then
                
                J = oTree.Nodes(i).Children
                OK = False
                For K = i + 1 To i + J
                    If Mid(oTree.NodesCollection(K).Key, 1, 1) = "N" And oTree.Value(K) = 1 Then Exit For
                    
                    If oTree.Value(K) = 1 Then
                        OK = True
                        Exit For
                    End If
                Next
                If Not OK Then
                    CamposTreeOK = False
                    Util.Avisa "Campo " & oTree.NodesCollection(i).Text & " é requerido."
                    Exit Function
                End If
                    
            End If
        Next
        CamposTreeOK = True
    End If
    Exit Function
Trata:
    If Err.Number <> 0 Then
        Util.Erro "Erro: " & Err.Number & " - " & Err.Description & "."
    End If
End Function

Public Function GravarDetalhes(Bdados As Object, Obj As Object, Tipo As TipoOBJ, Chave As String, Tabela As String, CampoChave As String, CampoComponente As String, CampoValor As String) As Boolean
    On Error GoTo Trata
    If bRegistrado Then
        Dim i As Integer
        
        Bdados.AbreTrans
        
        If Tipo = xGrid Then
            For i = 1 To Obj.ListItems.Count
                If Trim(Obj.ListItems(i).ListSubItems.item(2).Text) <> "" Then
                    If Not Bdados.InsereDados(Tabela, Bdados.PreparaValor(Chave, Mid(Obj.ListItems(i).Key, 2), Obj.ListItems(i).ListSubItems.item(2).Text), CampoChave & "," & CampoComponente & "," & CampoValor) Then
                        Bdados.CancelaTrans
                        GravarDetalhes = False
                        Exit Function
                    End If
                End If
            Next
            
        ElseIf Tipo = XTree Then
        
            For i = 1 To Obj.NodesCollection.Count
                If Left(Obj.Nodes(i).Key, 1) = "C" Then
                    If Obj.Value(i) = 1 Then
                        If Not Bdados.InsereDados(Tabela, Bdados.PreparaValor(Chave, Mid(Obj.Nodes(i).Key, 2)), CampoChave & "," & CampoComponente) Then
                            Bdados.CancelaTrans
                            GravarDetalhes = False
                            Exit Function
                        End If
                    End If
                End If
            Next
        End If
        
        Bdados.GravaTrans
        GravarDetalhes = True
    End If
    Exit Function
Trata:
    If Err.Number <> 0 Then
        Bdados.CancelaTrans
        GravarDetalhes = False
        Util.Erro "Erro: " & Err.Number & " - " & Err.Description & "."
        Exit Function
    End If
End Function

Public Function ApagaDetalhes(Bdados As Object, Tabela As String, Condicao As String) As Boolean
    If bRegistrado Then ApagaDetalhes = Bdados.DeletaDados(Tabela, Condicao)
End Function

Public Function PegaTabGeral(Bdados As Object, Tabela As String, Optional Objeto As Object, Optional Condicao As String) As String
    On Error GoTo Trata
    
    If bRegistrado Then
        Dim Sql As String
        Dim RS As VSRecordset
        
        
        Sql = "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_TIPO = " & _
        " (SELECT DISTINCT  TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME = '" & Tabela & "') AND TGE_CODIGO <> 0 "
        
        If Condicao <> "" Then
            Sql = Sql & " AND " & Condicao
        End If
        
        Sql = Sql & " ORDER BY TGE_CODIGO"
        
        If Bdados.AbreTabela(Sql, RS) Then
            RS.MoveFirst
            PegaTabGeral = RS(0)
            If Not Objeto Is Nothing Then
                
                If TypeOf Objeto Is TextBox Then
                    Objeto = RS(0)
                ElseIf TypeOf Objeto Is ComboBox Then
                    Objeto.Clear
                    Do Until RS.Eof
                        Objeto.AddItem RS(0)
                        RS.MoveNext
                    Loop
                End If
            End If
        End If
        Bdados.FechaTabela RS
    End If
    Exit Function
Trata:
    If Err.Number <> 0 Then
        Util.Erro "Erro: " & Err.Number & " - " & Err.Description & "."
    End If
End Function

Public Function PegaCodigoNaGeral(Bdados As Object, Tabela As String, Nome As String) As String
    On Error GoTo Trata
    
    If bRegistrado Then
        Dim Sql As String
        Dim RS As VSRecordset
        
        Sql = "SELECT TGE_CODIGO FROM TAB_GERAL WHERE TGE_TIPO = " & _
        " (SELECT DISTINCT TGE_TIPO FROM TAB_GERAL WHERE TGE_NOME = '" & Tabela & "') AND TGE_NOME = '" & Nome & "'"
        
        
        If Bdados.AbreTabela(Sql, RS) Then
            RS.MoveFirst
            PegaCodigoNaGeral = RS(0)
        End If
        Bdados.FechaTabela RS
    End If
    Exit Function
Trata:
    If Err.Number <> 0 Then
        Util.Erro "Erro: " & Err.Number & " - " & Err.Description & "."
    End If
End Function

Private Sub Class_Initialize()
    ValidaComponente "CLASS"
End Sub
