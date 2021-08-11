VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeViewBci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Public Type Edificacao
'    Subjetivo(1 To 200) As Boolean
'    Componente(1 To 200) As String
'End Type

Public Sub CarregaListaComponentes(TreeView As Object)
    Dim Sql As String
    Dim rs As VSRecordset
    Dim RsOpt As VSRecordset
    Dim Pai As Object
    Dim i As Integer
    TreeView.Clear
'    On Error Resume Next
    'tgc_cod_grupo < 100
    Sql = "Select tgc_cod_grupo,tgc_nome from tab_grupo_componente_avancado where  " & _
            " tgc_categoria= " & TreeView.Tag & " order by tgc_cod_grupo"
    If Bdados.AbreTabela(Sql, rs) Then
        rs.MoveFirst
        Do
            
            TreeView.AddCheck "G" & Format(rs!tgc_cod_grupo, "000"), , Format(rs!tgc_cod_grupo, "000") & " " & rs!tgc_nome, , True
            TreeView.Value("G" & Format(rs!tgc_cod_grupo, "000")) = 0
            Sql = "Select tco_cod_componente,tco_descricao_componente from tab_componente_avancado where tco_grupo=" & _
                rs!tgc_cod_grupo & " order by tco_cod_componente"
            If Bdados.AbreTabela(Sql, RsOpt) Then
                RsOpt.MoveFirst
                Set Pai = TreeView.Nodes("G" & Format(rs!tgc_cod_grupo, "000"))
                Do
                    TreeView.AddCheck Format(RsOpt!tco_cod_componente, "000") & Pai, _
                        Pai, Format(RsOpt!tco_cod_componente, "000") & " " & RsOpt!tco_descricao_componente, False
                    RsOpt.MoveNext
                Loop While Not RsOpt.EOF
            End If
            rs.MoveNext
        Loop While Not rs.EOF
'        For i = 1 To TreeView.NodesCollection.Count
'            If IsNumeric(Left(TreeView.NodesCollection(i).Key, 3)) Then
'                TreeView.Value(i) = 0
'            End If
'        Next
    End If
    On Error Resume Next
    TreeView.ExpandAll
    DoEvents
End Sub

Public Sub AdicionaEdificacao(TreeView As Object, ByRef Edificacoes() As Edificacao, Unidade As Integer)
    Dim i As Integer
    Edificacoes(Unidade).Deletado = False
    For i = 1 To TreeView.NodesCollection.Count
        If IsNumeric(Left(TreeView.NodesCollection(i).Key, 3)) Then
            If TreeView.Value(i) = 1 Then
                Edificacoes(Unidade).Componente(CInt(Mid(TreeView.NodesCollection(i).Key, 4, 3))) _
                = CInt(Left(TreeView.NodesCollection(i).Key, 3))
            End If
        End If
    Next
    CarregaListaComponentes TreeView
End Sub


Public Sub MarcaUnico(TreeView As Object, ItemNode_Key As String, Value As Integer)
    Dim i As Integer
    If Value = 1 And IsNumeric(Left(ItemNode_Key, 3)) Then
        For i = 1 To TreeView.NodesCollection.Count
            If Mid(TreeView.NodesCollection(i).Key, 4, 3) = Mid(ItemNode_Key, 4, 3) Then
                If IsNumeric(Left(TreeView.NodesCollection(i).Key, 3)) Then
                    If Left(ItemNode_Key, 3) <> Left(TreeView.NodesCollection(i).Key, 3) Then
                        TreeView.Value(i) = 0
                    End If
                End If
            End If
        Next
    ElseIf Value = 1 Then
        For i = 1 To TreeView.NodesCollection.Count
            If Mid(TreeView.NodesCollection(i).Key, 4, 3) = Mid(ItemNode_Key, 2, 3) Then
                If IsNumeric(Left(TreeView.NodesCollection(i).Key, 3)) Then
                    TreeView.Value(i) = 0
                End If
            End If
        Next
    End If
End Sub


Public Sub SetaTreeViewEdificacao(TreeView As Object, ByRef Edificacoes() As Edificacao, Unidade As Integer)
    Dim i As Integer
    For i = 1 To TreeView.NodesCollection.Count
        If IsNumeric(Left(TreeView.NodesCollection(i).Key, 3)) Then
            If CInt(Nvl(Edificacoes(Unidade).Componente(CInt(Mid(TreeView.NodesCollection(i).Key, 4, 3))), 0)) _
                    = CInt(Left(TreeView.NodesCollection(i).Key, 3)) Then
                TreeView.Value(i) = 1
            End If
        End If
    Next
End Sub
