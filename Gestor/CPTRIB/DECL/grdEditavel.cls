VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "grdEditavel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim FirstRow As Integer, FirstCol As Integer

Public Sub EditaCelula(Grid As Object, CaixaDeTexto As Object)
    CaixaDeTexto.Visible = True
    CaixaDeTexto.Width = Grid.CellWidth
    CaixaDeTexto.Height = Grid.CellHeight
    CaixaDeTexto.Top = Grid.CellTop + Grid.Top
    CaixaDeTexto.Left = Grid.CellLeft + Grid.Left
    CaixaDeTexto.Text = Grid.Text
    CaixaDeTexto.SelStart = 0
    CaixaDeTexto.SelLength = Len(CaixaDeTexto.Text)
    CaixaDeTexto.ZOrder
    CaixaDeTexto.SetFocus
End Sub

Public Sub TeclaPressionada(Grid As Object, CaixaDeTexto As Object, KeyCode As Integer)
    If KeyCode = vbKeyReturn Then
        Call EditaCelula(Grid, CaixaDeTexto)
    End If
    If KeyCode = vbKeyDelete Then
        Grid.Text = ""
    End If
End Sub

Public Sub TextoKeyDown(KeyCode As Integer, Grid As Object, CaixaDeTexto As Object)
    If KeyCode = vbKeyReturn Then
        Grid.Text = CaixaDeTexto.Text
        If Grid.Row = Grid.Rows - 1 Then
            Grid.Row = Grid.Row
        Else
            Grid.Row = Grid.Row + 1
        End If
        Grid.SetFocus
        CaixaDeTexto.Visible = False
    End If
End Sub

Public Sub CarregaGrid(Grid As Object, Sql As String)
    Dim Rs As VSRecordset
    Dim i As Integer
            
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        Grid.Cols = Rs.Fields.Count
        Grid.FixedCols = 2
        Grid.Row = 0
        For i = 0 To Rs.Fields.Count - 1
            Grid.Col = i
            Grid.ColWidth(i) = IIf(i <> 1, Len(Rs.Fields.Item(i).Name) * 220, 8000)
            Grid = Rs.Fields.Item(i).Name
        Next
            
        i = 1
        Do While Not Rs.EOF
            Grid.AddItem CStr(Rs!Item), i
            Grid.TextMatrix(i, 1) = Rs!Descricao
            i = i + 1
            Rs.MoveNext
        Loop
    End If
End Sub

'Private Sub Text1_LostFocus()
'    Text1.Visible = False
'End Sub

'Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    firstcol = Grid.Col
'    firstrow = Grid.Row
'End Sub
