VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Notificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Sub ExibirNotificacoes(Grid As Object, Optional Im As String, Optional PeriodoInicial As String, Optional PeriodoFinal As String, Optional Tipo As Integer, Optional SeqInicial As Double, Optional SeqFinal As Double)
    Dim Sql As String
    Dim Condicao As String
    
    Condicao = "TNT_INSCRICAO = TCI_IM and TNT_TIPO = 1 "
    If Tipo > 0 Then
        Condicao = Condicao & " and TNT_TIPO_NOTIFICACAO =" & Tipo
    Else
        Condicao = Condicao & " and TNT_TIPO_NOTIFICACAO =1"
    End If
    
    If Trim$(Im) <> "" Then
        Condicao = Condicao & "and TNT_INSCRICAO='" & Im & "'"
    End If
    If Trim$(PeriodoInicial) <> "" And Trim$(PeriodoFinal) <> "" Then
        Condicao = Condicao & " and TNT_DT_EMISSAO >= convert(datetime,'" & Format(PeriodoInicial, "DD/MM/YYYY") & "',103)" & _
            " AND TNT_DT_EMISSAO  <= convert(datetime,'" & Format(PeriodoFinal, "DD/MM/YYYY") & "',103)"
    End If
    
    If SeqInicial > 0 Then
        Condicao = Condicao & " and tnt_cod_notificacao >= " & SeqInicial
    End If
    If SeqFinal > 0 Then
        Condicao = Condicao & " and tnt_cod_notificacao <= " & SeqFinal
    End If
    Sql = "SELECT TNT_COD_NOTIFICACAO AS Numero,TCI_NOME AS Contribuinte, " & _
            " TNT_DT_EMISSAO as Emissao, " & _
            " TNT_VENCIMENTO as Vencimento, " & _
            " TNT_VALOR_NOTIFICACAO as Valor, TNT_TIPO_NOTIFICACAO as Tipo" & _
        " FROM TAB_NOTIFICACAO,TAB_CONTRIBUINTE "

    If Condicao <> "" Then
        Sql = Sql & " WHERE " & Condicao
    End If
    Sql = Sql & " ORDER BY TNT_VENCIMENTO"
    Grid.Preencher Bdados, Sql
    If Grid.ListItems.Count > 0 Then
        If Grid.ListItems.Count > 0 Then
            Grid.Mensagem = "Total : " & Format(Grid.Colunas(5).Soma, Const_Monetario)
        Else
            Grid.Mensagem = ""
        End If
    End If
End Sub



