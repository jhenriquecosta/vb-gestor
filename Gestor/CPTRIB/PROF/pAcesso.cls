VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Usuario As pUsuario
Public Lotacao As pLotacao

Public Enum TipoAcesso
    taSistema = 1
    taLotacao = 2
End Enum

Public Enum TipoPermissao
    tpPrivilegio = 1
    tpAcesso = 2
End Enum


Public Function GravarAcessoSistema(Tree As Object, Usuario As String, Optional TipoPermissao As TipoPermissao) As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim I As Integer
    Dim Tabela As String
    If TipoPermissao = tpAcesso Then
        Tabela = "TAB_ACESSO_USUARIO"
    Else
        Tabela = "TAB_ACESSO_USUARIO_PRIVILEGIO"
    End If
    ' LIMPA TODOS OS ACESSOS DO USUÁRIO
    BDados.DeletaDados Tabela, "TAU_TUS_COD_USUARIO = '" & Usuario & "'"
    
    'GRAVA OS ACESSOS INFORMADOS
    For I = 1 To Tree.Nodes.Count
        ' If Util.ParseString(Tree.Nodes.Item(I).Key, ":", 1) = "FORMULARIO" And Tree.Nodes.Item(I).Checked Then
        If Left(Tree.Nodes.Item(I).Key, 1) = "3" And Tree.Nodes.Item(I).Checked Then
            Valores = BDados.PreparaValor(Usuario, Mid(Tree.Nodes.Item(I).Key, 2, 4), Mid(Tree.Nodes.Item(I).Key, 6, 4), Mid(Tree.Nodes.Item(I).Key, 10, 3))
            Campos = "TAU_TUS_COD_USUARIO,TAU_TSI_COD_SISTEMA,TAU_TMO_COD_MODULO,TAU_TFO_COD_FORMULARIO"
            If BDados.InsereDados(Tabela, Valores, Campos) Then GravarAcessoSistema = True
       End If
    Next
    
End Function

Public Function GravarAcessoLotacao(Tree As Object, Usuario As String) As Boolean
    Dim Valores As String
    Dim Campos As String
    Dim I As Integer
        
    ' LIMPA TODOS OS ACESSOS DO USUÁRIO
    BDados.DeletaDados "TAB_ACESSO_LOTACAO", "TAL_TUS_COD_USUARIO = '" & Usuario & "'"
    
    'GRAVA OS ACESSOS INFORMADOS
    For I = 1 To Tree.Nodes.Count
        If Tree.Nodes.Item(I).Checked Then
            Valores = BDados.PreparaValor(Usuario, Util.ParseString(Tree.Nodes.Item(I).Key, ":", 2))
            Campos = "TAL_TUS_COD_USUARIO, TAL_TLO_CODIGO"
            If BDados.InsereDados("TAB_ACESSO_LOTACAO", Valores, Campos) Then GravarAcessoLotacao = True
        End If
    Next

End Function

Public Function ExibeSistema(Tree As Object, Optional TipoPermissao As TipoPermissao) As Boolean
    On Error GoTo Trata
    Dim RsSistemas As VSRecordset
    Dim RsModulos As VSRecordset
    Dim RsSubModulos As VSRecordset
    Dim RsFormularios As VSRecordset
    Dim ChaveSubFormulario As String
    Dim ChaveFormulario As String
    Dim ChaveSubModulo As String
    Dim RS As VSRecordset
    Dim sql As String
    Dim No As Node
    Dim Sistema As String
    Dim Modulo As String
    Dim Formulario As String
    Dim NohPai As String
    Dim Chave As String
    Dim Index As String
    Dim Fig As String
    Dim TabelaAcesso As String
    
    Tree.Nodes.Clear
    Index = "0VSIS"
    
    Fig = "TCTA401"
    'DoEvents
    Set No = Tree.Nodes.Add(, , Index, Temp.PegaParametro(BDados, "SISTEMA"), Fig)
    No.Expanded = True
     
    'Pego os sistemas...
    If TipoPermissao = tpAcesso Then
        sql = "SELECT TSI_COD_SISTEMA, TSI_NOME FROM TAB_SISTEMA WHERE " & _
            " TSI_COD_SISTEMA IN (SELECT DISTINCT TAU_TSI_COD_SISTEMA FROM " & _
            " TAB_ACESSO_USUARIO_PRIVILEGIO  WHERE TAU_TUS_COD_USUARIO = '" & User & "') " & _
            " ORDER BY TSI_NOME"
        TabelaAcesso = "TAB_ACESSO_USUARIO_PRIVILEGIO"
    Else
        sql = "SELECT TSI_COD_SISTEMA, TSI_NOME FROM TAB_SISTEMA ORDER BY TSI_NOME"
        TabelaAcesso = "TAB_ACESSO_USUARIO"
    End If
    If BDados.AbreTabela(sql, RsSistemas) Then
        Do Until RsSistemas.EOF
            Fig = "CLOSE"
            'Moto o sistema...
             NohPai = "1" & RsSistemas!TSI_COD_SISTEMA
             Set No = Tree.Nodes.Add(Index, tvwChild, NohPai, RsSistemas!TSI_NOME, Fig)
            'Pego os módulos...
            'TMO_TIPO = 1 AND
            sql = "SELECT TMO_COD_MODULO,TMO_NOME FROM TAB_MODULO "
            If TipoPermissao = tpAcesso Then
                sql = sql & " WHERE  TMO_COD_MODULO IN (SELECT DISTINCT TMO_REF_COD_MODULO FROM TAB_MODULO WHERE " & _
                " TMO_COD_MODULO IN (SELECT DISTINCT TAU_TMO_COD_MODULO FROM " & TabelaAcesso & _
                "  WHERE TAU_TSI_COD_SISTEMA = '" & RsSistemas!TSI_COD_SISTEMA & "'"
                sql = sql & " AND TAU_TUS_COD_USUARIO = '" & User & "'))"
            Else
                sql = sql & " WHERE TMO_TSI_COD_SISTEMA = '" & RsSistemas!TSI_COD_SISTEMA & "' AND TMO_TIPO <> 2"
            End If
            sql = sql & " ORDER BY TMO_NOME"
            If BDados.AbreTabela(sql, RsModulos) Then
                'Monto tree...
                Do Until RsModulos.EOF
                    Fig = "CLOSE"
                    Chave = "2" & CStr(RsSistemas!TSI_COD_SISTEMA) & CStr(RsModulos!TMO_COD_MODULO)
                    Set No = Tree.Nodes.Add(NohPai, tvwChild, Chave, RsModulos!TMO_NOME, Fig)
                    'Pego os Sub Modulos...
                    sql = "SELECT TMO_COD_MODULO, TMO_NOME FROM TAB_MODULO WHERE TMO_TIPO = 2  AND TMO_REF_COD_MODULO = '" & RsModulos!TMO_COD_MODULO & "'"
                    If TipoPermissao = tpAcesso Then
                        sql = sql & " AND " & _
                        " TMO_COD_MODULO IN (SELECT DISTINCT TAU_TMO_COD_MODULO FROM " & TabelaAcesso & _
                        " WHERE TAU_TSI_COD_SISTEMA = '" & RsSistemas!TSI_COD_SISTEMA & "'"
                        sql = sql & " AND TAU_TUS_COD_USUARIO = '" & User & "')"
                    Else
                        sql = sql & " AND TMO_TSI_COD_SISTEMA = '" & RsSistemas!TSI_COD_SISTEMA & "'"
                    End If
                    sql = sql & " ORDER BY TMO_NOME"
                    If BDados.AbreTabela(sql, RsSubModulos) Then
                        Do Until RsSubModulos.EOF
                            Fig = "CLOSE"
                            ChaveSubModulo = "2" & CStr(RsSistemas!TSI_COD_SISTEMA) & CStr(RsSubModulos.Fields("TMO_COD_MODULO"))
                            Set No = Tree.Nodes.Add(Chave, tvwChild, ChaveSubModulo, RsSubModulos!TMO_NOME, Fig)
                            'Pego os formulários do SubModulo..
                            sql = "SELECT TFO_COD_FORMULARIO, TFO_NOME FROM TAB_FORMULARIO "
                            If TipoPermissao = tpAcesso Then
                                sql = sql & " WHERE TFO_TMO_COD_MODULO " & BDados.Concatena & " TFO_COD_FORMULARIO IN (SELECT DISTINCT TAU_TMO_COD_MODULO " & BDados.Concatena & " TAU_TFO_COD_FORMULARIO FROM " & TabelaAcesso & _
                                " WHERE TAU_TMO_COD_MODULO = '" & RsSubModulos!TMO_COD_MODULO & "')"
                            Else
                                sql = sql & " WHERE TFO_TMO_COD_MODULO = '" & RsSubModulos!TMO_COD_MODULO & "'"
                            End If
                            sql = sql & "  ORDER BY TFO_NOME"
                            If BDados.AbreTabela(sql, RsFormularios) Then
                                
                                Do Until RsFormularios.EOF
                                    Fig = "CLOSE"
                                    'Monto o tree...
                                    ChaveSubFormulario = "3" & CStr(RsSistemas!TSI_COD_SISTEMA) & CStr(RsSubModulos!TMO_COD_MODULO & RsFormularios!TFO_COD_FORMULARIO)
                                    'Set No = Tree.Nodes.Add(ChaveSubModulo, tvwChild, ChaveSubFormulario, RsFormularios!TFO_NOME, Fig)
                                    Set No = Tree.Nodes.Add(ChaveSubModulo, tvwChild, ChaveSubFormulario, RsFormularios!TFO_NOME, Fig)
                                    Tree.Nodes.Item(ChaveSubFormulario).Tag = RsSistemas!TSI_COD_SISTEMA
                                    'DoEvents
                                    RsFormularios.MoveNext
                                Loop
                            End If
                            RsSubModulos.MoveNext
                        Loop
                    End If
                    'Pego os formulários do Modulo..
                    sql = "SELECT TFO_COD_FORMULARIO, TFO_NOME FROM TAB_FORMULARIO "
                    If TipoPermissao = tpAcesso Then
                        sql = sql & " WHERE " & _
                        " TFO_TMO_COD_MODULO " & BDados.Concatena & " TFO_COD_FORMULARIO IN (SELECT DISTINCT TAU_TMO_COD_MODULO " & BDados.Concatena & " TAU_TFO_COD_FORMULARIO FROM " & TabelaAcesso & _
                        " WHERE TAU_TMO_COD_MODULO = '" & RsModulos!TMO_COD_MODULO & "'"
                        sql = sql & " AND TAU_TUS_COD_USUARIO = '" & User & "')"
                    Else
                        sql = sql & " WHERE TFO_TMO_COD_MODULO = '" & RsModulos!TMO_COD_MODULO & "'"
                    End If
                    sql = sql & " ORDER BY TFO_NOME"
                    If BDados.AbreTabela(sql, RsFormularios) Then
                        Do Until RsFormularios.EOF
                            Fig = "CLOSE"
                            'Monto o tree...
                            ChaveFormulario = "3" & CStr(RsSistemas!TSI_COD_SISTEMA) & CStr(RsModulos!TMO_COD_MODULO & RsFormularios!TFO_COD_FORMULARIO)
                            Set No = Tree.Nodes.Add(Chave, tvwChild, ChaveFormulario, RsFormularios!TFO_NOME, Fig)
                            Tree.Nodes.Item(ChaveFormulario).Tag = RsSistemas!TSI_COD_SISTEMA
                            'DoEvents
                            RsFormularios.MoveNext
                        Loop
                    End If
                    RsModulos.MoveNext
                    'DoEvents
                Loop
            End If
            RsSistemas.MoveNext
            'DoEvents
        Loop
    End If
    Screen.MousePointer = 0
    Exit Function
Trata:
    If ERR.Number <> 3021 Then
        Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
        Screen.MousePointer = 0
        Exit Function
        Resume
    End If
End Function

Public Function MarcaAcessos(Tree As Object, Usuario As String, Acesso As TipoAcesso, Optional TipoPermissao As TipoPermissao) As Boolean
    On Error GoTo Trata
    
    Dim sql As String
    Dim RS As VSRecordset
    Dim Chave As String
    
    If Acesso = taSistema Then
        If TipoPermissao = tpAcesso Then
            sql = "SELECT * FROM TAB_ACESSO_USUARIO WHERE TAU_TUS_COD_USUARIO = '" & Usuario & "'"
        Else
            sql = "SELECT * FROM TAB_ACESSO_USUARIO_PRIVILEGIO WHERE TAU_TUS_COD_USUARIO = '" & Usuario & "'"
        End If
    ElseIf Acesso = taLotacao Then
        sql = "SELECT * FROM TAB_ACESSO_LOTACAO WHERE TAL_TUS_COD_USUARIO = '" & Usuario & "'"
    End If
    
    If BDados.AbreTabela(sql, RS) Then
        Do Until RS.EOF
            If Acesso = taSistema Then
                'Chave = "FORMULARIO:" & RS!TAU_TSI_COD_SISTEMA & ":" & RS!TAU_TMO_COD_MODULO & ":" & RS!TAU_TFO_COD_FORMULARIO
                Chave = "3" & RS!TAU_TSI_COD_SISTEMA & RS!TAU_TMO_COD_MODULO & RS!TAU_TFO_COD_FORMULARIO
            ElseIf Acesso = taLotacao Then
                Chave = "LOTACAO:" & RS!TAL_TLO_CODIGO
            End If
            Tree.Nodes(Chave).Checked = True
            RS.MoveNext
        Loop
    End If
    MarcaAcessos = True
    BDados.FechaTabela RS
    Exit Function
Trata:
        If ERR.Number <> 35601 Then
            Util.Avisa "Erro: " & ERR.Number & " - " & ERR.Description & "."
            Screen.MousePointer = 0
        Else
            If Usuario <> Aplicacoes.Usuario Then
                'Resume Next
                'RAIMUNDO
                Util.Avisa "Acesso negado. Usuário com acessos diferentes de seus privilégios."
            Else
                MarcaAcessos = True
            End If
        End If
End Function
