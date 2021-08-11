Attribute VB_Name = "MOD_AUXILIAR"
Function carregaCampo(xCampo As ADODB.Field) As Variant
  On Error Resume Next
    Dim Tipo As Integer
    Dim vlrc As Variant
    Tipo = xCampo.Type
    
    Select Case Tipo
        Case adBoolean
            vlrc = False
        Case adVarWChar Or adLongVarBinary Or adGUID Or adChar Or adVarChar
            vlrc = ""
        Case ElseadVarChar
            vlrc = 0
    End Select
    
    carregaCampo = IIf(IsNull(xCampo), vlrc, xCampo)
    
End Function

Public Sub configRelatorio(ByRef RPT As ActiveReport, Optional Sql As String)
  
  RPT.DataControl1.ConnectionString = Bdados.Conexao.ConnectionString
  If Sql <> "" Then RPT.DataControl1.Source = Sql
  RPT.Refresh
  
 'RPT.Restart

End Sub

Public Sub carregaCombo(Combo As Control, SQLOrigem As String, Optional ArgCompativel As String)
    On Error GoTo Erro
    Dim campo As Integer
    Dim pCampos()
    
    Dim Rs As Recordset
    Combo.Clear
    
    If InStr(1, SQLOrigem, "EXEC", 1) > 0 Then
         Set Rs = ExecutaSP(SQLOrigem, pCampos)
    Else
        If InStr(1, SQLOrigem, "SELECT", 1) > 0 Then
            Set Rs = ExecutaSP(SQLOrigem, pCampos)
        Else
            SQLOrigem = "SELECT * FROM " & SQLOrigem
            Set Rs = ExecutaSP(SQLOrigem, pCampos)
        End If
    End If
    
    If Rs.Fields.Count = 1 Then
        campo = 0
    Else
        campo = 1
    End If
    
    Do Until Rs.EOF
        Combo.AddItem carregaCampo(Rs(campo))
        Rs.MoveNext
    Loop
   
  '  fechaConexao (Rs.Status)
     Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
Erro:
    MostraErro "EncheCombo -> " & SQLOrigem
    
End Sub

Public Function abreConexao(ByRef Conex As Connection)

    Set Conex = New Connection
    Dim strconn As String
    strconn = Bdados.Conexao.DBConnection.ConnectionString
    Conex.Open strconn
    
    'Set Conex = abreConexao
End Function

Public Function fechaConexao(ByRef Conex As Connection)
    Conex.Close
End Function
Public Function ExecutaSP(Nome_Procedure As String, Campos_Valores() As Variant, Optional bRet As Boolean = False, Optional Banco As Connection) As ADODB.Recordset
    'seta os parametros
On Error GoTo erros
     
     Dim par() As ADODB.Parameter
     Dim cmd As ADODB.Command
     Set cmd = New ADODB.Command
     
     
     
     Dim P_Nome
     Dim P_Valor
     Dim linha As Integer
     Dim coluna As Integer
     
   '  bdcnn = abreConexao()
     If Banco Is Nothing Then abreConexao Banco
        
     cmd.ActiveConnection = Banco
     cmd.CommandText = Nome_Procedure
     cmd.CommandType = adCmdText
          
     coluna = 0
     mlinha = 0
     
     If InStr(1, Nome_Procedure, "SP_UPD") > 0 Then bRet = True
     If InStr(1, Nome_Procedure, "SP_INS") > 0 Then bRet = True
        
     If Uboundsp(Campos_Valores) >= 0 Then
        cmd.CommandType = adCmdStoredProc
        ReDim par(UBound(Campos_Valores))
        For linha = 0 To UBound(Campos_Valores)
            P_Nome = Campos_Valores(linha, coluna)
            P_Valor = Campos_Valores(linha, coluna + 1)
               
            If InStr(1, P_Nome, "OUTPUT") > 0 Then
               mParam = adParamOutput
               bRet = True
               mlinha = linha
            Else
               mParam = adParamInput
            End If
            
            If VarType(P_Valor) = vbDate Then
               If P_Valor = "00:00:00" Then P_Valor = Null
               Set par(linha) = cmd.CreateParameter(P_Nome, adDate, adParamInput, 8, P_Valor)
            ElseIf VarType(P_Valor) = vbInteger Or VarType(P_Valor) = vbLong Or VarType(P_Valor) = vbDecimal Or VarType(P_Valor) = vbDouble Then
               Set par(linha) = cmd.CreateParameter(P_Nome, adInteger, mParam, 4, P_Valor)
            ElseIf VarType(P_Valor) = vbString Then
               Set par(linha) = cmd.CreateParameter(P_Nome, adVarChar, adParamInput, IIf(Len(P_Valor) = 0, 1, Len(P_Valor)), P_Valor)
            ElseIf VarType(P_Valor) = vbCurrency Then
               Set par(linha) = cmd.CreateParameter(P_Nome, adCurrency, adParamInput, 4, P_Valor)
            ElseIf VarType(P_Valor) = vbLong Then
               Set par(linha) = cmd.CreateParameter(P_Nome, adBigInt, adParamInput, 8, P_Valor)
            ElseIf VarType(P_Valor) = vbBoolean Then
               Set par(linha) = cmd.CreateParameter(P_Nome, adBoolean, adParamInput, 1, P_Valor)
            ElseIf VarType(P_Valor) = vbSingle Then
               Set par(linha) = cmd.CreateParameter(P_Nome, adSingle, adParamInput, 4, P_Valor)
            End If
            cmd.Parameters.Append par(linha)
        Next
     End If
     
     Dim LngRec As Long
     Dim rsRetorno As New Recordset
          
     If bRet = True Then
        cmd.Execute
        If cmd.Parameters(mlinha).Direction = adParamOutput Then
            rsRetorno.Fields.Append "Codigo", adInteger, 4
            rsRetorno.Open
            rsRetorno.AddNew
            rsRetorno(0).Value = cmd.Parameters(mlinha).Value
            rsRetorno.Update
        End If
     '   INTULTCODCAD = cmd.Parameters(mlinha).Value
     Else
        With rsRetorno
             .CursorLocation = adUseClient
            .Open cmd, CursorType:=adOpenStatic, _
                    Options:=adCmdText
            Set .ActiveConnection = Nothing
        End With
        Set cmd = Nothing
     End If
     
     Set ExecutaSP = rsRetorno
     Set ExecutaSP.ActiveConnection = Nothing
     fechaConexao Banco
     Exit Function

erros:
     
     If Err.Number = -2147217873 Then
        MsgBox "Erro Na Tentativa de Gravar Registro Solicitado " & vbcrfl & "Erro: " & Err.Description
     Else
        MsgBox "Erro: " & Err.Description
     End If
     
End Function


Function Uboundsp(v As Variant) As Long
  On Error Resume Next
  Uboundsp = UBound(v)
  If Err <> 0 Then Uboundsp = -1
  On Error GoTo 0
End Function
Public Sub MostraErro(Optional FrmErr As String)   'exibe mensagem de erro personalizada
'by henrique
Dim Mens As String
Dim mErr As Double
Dim mDes As String
Dim mSou As String
Dim mDll As String

mErr = Err.Number
mDes = Err.Description
mSou = Err.Source
mDll = Err.LastDllError

    Select Case mErr
        Case 2000
          Exit Sub
        Case 91
           Mens = "Voce Tentou Executar Uma Operacao Com Um Objeto Que Ainda Nao Foi Criado"
        Case 13
           Mens = "Um Caracter Inválido Gerou Um ERRO No Sistema"
        Case 482
           Mens = "O Sistema Não Conseguiu Inicializar o Dispositivo de Impressão"
        Case 3021
          Mens = "Não Há Registros Disponiveis Na Tabela Solicitada"
        Case 3022
          Mens = "Codigo Já Cadastrado"
        Case 3045
          Mens = "Arquivo em Uso, Tente Novamente !"
        Case 3197 Or 3046 Or 3158 Or 3186 Or 3187 Or 3188 Or 3189 Or 3218 Or 3260
          Mens = "Impossivél Gravar, Registro em Uso !"
        Case 3200
          Mens = "Impossivél Apagar, Registro em Uso !"
        Case 3265
          Mens = "Um Campo Informado Não Existe Na Tabela Ou Na Consulta Solicitada!" & vbCrLf & "Consulte o Programador"
        Case 3262
         Mens = "O Sistema Não Conseguiu Travar a Tabela!"
       Case 3029
          Mens = "O USUARIO OU A SENHA INFORMADA É INVÁLIDO"
       Case 3033
          Mens = "Você não tem as permissões necessárias para usar o objeto <nome>." & vbCrLf & "Peça ao administrador de sistema ou à pessoa que criou este objeto que estabeleça as permissões apropriadas para você. (Erro 3033)"
       Case 3315
          mensd = mDes
          Mens = "O Sistema Não Conseguiu Gravar Por Que Um Determinado Campo Está Vazio"
          Mens = mensd & vbCrLf & Mens
       Case 3421
          Mens = "Voçê Tentou Gravar Uma Informação Diferente Da Esperada Pelo Sistema! " & vbCrLf & " Ex: O Sistema Espera Um Valor Numerico e Voçê Informa Um Texto"
       Case 440
          Mens = mDes & vbCrLf
          Mens = Mens & "Não Há Registro Disponivel Na Tabela Solicitada"
       Case Else
          Mens = "Uma Operação Inválida Resultou Em Um Erro" & vbCrLf & _
          "SE O ERRO PERSISTIR! ANOTE A MENSAGEM ABAIXO E CHAME O PROGRAMADOR" & vbCrLf & _
          "Descrição: " & mDes & vbCrLf & _
          "Número   : " & mErr & vbCrLf & _
          "Dll      : " & mDll
       End Select
       Mens = Mens & vbCrLf & mDll
       On Error Resume Next
       MeuArq = App.Path & "\MeusErros.Txt"
       
       Open MeuArq For Append As #2
    
       Write #2, Err.Source & "->"; Date & "-" & Mid(Time, 1, 5) & "-" & strUsuario & "-" & FrmErr & "-" & mErr & "-" & mDes; "-" & mDll
       Close #2
       
       Mensagem Mens
       
       Mens = "Num Erro : " & mErr & vbCrLf & _
              "Descricao: " & mDes & vbCrLf & _
              "Origem   : " & mDll & vbCrLf & _
              "Formul.  : " & FrmErr & vbCrLf & _
              "Usuario  : " & strUsuario & vbCrLf & _
              "Maquina  : " & GetComputador & vbCrLf & _
              "Data/Hora: " & Date & " " & Mid(Time, 1, 5)
       
     '  Call Enviar_Erro_Email(Mens)
       
       
End Sub

Sub Mensagem(Mens As String)
    Util.Mensagem Mens
End Sub

Public Function NumeroSQL(Valor)
    Dim ValorRetorno
    ValorRetorno = CStr(Valor)
    ValorRetorno = Replace(ValorRetorno, ".", "")
    ValorRetorno = Replace(ValorRetorno, ",", ".")
    NumeroSQL = ValorRetorno
End Function

Function montaSqlWhere(InstrucaoSql As String, j As Collection) As String
Dim SqlJoin As String
Dim Initial As Boolean
SqlJoin = ""
Initial = True
    For M = 1 To j.Count
        If j.Item(M) <> "" Then
           If Initial Then
              SqlJoin = SqlJoin & j.Item(M)
              Initial = False
           Else
              SqlJoin = SqlJoin & " And " & j.Item(M)
           End If
        End If
    Next
    montaSqlWhere = InstrucaoSql & IIf(Len(SqlJoin) = 0, "", " WHERE " & SqlJoin)
    
End Function

Function montaClausulaSqlWhere(j As Collection) As String

Dim SqlJoin As String
Dim Initial As Boolean
SqlJoin = ""
Initial = True
    For M = 1 To j.Count
        If j.Item(M) <> "" Then
           If Initial Then
              SqlJoin = SqlJoin & j.Item(M)
              Initial = False
           Else
              SqlJoin = SqlJoin & " And " & j.Item(M)
           End If
        End If
    Next
    montaClausulaSqlWhere = SqlJoin
    
End Function
Function DtoSQL(data As Variant) As Variant

    If Not IsDate(data) Then
       DtoSQL = Null
    Else
       DtoSQL = Format(data, "dd/mm/yyyy")
    End If

End Function
