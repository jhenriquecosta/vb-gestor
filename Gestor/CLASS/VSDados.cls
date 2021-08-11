VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public ModoTexto As Boolean 'SQz (22/09/02) : vide descricao da propriedade
Attribute ModoTexto.VB_VarDescription = "Habilita/Desabilita a execucao dos comandos SQL.Usado para capturar somente a query gerada."
Public Conexao As VSClass.VSConexao
Public Tabela As VSRecordset
Private Util As New VSUtil
Private Edita As New VSTexto
Private strUltimaQuery As String
Private cteSeparador As String

Public Enum TipoConversao
    TCBit 'Integer data with either a 1 or 0 value.
    TCSimples 'int Integer (whole number) data from -2^31 (-2,147,483,648) through 2^31 - 1 (2,147,483,647).
    TCInteiro 'SmallInt Integer data from 2^15 (-32,768) through 2^15 - 1 (32,767).
    TCByte 'tinyint Integer data from 0 through 255.
    'TCdecimal 'Fixed precision and scale numeric data from -10^38 -1 through 10^38 -1.
    'TCNumeric 'A synonym for decimal.
    TCMonetario '(monetario)Monetary data values from -2^63 (-922,337,203,685,477.5808) through 2^63 - 1 (+922,337,203,685,477.5807), with accuracy to a ten-thousandth of a monetary unit.
    'TCsmallmoney 'Monetary data values from -214,748.3648 through +214,748.3647, with accuracy to a ten-thousandth of a monetary unit.
    'TCFloat 'Floating precision number data from -1.79E + 308 through 1.79E + 308.
    TCDuplo 'Real Floating precision number data from -3.40E + 38 through 3.40E + 38.
    TCDataHora '(data)Date and time data from January 1, 1753, to December 31, 9999, with an accuracy of three-hundredths of a second, or 3.33 milliseconds.
    'TCsmalldatetime 'Date and time data from January 1, 1900, through June 6, 2079, with an accuracy of one minute.
    'TCcursor 'A reference to a cursor.
    'TCTimeStamp 'A database-wide unique number.
    'TCuniqueidentifier 'A globally unique identifier (GUID).
    tctexto 'Char Fixed-length non-Unicode character data with a maximum length of 8,000 characters.
    TCMemo 'VarChar Variable-length non-Unicode data with a maximum of 8,000 characters.
    TCLob 'Text Variable-length non-Unicode data with a maximum length of 2^31 - 1 (2,147,483,647) characters.
    'TCnchar 'Fixed-length Unicode data with a maximum length of 4,000 characters.
    'TCnvarchar 'Variable-length Unicode data with a maximum length of 4,000 characters. sysname is a system-supplied user-defined data type that is a synonym for nvarchar(128) and is used to reference database object names.
    'TCntext 'Variable-length Unicode data with a maximum length of 2^30 - 1 (1,073,741,823) characters.
    'TCBinary 'Fixed-length binary data with a maximum length of 8,000 bytes.
    TCBinario 'Variable-length binary data with a maximum length of 8,000 bytes.
    'TCImage 'Variable-length binary data with a maximum length of 2^31 - 1 (2,147,483,647) bytes.
End Enum
Public Enum TipoCursor
    SomenteAvanco = 0
    Registros = 1
    Dinamico = 2
    Estatico = 3
End Enum
Public Enum TipoTrava
    SomenteLeitura = 1
    Pessimista = 2
    Otimista = 3
    OtimistaBatch = 3
End Enum
Public Enum TipoParteTexto
    LeftVs = 1
    MidVs = 2
    RightVs = 3
End Enum

Public Sub AbreTrans()
    If bRegistrado Then Conexao.BeginTrans
End Sub

Public Sub GravaTrans()
    If bRegistrado Then Conexao.CommitTrans
End Sub

Public Sub CancelaTrans()
    If bRegistrado Then Conexao.RollbackTrans
End Sub

Public Function AbreBanco(Tipo As TipoBanco, DataSource As String, User As String, Optional Password As String, Optional Parametro_Opcional As String) As Boolean
    On Error GoTo Trata
    If bRegistrado Then
        Dim Str As String
        
        Select Case Tipo
            Case Access
                Str = "Provider=Microsoft.Jet.OLEDB.4.0;Mode=ReadWrite;Persist Security Info=True;Password=''" & _
                ";Data Source=" & DataSource & _
                ";User ID=" & User
                
                If Password <> "" Then Str = Str & ";Jet OLEDB:Database Password=" & Password
                
            Case SQLServer
                Str = "Provider=SQLOLEDB.1;Persist Security Info=True" & _
                ";Data Source=" & DataSource & _
                ";User ID=" & User
                
                If Password <> "" Then Str = Str & ";Password=" & Password
                If Parametro_Opcional <> "" Then Str = Str & ";Initial Catalog=" & Parametro_Opcional
                
            Case oracle
                Str = "Provider=MSDAORA.1;Persist Security Info=True;" & _
                ";Data Source=" & DataSource & _
                ";User ID=" & User
                
                If Password <> "" Then Str = Str & ";Password=" & Password
            Case Interbase
                Str = "Provider=IbOleDb.1;Persist Security Info=True;Mode=ReadWrite" & _
                ";Data Source=" & DataSource & _
                ";User ID=" & User
                
                If Password <> "" Then Str = Str & ";Password=" & Password
                If Parametro_Opcional <> "" Then Str = Str & ";Loation=" & Parametro_Opcional
                
        End Select
        FechaBanco
        Set Conexao = New VSConexao
        Conexao.Abrir Str
        
        Conexao.Dsn = DataSource
        Conexao.User = User
        Conexao.Password = Password
        Conexao.Catalog = Parametro_Opcional
        Conexao.FormatoBanco = Tipo
        
        AbreBanco = True
    End If
    Exit Function
Trata:
    If Conexao.Errors.Count > 0 Then
        AbreBanco = False
        Util.Avisa "Erro: " & Conexao.Errors(0).Number & " - " & Conexao.Errors(0).Description & "."
        Screen.MousePointer = 0
    End If
End Function

Public Function AbreBancoDados(Banco As String) As Boolean
    On Error GoTo Trata
    Dim Tipo As TipoBanco, DataSource As String, User As String, Password As String, Parametro_Opcional As String
    If bRegistrado Then
        Dim Str As String
        Dim Instala As New VSInstala
        
        Tipo = Instala.PegaConfiguracao(Banco, tcTipo)
        DataSource = Instala.PegaConfiguracao(Banco, tcDsn)
        User = Instala.PegaConfiguracao(Banco, tcUser)
        Parametro_Opcional = Instala.PegaConfiguracao(Banco, tcCatalog)
        Password = Instala.PegaConfiguracao(Banco, tcPassword)
                    
        Select Case Tipo
            Case Access
                Str = "Provider=Microsoft.Jet.OLEDB.4.0;Mode=ReadWrite;Persist Security Info=True;Password=''" & _
                ";Data Source=" & DataSource & _
                ";User ID=" & User
                
                If Password <> "" Then Str = Str & ";Jet OLEDB:Database Password=" & Password
                
            Case SQLServer
                Str = "Provider=SQLOLEDB.1;Persist Security Info=True" & _
                ";Data Source=" & DataSource & _
                ";User ID=" & User
                
                If Password <> "" Then Str = Str & ";Password=" & Password
                If Parametro_Opcional <> "" Then Str = Str & ";Initial Catalog=" & Parametro_Opcional
                
            Case oracle
                Str = "Provider=MSDAORA.1;Persist Security Info=True;" & _
                ";Data Source=" & DataSource & _
                ";User ID=" & User
                
                If Password <> "" Then Str = Str & ";Password=" & Password
            Case Interbase
                Str = "Provider=IbOleDb.1;Persist Security Info=True;Mode=ReadWrite" & _
                ";Data Source=" & DataSource & _
                ";User ID=" & User
                
                If Password <> "" Then Str = Str & ";Password=" & Password
                If Parametro_Opcional <> "" Then Str = Str & ";Loation=" & Parametro_Opcional
            Case 4
                Str = "PROVIDER=PostgreSQL OLE DB Provider;DATA SOURCE=" & DataSource & "; USER ID= " & User & _
                    "; PASSWORD=" & Password & ";"
        End Select
        FechaBanco
        Set Conexao = New VSConexao
        Conexao.Abrir Str
        
        Conexao.Dsn = DataSource
        Conexao.User = User
        Conexao.Password = Password
        Conexao.Catalog = Parametro_Opcional
        Conexao.FormatoBanco = Tipo
        
        AbreBancoDados = True
    End If
    Exit Function
Trata:
    If Conexao.Errors.Count > 0 Then
        AbreBancoDados = False
        Util.Avisa "Erro: " & Conexao.Errors(0).Number & " - " & Conexao.Errors(0).Description & "."
        Screen.MousePointer = 0
    End If
End Function

Public Function Executa(Sql As String) As Boolean
    On Error GoTo Trata
    
    If bRegistrado Then
        If Not ModoTexto Then
            Conexao.Execute Sql
        End If
        
        Executa = True
        strUltimaQuery = Sql
    End If
    Exit Function
Trata:
    If Conexao.Errors.Count > 0 Then
        Executa = False
        Util.Erro "Erro: " & Conexao.Errors(0).Number & " - " & Conexao.Errors(0).Description & "."
        Screen.MousePointer = 0
    End If
End Function

Public Function AbreTabela(Sql As String, Optional Record As Variant, Optional TipoAbertura As TipoCursor = SomenteAvanco, Optional TipoTrava As TipoTrava = SomenteLeitura) As Boolean
    On Error GoTo Trata
    
    If bRegistrado Then
        If Not IsMissing(Record) Then
            FechaTabela Record
            Set Record = New VSRecordset
            Record.Abrir Sql, Conexao, TipoAbertura, TipoTrava
            If Not Record.Eof Then AbreTabela = True
        Else
            FechaTabela Tabela
            Set Tabela = New VSRecordset
            Tabela.Abrir Sql, Conexao, TipoAbertura, TipoTrava
            If Not Tabela.Eof Then AbreTabela = True
        End If
        strUltimaQuery = Sql
    End If
    Exit Function
    
Trata:
    If Conexao.Errors.Count <> 0 Then
        AbreTabela = False
        Util.Avisa "Erro: " & Conexao.Errors(0).Number & " - " & Conexao.Errors(0).Description & "."
        Screen.MousePointer = 0
    End If
End Function

Public Sub ApagaTabela(Tabela As String)
    On Error GoTo Trata
    
    If bRegistrado Then
        Executa "DROP TABLE " & Tabela
        strUltimaQuery = "DROP TABLE " & Tabela
    End If
    Exit Sub

Trata:
    If Conexao.Errors.Count > 0 Then
        Util.Avisa "Erro: " & Conexao.Errors(0).Number & " - " & Conexao.Errors(0).Description & "."
        Screen.MousePointer = 0
    End If
End Sub

Public Function InsereDados(Tabela As String, Valores As String, Optional Campos As String) As Boolean
' VSClass.VSDados.Function InsereDados
'================================================================================
' Queiroz em VTDES_01
' 24/05/2002-12:11:28
'
' Descricao  : Gera uma sql do tipo Insert, e envia
'
' Parametros : Tabela (String)
'              Valores (String)
'              Campos (String)
'
' Ex:
'--------------------------------------------------------------------------------
    Dim Sql As String
    Dim i As Integer
    Dim Valor As String
    Dim ListaDeValores As String
    Dim PosCaracter As Integer
    Dim ValorData As String
    Dim X As Integer
    On Error GoTo Trata

    If bRegistrado Then
        Sql = "INSERT into " & Tabela & " " & IIf(Trim(Campos) = "", "", "(" & Campos & ")")
        Sql = Sql & " VALUES("
        
        i = 1
        X = Len(Valores)
        Do
            Valor = Util.ParseString(Valores, cteSeparador, i)
'            If (UCase(Left(Trim(Valor), 8)) = "CONVERT(" And Right(Trim(Valor), 1) = ")") Then
            If (Conexao.FormatoBanco = SQLServer And (UCase(Left(Trim(Valor), 8)) = "CONVERT(" And Right(Trim(Valor), 1) = ")")) Or (Conexao.FormatoBanco = oracle And (UCase(Left(Trim(Valor), 3)) = "TO_" And Right(Trim(Valor), 1) = ")")) Then
                Sql = Sql & IIf(i = 1, "", ",") & IIf(Edita.PosPic(Valor, "'") = 0, Left(Left(Valor, Edita.PosPic(Valor, ",")) & "'" & Mid(Valor, Edita.PosPic(Valor, ",") + 1), Len(Valor)) & "')", Valor)
            Else
                Sql = Sql & IIf(i = 1, "", ",") & FormataValorCampo(Valor)
            End If
            i = i + 1
            X = X - (Len(Valor) + Len(cteSeparador))
        Loop Until X = 0
        Sql = Sql & ")"
        
        InsereDados = Executa(Sql)
        strUltimaQuery = Sql
    End If
    Exit Function
Trata:
    If Err.Number = 6 Then
        Resume Next
    End If
End Function

Public Function DeletaDados(Tabela As String, Optional Condicao As String) As Boolean
    Dim Sql As String
    
    If bRegistrado Then
        Sql = "DELETE  from " & Tabela & " "
        If Trim(Condicao) <> "" Then Sql = Sql & " WHERE " & Condicao
    
        DeletaDados = Executa(Sql)
        strUltimaQuery = Sql
    End If
End Function

Public Function AtualizaDados(Tabela As String, Valores As String, Campos As String, Optional Condicao As String) As Boolean
' VSClass.VSDados.Function AtualizaDados
'================================================================================
' Queiroz em VTDES_01
' 24/05/2002-14:34:11
'
' Descricao  : Prepara uma sql do tipo UPDATE e envia para o banco. O valor de retorno
'               indica sucesso ou fracasso da execucao.
'
' Parametros : Tabela (String) - Nome da tabela
'              Valores (String) - Valores a serem atualizados (preparada pela PreparaValores)
'              Campos (String) - Nomes dos campos, separados por virgula
'              Condicao (String) - Condicao de atualizacao (clausula Where)
'
' Ex:
'--------------------------------------------------------------------------------
    On Error GoTo ErroAtualiza
    Dim Sql As String
    Dim i As Integer
    Dim Campo As String, Valor As String
    Dim ListaDeValores As String
    Dim PosCaracter As Integer
    Dim ValorData As String
    
    If bRegistrado Then
        Sql = "UPDATE " & Tabela & " SET "
        
        i = 1
        Do
            Campo = Util.ParseString(Campos, ",", i)
            If Campo <> "" Then
                Sql = Sql & IIf(i = 1, "", ",") & Campo & " = "
                Valor = Util.ParseString(Valores, cteSeparador, i)
                If UCase(Left(Trim(Valor), 8)) = "CONVERT(" And Right(Trim(Valor), 1) = ")" Then
                    Sql = Sql & IIf(Edita.PosPic(Valor, "'") = 0, Left(Left(Valor, _
                        Edita.PosPic(Valor, ",")) & "'" & Mid(Valor, Edita.PosPic(Valor, ",") + 1), _
                        Len(Valor)) & "')", Valor)
                Else
                    Sql = Sql & FormataValorCampo(Valor)
                End If
                i = i + 1
            Else
                Exit Do
            End If
        Loop Until Campo = ""
        
        If Condicao <> "" Then Sql = Sql & " WHERE " & Condicao
    
        AtualizaDados = Executa(Sql)
        strUltimaQuery = Sql
    End If
    Exit Function

ErroAtualiza:
    MsgBox Err.Description
End Function

Public Function GravaDados(Tabela As String, Valores As String, Campos As String, Condicao As String) As Boolean
    Dim RSTEMP As VSRecordset
    If bRegistrado Then
        If AbreTabela("SELECT * FROM " & Tabela & " WHERE " & Condicao, RSTEMP) Then
            GravaDados = AtualizaDados(Tabela, Valores, Campos, Condicao)
        Else
            GravaDados = InsereDados(Tabela, Valores, Campos)
        End If
        FechaTabela RSTEMP
    End If
End Function

Public Sub FechaTabela(Optional Record As Variant)
    If bRegistrado Then
        If Not IsMissing(Record) Then
            If Not Record Is Nothing Then
                If Record.State <> adStateClosed Then
                    Record.Fechar
                    Set Record = Nothing
                End If
            End If
        Else
            If Not Tabela Is Nothing Then
                If Tabela.State <> adStateClosed Then
                    Tabela.Fechar
                    Set Tabela = Nothing
                End If
            End If
        End If
    End If
End Sub

Public Function PreparaValor(ParamArray Valores()) As String
    Dim i As Byte, Min As Byte, Max As Byte
    Dim Valor As String
    
    Valor = ""
    If bRegistrado Then
        Min = LBound(Valores)
        Max = UBound(Valores)
        For i = Min To Max
            Valor = Valor & Util.Nvl(CStr(Valores(i)), "Null") & cteSeparador
'            If CStr(Valores(i)) <> "" Then
'                Valor = Valor & Seg.Criptografa(Seg.Criptografa(Seg.Criptografa(CStr(Valores(i))))) & cteSeparador
'            Else
'                Valor = Valor & "Null" & cteSeparador
'            End If
        Next
        PreparaValor = Valor
    End If
End Function

Public Sub FechaBanco()
    If bRegistrado Then
        If Not Conexao Is Nothing Then
            If Conexao.State <> adStateClosed Then
                Conexao.Fechar
                Set Conexao = Nothing
            End If
        End If
    End If
End Sub

Public Function BuscaCodigo(Tabela As String) As String
    Dim RS As VSRecordset
    
    If bRegistrado Then
        If AbreTabela(Tabela, RS) Then
            BuscaCodigo = IIf(IsNull(RS(0)), "0", RS(0))
        Else
            BuscaCodigo = "0"
        End If
        FechaTabela RS
        strUltimaQuery = Tabela
    End If
End Function

Public Function DescricaoGeral(Tabela As String, Codigo As Long) As String
    Dim RS As VSRecordset
    Dim strSql As String
    
    If bRegistrado Then
        strSql = "SELECT TGE_NOME FROM TAB_GERAL WHERE TGE_CODIGO=" & Codigo & " AND TGE_TIPO=" & _
                    "(SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_CODIGO=0 AND TGE_NOME='" & Tabela & "')"
        
        If AbreTabela(strSql, RS) Then
            DescricaoGeral = RS!TGE_NOME
        End If
        FechaTabela RS
        strUltimaQuery = strSql
    End If
End Function

Public Function CodigoGeral(Tabela As String, Descricao As String) As Long
    Dim RS As VSRecordset
    Dim strSql As String
    
    If bRegistrado Then
        strSql = "SELECT TGE_CODIGO FROM TAB_GERAL WHERE TGE_NOME='" & Descricao & "' AND TGE_TIPO=" & _
                    "(SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_CODIGO=0 AND TGE_NOME='" & Tabela & "')"
        
        If AbreTabela(strSql, RS) Then
            CodigoGeral = RS!TGE_CODIGO
        End If
        FechaTabela RS
        strUltimaQuery = strSql
    End If
End Function

Private Sub Class_Initialize()
    ValidaComponente "CLASS"
    If bRegistrado Then
        Set Conexao = New VSConexao
        cteSeparador = "VTSEP"
    End If
End Sub

Public Function Concatena() As String
    If bRegistrado Then
        Select Case Conexao.FormatoBanco
            Case Access
                Concatena = " & "
            Case SQLServer
                Concatena = " + "
            Case oracle
                Concatena = " || "
            Case Interbase
                Concatena = " || "
        End Select
    End If
End Function

Public Function Converte(ByVal Dado, Tipo As TipoConversao) As String
    If bRegistrado Then
        Dim T As New VSTexto
        Dim pos As Integer
        Select Case Conexao.FormatoBanco
            Case Access
                Select Case Tipo
                    Case TCDataHora
                        If InStr(1, Dado, ",103") > 0 Then
                            Dado = Mid(Dado, 1, Len(Dado) - 4)
                        End If
                        If IsDate(Dado) Then
                            Converte = " cdate('" & Dado & "') "
                        Else
                            Converte = " cdate(" & Dado & ") "
                        End If
                    Case TCLob, TCMemo, tctexto: If Trim(Dado) <> "" Then Converte = " CSTR('" & Dado & "') "
                    Case TCMonetario: Converte = " ccur('" & Dado & "') "
                    Case TCByte: Converte = " cbyte('" & Dado & "') "
                    Case TCInteiro: Converte = " cint('" & Dado & "') "
                    Case TCDuplo, TCSimples: Converte = " cdbl('" & Dado & "') "
                    'Case TCDuplo, TCSimples: Dado = Format(Dado, "#,##0.00"): Converte = " cdbl(" & Dado & ") "
                    Case Else
                        Call Util.Erro("Tipo de dados não programado.")
                End Select
            Case SQLServer
                    If IsNumeric(Dado) Then
                        Dim ValorAux As String
                        ValorAux = Edita.TiraPic(CStr(Dado), ".")
                        ValorAux = Edita.TrocaPic(CStr(ValorAux), ",", ".")
                        Converte = " convert(" & NomeConvSQL(Tipo) & "," & IIf(Tipo = tctexto, "'" & ValorAux & "'", ValorAux) & ")"
                    Else
'                        If Trim(Dado) = "" Then Dado = "Null"
                        If Tipo = tctexto Then
                            If Trim(Dado) <> "" Then Converte = " convert(" & NomeConvSQL(Tipo) & "(" & Len(Dado) & "),'" & Dado & "')"
                        Else
                            If Trim(Dado) <> "" Then Converte = " convert(" & NomeConvSQL(Tipo) & "," & IIf(Tipo = TCDataHora, " '" & Dado & "',103", Dado) & ")"
                        End If
                    End If
            Case oracle
                If Tipo = TCDataHora Then
                    Converte = " to_date('" & Dado & "','dd/mm/yyyy') "
                ElseIf Tipo = tctexto Then
                    Converte = " to_char('" & Dado & "') "
                ElseIf Tipo = TCMemo Then
                    Converte = " to_long('" & Dado & "') "
                Else
                    If Tipo = TCMonetario Then
                        Dado = Format(Dado, "#,##0.00")
                    End If
                    If Fix(Dado) > 0 Then
                        pos = InStr(Dado, ",")
                        Dado = T.TiraTudo(CStr(Dado))
                        If pos > 0 Then
                            Dado = Left(Dado, Len(Dado) - 2) & "," & Right(Dado, 2)
                        End If
                    End If
                    If Tipo = TCMonetario Then
                        Converte = " to_number('" & T.TrocaPic(CStr(Dado), ",", ".") & "','9999999999.99')"
                    Else
                        Converte = " to_number('" & T.TrocaPic(CStr(Dado), ",", ".") & "','9999999999.9999')"
                    End If
                End If
            Case Interbase
                If Tipo = TCDataHora Then
                    Converte = " to_date(" & Dado & ",'dd/mm/yyyy') "
                ElseIf Tipo = tctexto Then
                    Converte = " to_char(" & Dado & ") "
                ElseIf Tipo = TCMemo Then
                    Converte = " to_long(" & Dado & ") "
                Else
                    Converte = " to_number(" & Dado & ") "
                End If
        End Select
    End If
End Function

Private Function NomeConvSQL(EnumConv As TipoConversao) As String
    If bRegistrado Then
        Select Case EnumConv
            Case TCBit: NomeConvSQL = "Bit"
            Case TCSimples: NomeConvSQL = "int"
            Case TCInteiro: NomeConvSQL = "SmallInt"
            Case TCByte: NomeConvSQL = "tinyint"
            Case TCMonetario: NomeConvSQL = "Money"
            Case TCDuplo: NomeConvSQL = "Real"
            Case TCDataHora: NomeConvSQL = "DateTime"
            Case tctexto: NomeConvSQL = "Char"
            Case TCMemo: NomeConvSQL = "VarChar"
            Case TCLob: NomeConvSQL = "Text"
            Case TCBinario: NomeConvSQL = "VarBinary"
        End Select
    End If
End Function

Public Function ParteTexto(Dado, Parte As TipoParteTexto, Optional Inicio As Integer = 1, Optional Tamanho As Integer = 0, Optional DadoEhCampo As Boolean = False) As String
    If bRegistrado Then
        Select Case Conexao.FormatoBanco
            Case Access
            
                If Parte = LeftVs Then
                    ParteTexto = " Left(" & Dado & "," & Tamanho & ")"
                ElseIf Parte = MidVs Then
                    ParteTexto = " Mid(" & Dado & "," & Inicio & IIf(Tamanho = 0, ")", "," & Tamanho & ")")
                ElseIf Parte = RightVs Then
                    ParteTexto = " Right(" & Dado & "," & Tamanho & ")"
                End If
            Case SQLServer
                If Parte = LeftVs Then
                    ParteTexto = " Left(" & Dado & "," & Tamanho & ")"
                ElseIf Parte = MidVs Then
                '<Removed by: Queiroz at: 28/05/2002-20:08:16 on machine: VTDES_01>
    '                ParteTexto = " Left(Right(" & Dado & "," & Len(Dado) + 1 - Inicio & ")," & IIf(Tamanho <> 0, Tamanho, Len(Dado) + 1 - Inicio) & ")"
                '</Removed by: Queiroz at: 28/05/2002-20:08:16 on machine: VTDES_01>
                
                '<Added by: Queiroz at: 28/05/2002-20:12:39 on machine: VTDES_01>
                    ParteTexto = " Substring(Cast(" & IIf(DadoEhCampo, "", "'") & Dado & IIf(DadoEhCampo, "", "'") & " AS varchar)," & Inicio & "," & Tamanho & ")"
                '</Added by: Queiroz at: 28/05/2002-20:12:39 on machine: VTDES_01>
                ElseIf Parte = RightVs Then
                    ParteTexto = " Right(" & Dado & "," & Tamanho & ")"
                End If
            Case oracle
                If Parte = LeftVs Then
                    ParteTexto = " substr(" & Dado & "," & Tamanho & ")"
                ElseIf Parte = MidVs Then
                    ParteTexto = " substr(" & Dado & "," & Inicio & IIf(Tamanho = 0, ")", "," & Tamanho & ")")
                ElseIf Parte = RightVs Then
                    ParteTexto = " right(" & Dado & "," & Tamanho & ")" 'Falta este
                End If
            Case Interbase
                If Parte = LeftVs Then
                    ParteTexto = " substr(" & Dado & "," & Tamanho & ")"
                ElseIf Parte = MidVs Then
                    ParteTexto = " substr(" & Dado & "," & Inicio & IIf(Tamanho = 0, ")", "," & Tamanho & ")")
                ElseIf Parte = RightVs Then
                    ParteTexto = " right(" & Dado & "," & Tamanho & ")" 'Falta este
                End If
        End Select
    End If
End Function

Public Property Get UltimoComando() As String
    If bRegistrado Then UltimoComando = strUltimaQuery
End Property

Public Function FormataValorCampo(Valor) As String
' VSClass.VSDados.Function FormataValorCampo
'================================================================================
' Queiroz em VTDES_01
' 24/05/2002-14:01:22
'
' Descricao  : Formata um valor da forma que a sintaxe sql dos bancos solicita
'
' Parametros : Valor - Valor a ser formatado
'
' Ex: FormataValorCampo("27/03/1978") = Convert('27/03/1978', 103, DateTime), num banco SQL Server
'--------------------------------------------------------------------------------
    On Error GoTo ErroFormata
    If bRegistrado Then
        Dim resultado As String
        If Trim$(UCase(Valor)) = "NULL" Then
            resultado = "Null"
        Else
            If IsNumeric(Valor) Then
                If Conexao.FormatoBanco = Access Then
                    'resultado = "'" & Valor & "'"
                    resultado = Valor
                Else
                    resultado = Edita.TiraPic(CStr(Valor), ".")
                    resultado = Edita.TrocaPic(resultado, ",", ".")
                End If
            ElseIf Mid(Valor, 3, 1) = "/" And Mid(Valor, 6, 1) = "/" And Len(Valor) = 10 Then
                resultado = Converte(Valor, TCDataHora)
            Else
                If Conexao.FormatoBanco = oracle Then
                    resultado = CStr(Valor)
                Else
                    resultado = Edita.TiraPic(CStr(Valor), "'")
                End If
                If Conexao.FormatoBanco = Access Then
                    If Not ((UCase(Left(Trim(Valor), 5)) = "CSTR(" Or UCase(Left(Trim(Valor), 5)) = "CDBL(" Or UCase(Left(Trim(Valor), 6)) = "CDATE(") And Right(Trim(Valor), 1) = ")") Then
                        resultado = "'" & resultado & "'"
                    Else
                        resultado = IIf(Edita.PosPic(CStr(Valor), "'") = 0, Left(Left(Valor, Edita.PosPic(CStr(Valor), "(")) & "'" & Mid(Valor, Edita.PosPic(CStr(Valor), "(") + 1), Len(Valor)) & "')", Valor)
                    End If
                ElseIf Conexao.FormatoBanco = SQLServer Then
                    resultado = "'" & resultado & "'"
                ElseIf Conexao.FormatoBanco = oracle Then
                    If Not UCase(Left(Trim(Valor), 3)) = "TO_" Then
                         resultado = Converte(resultado, tctexto)
                    End If
                End If
            End If
        End If
        FormataValorCampo = resultado
    End If
    Exit Function
ErroFormata:
        MsgBox Err.Description
End Function

Public Function AnexarDocumento(Tabela As String, Campos As String, CampoDoc As String, ByRef objRTF As Object, Optional Condicao As String, Optional Valores As String) As Boolean
    Dim Sql As String, Valor As String
    Dim X As Integer
    Dim cmmDocumento As New ADODB.Command
    Dim Param1 As Parameter
    Dim Param2 As Parameter
    Dim i As Integer
    
    On Error GoTo Trata_Erro
    
    AnexarDocumento = True
    
        
    Sql = "INSERT INTO " & Tabela & " " & IIf(Trim(Campos) = "", "", "(" & Campos & ")")
    Sql = Sql & " VALUES("
    
    i = 1
    X = Len(Valores)
    Do
        Valor = Util.ParseString(Valores, cteSeparador, i)
        If UCase(Left(Trim(Valor), 8)) = "CONVERT(" And Right(Trim(Valor), 1) = ")" Then
            Sql = Sql & IIf(i = 1, "", ",") & IIf(Edita.PosPic(Valor, "'") = 0, Left(Left(Valor, Edita.PosPic(Valor, ",")) & "'" & Mid(Valor, Edita.PosPic(Valor, ",") + 1), Len(Valor)) & "')", Valor)
        Else
            Sql = Sql & IIf(i = 1, "", ",") & FormataValorCampo(Valor)
        End If
        i = i + 1
        X = X - (Len(Valor) + Len(cteSeparador))
    Loop Until X = 0
    Sql = Sql & ")"
    
    Sql = "INSERT INTO " & Tabela & "(HRD_HRO_NUM_CORRELATIVO,HRD_ARQUIVO_DOC)"
    Sql = Sql & " VALUES(?,?)"
        
    Set cmmDocumento.ActiveConnection = Me.Conexao
    cmmDocumento.CommandText = Sql
    Set Param1 = New Parameter
    Param1.Direction = adParamInput
    Param1.Type = adNumeric
    'Param1.Value = Num_Regime
    cmmDocumento.Parameters.Append Param1
        
    Set Param2 = New Parameter
    Param2.Direction = adParamInput
    Param2.Type = adBSTR
    'Param2.Value = rct_Documento.TextRTF
    cmmDocumento.Parameters.Append Param2
        
    cmmDocumento.Execute
Trata_Erro:
End Function

Public Function sp_GeraCorrelativo(Banco As String, Operacao As Integer) As String
    Util.Erro "Use a funcao Correlativo!"
'    Dim prBanco As Parameter
'    Dim prOperacao As Parameter
'    Dim prRetorno As Parameter
'    Dim prSequencial As Parameter
'
'    Dim Store As New adodb.Command
'    Dim Retorno As String
'    Set Store.ActiveConnection = Me.Conexao.DBConnection
'
'    Store.CommandText = "USE VTSeg"
'    Store.CommandType = adCmdText
'    Store.Execute
'
'    Store.CommandText = "gp_GeraNumCorrelativo"
'    Store.CommandType = adCmdStoredProc
'
'    ' Seta o parâmetro de retorno padrão
'    Set prRetorno = Store.CreateParameter("Return", adInteger, adParamReturnValue)
'    Store.Parameters.Append prRetorno
'
'    ' Seta o parâmetro de entrada Banco
'    Set prBanco = Store.CreateParameter("NomeDoBanco", adVarChar, adParamInput, 8)
'    Store.Parameters.Append prBanco
'    prBanco.Value = Banco
'
'    ' Seta o parâmetro de entrada Operacao
'    Set prOperacao = Store.CreateParameter("OperadorCorrelativo", adInteger, adParamInput)
'    Store.Parameters.Append prOperacao
'    prOperacao.Value = Operacao
'
'    ' Seta o parâmetro de retorno com  o número sequencial
'    Set prSequencial = Store.CreateParameter("NumCorrelativo", adInteger, adParamOutput)
'    Store.Parameters.Append prSequencial
'
'    Store.Execute
'    sp_GeraCorrelativo = prSequencial.Value
'
'    While Store.Parameters.Count > 0
'        Store.Parameters.Delete 0
'    Wend
'
'
'    Store.CommandText = "USE " & Banco
'    Store.CommandType = adCmdText
'    Store.Execute
End Function

Public Function Correlativo(Sistema As String, Operador As Integer, Optional Descricao As String, _
                            Optional Incrementar As Boolean = True, _
                            Optional Referencia As String, Optional Incremento As Integer = 1, _
                            Optional Inicio As Integer = 1, _
                            Optional Mascara As String = "999999") As String
    Dim Comando As VSComando
    
    Set Comando = New VSComando
    Comando.Texto Me, "vtseg.dbo.sp_num_correlativo", cmdStoredProcedure
    Comando.setarParametro "NomeDoSistema", tipChar, parEntrada, 4, Sistema
    Comando.setarParametro "NumCorrelativo", tipInteger, parEntradaSaida, , Operador
    Comando.setarParametro "ReferenciaDoOperador", tipLongVarChar, parEntrada, 20, Referencia
    Comando.setarParametro "SeqDoOperador", tipInteger, parEntrada, , Inicio
    Comando.setarParametro "IncDaSequencia", tipInteger, parEntrada, , Incremento
    Comando.setarParametro "AutoInc", tipInteger, parEntrada, , Abs(Incrementar)
    Comando.setarParametro "DisplayFormat", tipChar, parEntrada, 6, Mascara
    Comando.setarParametro "DescrDoOperador", tipVarChar, parEntrada, 50, Descricao
    Comando.Executa
    Correlativo = Comando.Parametro("NumCorrelativo").Value
End Function


