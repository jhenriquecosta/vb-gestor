VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSTexto"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Util As New VSUtil

Const Ponto = "."
Const Virgula = ","

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Enum TipoChar
    Letra = 0
    Numero = 1
    Valores = 2
End Enum

Public Enum TipoFormato
    Data = 0
    Cpf = 1
    Cgc = 2
    Telefone = 3
    CEP = 4
    Monetario = 5
    Hora = 6
    UmDV = 7
    DoisDV = 8
    PASEP = 9
    Documento = 10
End Enum

Public Function ListIndexDe(Combo As Object, Texto As String) As Integer
    If bRegistrado Then
        Dim i As Integer
        For i = 0 To Combo.ListCount
            If Combo.List(i) = Texto Then
                ListIndexDe = i
                Exit Function
            End If
        Next
        ListIndexDe = -1
    End If
End Function

Public Function Minuscula(nKeyAscii As Integer) As Integer
    If bRegistrado Then Minuscula = Asc(LCase(Chr(nKeyAscii)))
End Function

Public Function Maiuscula(nKeyAscii As Integer) As Integer
    If bRegistrado Then Maiuscula = Asc(UCase(Chr(nKeyAscii)))
End Function

Public Function AceitaDig(Dig_KeyAscii As Integer, Dig_Tipo As TipoChar) As Integer
    If bRegistrado Then
        'Permite ou nao a digitacao de um caracter
        AceitaDig = Dig_KeyAscii
        If Dig_KeyAscii = vbKeyBack Or Dig_KeyAscii = vbKeySpace Then Exit Function
        Select Case Dig_Tipo
            Case Letra
                ' Permite aceitação de caracter com ou sem acentuação
                If (Dig_KeyAscii < 65 Or Dig_KeyAscii > 90) And (Dig_KeyAscii < 97 Or Dig_KeyAscii > 122) Then
                    If (Dig_KeyAscii < 192 Or Dig_KeyAscii > 197) And (Dig_KeyAscii < 200 Or Dig_KeyAscii > 207) And (Dig_KeyAscii < 210 Or Dig_KeyAscii > 214) And (Dig_KeyAscii < 217 Or Dig_KeyAscii > 220) And (Dig_KeyAscii < 199) Then
                        AceitaDig = 0
                    End If
                End If
            Case Numero
                If Dig_KeyAscii < 48 Or Dig_KeyAscii > 57 Then
                    AceitaDig = 0
                End If
            Case Valores
                If (Dig_KeyAscii < 48 Or Dig_KeyAscii > 57) And Dig_KeyAscii <> 44 Then
                    AceitaDig = 0
                End If
        End Select
    End If
End Function

Public Sub BuscaItemNaLista(EditControl As Object, KeyAscii As Integer)
    If bRegistrado Then
        On Error Resume Next
        Dim buffer As String
        Dim RetVal As Long
        buffer = Left(EditControl.Text, EditControl.SelStart) & Chr(KeyAscii)
        RetVal = SendMessage((EditControl.hWnd), &H14C, -1, ByVal buffer)
        If RetVal <> -1 Then
            EditControl.ListIndex = RetVal
            EditControl.Text = EditControl.List(RetVal)
            EditControl.SelStart = Len(buffer)
            EditControl.SelLength = Len(EditControl.Text)
            KeyAscii = 0
        End If
    End If
End Sub

Public Function CriticaCampos(Form As Object) As Boolean
    If bRegistrado Then
        On Error Resume Next
        Dim Controle As Control
        
        For Each Controle In Form.Controls
            If Trim(Controle.Tag) <> "" And Trim(Controle.Text) = "" Then
                If Err.Number = 0 Then
                    Call Util.Avisa("Campo '" & Controle.Tag & "' deve ser informado.")
                    Controle.SetFocus
                    CriticaCampos = False
                    Exit Function
                End If
                Err.Clear
            End If
        Next
        CriticaCampos = True
    End If
End Function

Public Sub DestacaCaixa(EditControl As Object, Status As Boolean)
    If bRegistrado Then
        'Destaca caixa de texto e combo ao receber o foco
        If Not Status Then
            EditControl.ForeColor = &H800000
            EditControl.BackColor = &H80000005
        Else
            EditControl.SelStart = 0
            EditControl.SelLength = Len(Trim(EditControl))
            EditControl.ForeColor = &HFFFFFF
            EditControl.BackColor = &H800000
        End If
    End If
End Sub

Public Function FormataTexto(ByRef EditControl As Object, Tipo As TipoFormato, Optional Agrupar As Boolean = True) As String
    If bRegistrado Then
        'Mascara texto de acordo com seu tipo
        If Trim$(EditControl) = "" Then Exit Function
        Dim i As Integer, Texto As String
        
        Select Case Tipo
            Case Data
                If Mid(EditControl, 3, 1) = "/" Then
                    FormataTexto = EditControl
                    Exit Function
                End If
                If Len(EditControl) = 6 Then
                    FormataTexto = Left(EditControl, 2) _
                                    & "/" & Mid(EditControl, 3, 2) _
                                    & "/" & Right(EditControl, 2)
                ElseIf Len(EditControl) = 8 Then
                    FormataTexto = Left(EditControl, 2) _
                                    & "/" & Mid(EditControl, 3, 2) _
                                    & "/" & Right(EditControl, 4)
                Else
                    FormataTexto = ""
                End If
                If IsDate(FormataTexto) Then
                    EditControl = FormataTexto
                Else
                    FormataTexto = ""
                    EditControl = ""
                End If
                
            Case Cpf
                If Util.ValidaCpf(EditControl.Text) Then
                    If Len(EditControl) = 11 And Mid(EditControl, 4, 1) <> "." Then
                        FormataTexto = Left(EditControl, 3) & "." & Mid(EditControl, 4, 3) & "." & _
                            Mid(EditControl, 7, 3) & "-" & Right(EditControl, 2)
                    End If
                    EditControl = FormataTexto
                Else
                    FormataTexto = ""
                    EditControl = ""
                End If
                
            Case CEP
                If Len(EditControl) = 8 Then
                    FormataTexto = Mid(EditControl, 1, 5) & "-"
                    FormataTexto = FormataTexto + Mid(EditControl, 6, 3)
                End If
                EditControl = FormataTexto
            
            Case Monetario
                Texto = ""
                If IsNumeric(EditControl.Text) Then
                    i = InStr(1, EditControl.Text, ",")
                    If i > 0 Then
                        Texto = Util.Nvl(Mid(EditControl.Text, 1, i - 1), 0)
                        Texto = Format(Texto, "#,##0")
                        Texto = Texto & "," & Mid(EditControl.Text, i + 1)
                    End If
                    Texto = Format(EditControl.Text, "Standard")
                End If
                FormataTexto = Texto
                EditControl = Texto
               
            Case Cgc
                If Len(EditControl) = 14 And Mid(EditControl, 4, 1) <> "." Then
                    If Mid(EditControl, 3, 1) <> "." Then
                        FormataTexto = Left(EditControl, 2) & "." & Mid(EditControl, 3, 3) & "." & _
                            Mid(EditControl, 6, 3) & "/" & Mid(EditControl, 9, 4) & "-" & _
                            Right(EditControl, 2)
                    End If
                End If
                EditControl = FormataTexto
    '            If Not Util.ValidaCgc(EditControl.Text) Then
    '                FormataTexto = ""
    '                EditControl = ""
    '            End If
                    
            Case Telefone
                If Len(EditControl) = 7 Then
                    FormataTexto = Left(EditControl, 3) & "-" & Mid(EditControl, 4)
                    EditControl = FormataTexto
            
                ElseIf Len(EditControl) = 8 Then
                    FormataTexto = Left(EditControl, 4) & "-" & Mid(EditControl, 5)
                    EditControl = FormataTexto
                    
                ElseIf Len(EditControl) = 10 Then
                    FormataTexto = "(" & Trim$(Left(EditControl, 3)) & ") " & Mid(EditControl, 4, 3) & _
                                    "-" & Mid(EditControl, 7, 4)
                    EditControl = FormataTexto
                    
                ElseIf Len(EditControl) = 11 Then
                        FormataTexto = "(" & Trim$(Left(EditControl, 3)) & ") " & Mid(EditControl, 4, 4) & _
                                        "-" & Mid(EditControl, 8, 4)
                    EditControl = FormataTexto
                ElseIf Len(EditControl) = 13 Then
                    FormataTexto = EditControl
                End If
                
            Case Hora
                If Len(EditControl) > 2 And (Mid(EditControl, 3, 1) = ":" Or Mid(EditControl, 3, 1) > 5) Then
                    FormataTexto = EditControl
                    Exit Function
                ElseIf Len(EditControl) = 1 Then
                    FormataTexto = "0" & Left(EditControl, 2) & ":00"
                    EditControl = FormataTexto
                ElseIf Len(EditControl) = 2 Then
                    FormataTexto = Left(EditControl, 2) & ":00"
                    EditControl = FormataTexto
                ElseIf Len(EditControl) = 4 Then
                    FormataTexto = Left(EditControl, 2) & ":" & Right(EditControl, 2)
                    EditControl = FormataTexto
                ElseIf Len(EditControl) = 6 Then
                    FormataTexto = Left(EditControl, 2) _
                                    & ":" & Mid(EditControl, 3, 2) _
                                    & ":" & Right(EditControl, 4)
                    EditControl = FormataTexto
                Else
                    FormataTexto = EditControl
                End If
                
            Case UmDV
                Dim Str As String
                If Agrupar Then
                    Texto = "#,##0"
                Else
                    Texto = "0"
                End If
                Str = Mid$(EditControl, 1, Len(EditControl) - 1)
                Str = Format(Str, Texto)
                EditControl = Str & "-" & Mid$(EditControl, Len(EditControl))
                FormataTexto = EditControl
                
            Case DoisDV
                If Agrupar Then
                    Texto = "#,##0"
                Else
                    Texto = "0"
                End If
                Str = Mid$(EditControl, 1, Len(EditControl) - 2)
                Str = Format(Str, Texto)
                EditControl = Str & "-" & Mid$(EditControl, Len(EditControl) - 1)
                FormataTexto = EditControl
                
            Case PASEP
                Str = BotaPic(EditControl.Text, ".", 3)
                Str = BotaPic(Str, ".", 9)
                Str = BotaPic(Str, ".", 12)
                EditControl = Str
                FormataTexto = EditControl
                
            Case Documento
                Select Case Len(EditControl.Text)
                    Case 14
                        EditControl.Text = FormataTexto(EditControl, Cgc, Agrupar)
                    Case 11
                        EditControl.Text = FormataTexto(EditControl, Cpf, Agrupar)
                    Case Else
                        EditControl.Text = FormataTexto(EditControl, UmDV, Agrupar)
                End Select
                FormataTexto = EditControl.Text
                
        End Select
    End If
End Function

Public Sub LimpaCampos(Form_Name As Object)
    If bRegistrado Then
        Dim Controle As Control
        On Error Resume Next
        For Each Controle In Form_Name.Controls
            Controle.Text = ""
            If Controle.Style <> 0 Then Controle.ListIndex = -1
        Next
    End If
End Sub

Public Sub HabilitaCampos(Form_Name As Object, Valor As Boolean)
    If bRegistrado Then
        Dim Controle As Control
        On Error Resume Next
        For Each Controle In Form_Name.Controls
            If Not TypeOf Controle Is Label Then
                Controle.Enabled = Valor
            End If
        Next
    End If
End Sub

Public Sub SelecionaTexto(txt As Object)
    If bRegistrado Then
        txt.SelStart = 0
        txt.SelLength = Len(txt)
    End If
End Sub

Public Sub FocalizaCaixa(Form As Object)
    On Error GoTo Trata
    If bRegistrado Then
        Static Objeto As Object
        If Not IsMissing(Form.ActiveControl.Text) Then
            If Objeto Is Nothing Then
                Set Objeto = Form.ActiveControl
                Call DestacaCaixa(Objeto, True)
            ElseIf Form.ActiveControl.Name <> Objeto.Name Then
                Call DestacaCaixa(Objeto, False)
                Call DestacaCaixa(Form.ActiveControl, True)
                Set Objeto = Form.ActiveControl
            End If
        ElseIf Not (Objeto Is Nothing) Then
            If Form.ActiveControl.Name <> Objeto.Name Then
                Call DestacaCaixa(Objeto, False)
                Set Objeto = Form.ActiveControl
            End If
        End If
    End If
Trata:
    Exit Sub
End Sub

Public Sub AtualizaCombo(Bdados As Object, Combo As Object, Tabela As String)
    If bRegistrado Then
        Dim RS As VSRecordset
        Combo.Clear
        If Bdados.AbreTabela(Tabela, RS) Then
            Do Until RS.Eof
                If Not IsNull(RS(0)) Then
                    If Trim(RS(0)) <> "" Then
                        Combo.AddItem RS(0)
                    End If
                End If
                RS.MoveNext
            Loop
        End If
    End If
End Sub

Public Sub AtualizaComboGeral(Bdados As Object, Combo As Object, Tabela As String)
    If bRegistrado Then
        Dim RS As VSRecordset
        Dim strSql As String
        
        Combo.Clear
        strSql = "SELECT TGE_NOME, TGE_CODIGO FROM TAB_GERAL WHERE TGE_CODIGO>0 AND TGE_TIPO=" & _
                    "(SELECT TGE_TIPO FROM TAB_GERAL WHERE TGE_CODIGO=0 AND TGE_NOME='" & Tabela & "')"
        If Bdados.AbreTabela(strSql, RS) Then
            Do Until RS.Eof
                If Not IsNull(RS!TGE_NOME) Then
                    If Trim(RS!TGE_NOME) <> "" Then
                        Combo.AddItem RS!TGE_NOME
                        Combo.ItemData(Combo.NewIndex) = RS!TGE_CODIGO
                    End If
                End If
                RS.MoveNext
            Loop
        End If
        Bdados.FechaTabela RS
    End If
End Sub

Public Function PassaTamanhoCombo(Combo As Object, MaxLength As Integer) As Boolean
    If bRegistrado Then
        If Len(Combo.Text) > MaxLength Then
            Util.Avisa "Campo com tamanho máximo de " & MaxLength & " caracteres."
            Combo.SetFocus
            PassaTamanhoCombo = True
        End If
    End If
End Function

Public Function TiraPic(cString As String, cChar As String) As String
    If bRegistrado Then
        Dim nPoint As Integer
         If Len(cChar) = 1 Then
              nPoint = InStr(cString, cChar)
              While nPoint
                   cString = Left(cString, nPoint - 1) & Mid$(cString, nPoint + 1, Len(cString))
                   nPoint = InStr(nPoint, cString, cChar)
              Wend
         End If
         TiraPic = cString
    End If
End Function

Public Function TiraTudo(strValor As String) As String
    Dim strRetorno As String
    
    strRetorno = TiraPic(strValor, ".")
    strRetorno = TiraPic(strRetorno, ",")
    strRetorno = TiraPic(strRetorno, "/")
    strRetorno = TiraPic(strRetorno, "\")
    strRetorno = TiraPic(strRetorno, "-")
    strRetorno = TiraPic(strRetorno, ":")
    strRetorno = TiraPic(strRetorno, ";")
    strRetorno = TiraPic(strRetorno, "(")
    strRetorno = TiraPic(strRetorno, ")")
    TiraTudo = strRetorno
End Function

Public Function BotaPic(cString As String, cChar As String, nPos As Integer) As String
    If bRegistrado Then
        If Len(cString) > nPos Then
            BotaPic = Mid$(cString, 1, nPos) & cChar & Mid$(cString, nPos + 1)
        Else
            BotaPic = cString
        End If
    End If
End Function

Public Function PosPic(cString As String, cChar As String) As Integer
    If bRegistrado Then PosPic = InStr(1, cString, cChar)
End Function

Public Function TrocaPic(cString As String, cChar As String, Troca As String) As String
    If bRegistrado Then
        Dim nPoint As Integer
         If Len(cChar) = 1 And Len(Troca) = 1 Then
              nPoint = InStr(cString, cChar)
              While nPoint
                   cString = Left(cString, nPoint - 1) & Troca & Mid$(cString, nPoint + 1, Len(cString))
                   nPoint = InStr(nPoint, cString, cChar)
              Wend
         End If
        TrocaPic = cString
    End If
End Function

Public Function BuscaItemListView(EditControl As Object, Coluna As Integer, item As String) As Integer
    On Error GoTo Retorno
    If bRegistrado Then
        Dim Linha As Long
        
        For Linha = 1 To EditControl.ListItems.Count
        Set EditControl.SelectedItem = EditControl.ListItems(Linha)
            If Coluna <> 0 Then
                If EditControl.SelectedItem.SubItems(Coluna) = item Then
                    BuscaItemListView = Linha
                    Exit For
                End If
            Else
                If EditControl.SelectedItem = item Then
                    BuscaItemListView = Linha
                    Exit For
                End If
            End If
        Next
    End If
    Exit Function
Retorno:
    BuscaItemListView = 0
End Function

Private Sub Class_Initialize()
    ValidaComponente "CLASS"
End Sub
