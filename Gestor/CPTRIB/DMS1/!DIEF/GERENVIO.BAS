Attribute VB_Name = "Module3"
Option Explicit

Dim snap_detalhe As Recordset
Type gerador
    inscricao       As String
    periodo_inicial As String
    tipo_decla      As Integer
    data_geracao    As String
    condicao        As Byte
End Type

Global gera() As gerador
Global Versao As String

Global Gridgeral() As Variant
Global linha_erro As Integer
Global retorno_critica As Integer
Global motivo_entradas As String
Global val_fatura_normal As String
Global Val_fat_mensal_epp As String
Global Val_aliquota_epp As String
Global Val_imposto_epp As String


Global val_doEstado As String
Global val_deOutr_Est As String
Global val_doExterior As String
Global val_total As String

Global val_txt400 As String
Global val_txt416 As String
Global val_txt417 As String
Global val_txt418 As String
Global val_txt419 As String
Global val_txt420 As String
Global val_txt421 As String
Global val_txt422 As String
Global val_txt423 As String

Function calcula_controle(num_inscricao As String, periodo_inicial As String, tipo_decla As Integer) As Double
    
    
    query = "Select * From TAB_DETALHE_DEClARACAO "
    query = query & "Where tdd_num_inscricao  = '" & num_inscricao & "'"
    query = query & " and tdd_periodo_inicial = '" & periodo_inicial & "'"
    query = query & " and tdd_substitutiva = " & tipo_decla
    
    Set snap_detalhe = dbgim.OpenRecordset(query, dbOpenSnapshot)
    
    Dim acumulador As Double
    acumulador = 0
    While Not snap_detalhe.EOF
        
        acumulador = acumulador + CDbl(snap_detalhe!tdd_item * snap_detalhe!tdd_valor)
        snap_detalhe.MoveNext
        
    Wend
    
    calcula_controle = acumulador
    snap_detalhe.Close
    
End Function

Private Function PutZeros(ByVal valor As String)
    
   
    valor = Trim(Format$(valor, "0000000000000.00"))
    
    If Left$(Trim(valor), 1) = "-" Then
        PutZeros = "-" & String$(Len(valor) - Len(Trim(valor)), "0") + (Left$(Right$(Trim(valor), (Len(Trim(valor)) - 1)), (Len(Trim(valor)) - 4)) & Right$(Trim(valor), 2))
    Else
        PutZeros = String$(Len(valor) - Len(Trim(valor)), "0") + (Left$(Trim(valor), Len(Trim(valor)) - 3) & Right$(Trim(valor), 2))
    End If
    
End Function














Function trata_retorno() As Integer
'
' Função trata retorno de acesso a disco
'
    If Err <> 0 Then
        
        Select Case Err
        Case 52
'            trata_retorno = MsgBox(Error, 53, NOME_SISTEMA_DEFAULT)
            trata_retorno = MsgBox("Nome/Número do Arquivo inválido.", 53, NOMBRE_SISTEMA_DEFAULT)
        Case 53
'            trata_retorno = MsgBox(Error, 53, NOME_SISTEMA_DEFAULT)
            trata_retorno = MsgBox("Arquivo não existe.", 53, NOME_SISTEMA_DEFAULT)
        Case 57, 67
'            trata_retorno = MsgBox(Error, 53, NOME_SISTEMA_DEFAULT)
            trata_retorno = MsgBox("Erro de Dispositivo de I/O.", 53, NOME_SISTEMA_DEFAULT)
        Case 61
  '          trata_retorno = MsgBox(Error, 53, NOME_SISTEMA_DEFAULT)
            trata_retorno = MsgBox("Disco Cheio.", 53, NOME_SISTEMA_DEFAULT)
        Case 64
             trata_retorno = MsgBox(Error, 53, NOME_SISTEMA_DEFAULT)
        Case 68
'            trata_retorno = MsgBox(Error, 53, NOME_SISTEMA_DEFAULT)
            trata_retorno = MsgBox("Dispositivo não disponivel.", 53, NOME_SISTEMA_DEFAULT)
        Case 71
'            trata_retorno = MsgBox(Error, 53, NOME_SISTEMA_DEFAULT)
            trata_retorno = MsgBox("Dispositivo não está pronto./Disco ausente.", 53, NOME_SISTEMA_DEFAULT)
        Case 75
'            trata_retorno = MsgBox(Error, 53, NOME_SISTEMA_DEFAULT)
            trata_retorno = MsgBox("Erro de Acesso ou caminho não especificado.", 53, NOME_SISTEMA_DEFAULT)
        Case 70
            'If UCase(Mid$(CurDir, 1, 1)) = "A" Or UCase(Mid$(CurDir, 1, 1)) = "B" Then
'               trata_retorno = MsgBox(Error, 53, NOME_SISTEMA_DEFAULT)
               trata_retorno = MsgBox("Permissão negada/Disquete protegido contra gravação.", 53, NOME_SISTEMA_DEFAULT)
            'Else
            '   trata_retorno = MsgBox(S3(5), 53)
            'End If
        Case 1000
               trata_retorno = MsgBox(Error, 53, NOME_SISTEMA_DEFAULT)
        Case Else
               trata_retorno = MsgBox("Problemas na unidade selecionada." & Chr(13) & "Erro " & Error & " ocorreu.", 53, NOME_SISTEMA_DEFAULT)
        End Select
    End If

End Function












