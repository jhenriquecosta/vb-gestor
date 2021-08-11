VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodigoDeBarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public LinhaBarraGerada As String
Public Enum TipoModuloBarra
    etmModulo10 = 10
    etmModulo11 = 11
    etmModulo11ACN = 12
End Enum

Public Function CalculaDV(ByVal Linha As String, ByVal Modulo As TipoModuloBarra) As Integer
        Dim I As Integer
        Dim multiplicador  As Integer
        Dim somatorio As Integer
        Dim resultado  As Integer
        Dim Divisor As Integer
        Dim Aux As Integer
        
        multiplicador = 1
        somatorio = 0
        resultado = 0
        Aux = 0
        Divisor = IIf(Modulo = etmModulo11ACN, Modulo - 1, Modulo)
        I = Len(Linha)
        If Modulo = etmModulo10 Then
            Do
                multiplicador = IIf(multiplicador = 2, 1, 2)
                Aux = multiplicador * CInt(Mid(Linha, I, 1))
                resultado = IIf(Aux > 9, ((Aux - 10) + 1), Aux)
                somatorio = somatorio + resultado
                I = I - 1
            Loop While I > 0
        ElseIf Modulo = etmModulo11ACN Or Modulo = etmModulo11 Then
            'etmModulo11ACN - PARA AGENCIA, CEDENTE, NOSSO-NUMERO
            'etmModulo11 - PARA CODIGO DE BARRA
            multiplicador = IIf(Modulo = etmModulo11ACN, 9, 2)
            Do
                resultado = multiplicador * CInt(Mid(Linha, I, 1))
                somatorio = somatorio + resultado
                If Modulo = etmModulo11ACN Then
                    multiplicador = IIf(multiplicador = 2, 9, multiplicador - 1)
                Else
                    multiplicador = IIf(multiplicador = 9, 2, multiplicador + 1)
                End If
                I = I - 1
            Loop While I > 0
        End If
        CalculaDV = Abs(Divisor - (somatorio Mod Divisor))
        If CalculaDV = 10 Then
            If Modulo = etmModulo10 Then
                CalculaDV = "0"
            ElseIf Modulo = etmModulo11 Then
                CalculaDV = "1"
            ElseIf Modulo = etmModulo11ACN Then
                CalculaDV = "X"
            End If
        ElseIf CalculaDV = 11 Then
            CalculaDV = "1"
        End If
End Function

Private Function ConvenioTributo(CodImposto As String) As Double
    Dim Sql As String
    Dim Rs As VSRecordset
    If CodImposto = Const_Extrato Or CodImposto = Const_Notificacao Then CodImposto = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU))
    Sql = "select TIP_TCB_CONVENIO from tab_imposto where tip_cod_imposto ='" & CodImposto & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        ConvenioTributo = Nvl("" & Rs!TIP_TCB_CONVENIO, 0)
    End If
End Function

Public Function CriaLinhaDigitavel(ByVal Contribuinte As String, ByVal CodigoImposto As String, ByVal ValorTotal As Double, ByVal Periodo As String, ByVal DtVenc As String, Optional ByVal Parcela As String, Optional CodigoDocumento As String) As String
    On Error Resume Next
    Dim Codebar As New CodBarBB

    Dim Pos As Integer
    Dim Linha1 As String
    Dim Linha2 As String
    Dim Linha3 As String
    Dim Linha4  As String
    Dim Linha As String
    Dim StrID  As String
    Dim StrValor  As String
    Dim StrOrgao As String
    Dim StrNossoNumero As String
    Dim StrDigitavel As String
    Dim StrDV  As String
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Dim Correlativo As String
    Dim VersaoCodBarra As Integer
    Dim TipoContribuinte As Byte
    LinhaBarraGerada = ""
    'Identificacao do Produto = 8(Arrecadacao) / Segmento = 1(Prefeitura) / Identificador=6(Real)
    
    StrID = "81"
    
    'VERSAO DO COD. BARRA
    VersaoCodBarra = Format(Temp.PegaParametro(Bdados, "VERSAO BARRA"), "0")
    'CONTRIBUINTE
    Contribuinte = TiraTudo(Contribuinte)
    TipoContribuinte = IIf(Left(Contribuinte, 2) = "11" And Len(Contribuinte) = 10, 1, 2) '1=im ; 2=ic
    If TipoContribuinte = 1 Then
        Contribuinte = Contribuinte & "000"
    Else
        Contribuinte = Format(Contribuinte, "000000000000000")
    End If
    'PERIODO: AAMM
    Periodo = Format(Periodo, "000000")
    Periodo = IIf(TipoContribuinte = 1, Right(Periodo, 4) & Left(Periodo, 2), Right(Periodo, 4))
    'CORRELATIVO DO IMPOSTO
    Sql = "Select TIP_COD_CORRELATIVO from tab_imposto where tip_cod_imposto ='" & CodigoImposto & "'"
    If Bdados.AbreTabela(Sql, Rs) Then Correlativo = Format(Rs!TIP_COD_CORRELATIVO, "00")
    'PARCELA
    Parcela = Format(Parcela, "00")
    'FEBRABAN DA PREFEITURA
    
    StrOrgao = Format(Nvl(ConvenioTributo(CodigoImposto), 0), "0000")
    If CInt(Trim(StrOrgao)) = 0 Then
        StrOrgao = Edita.TiraTudo(Temp.PegaParametro(Bdados, "FEBRABAN"))
    End If
           
    'Campo Livre 25 POSICOES
    '********** 14-08-2004 - A partir desta data a forma de identificacao do pagamento serah o Numero do DAM*********
    If Trim(CodigoDocumento) <> "" Then
        If Len(StrOrgao) <= 4 Then
            If Nvl(Temp.PegaParametro(Bdados, "POSICAO DATA BARRA"), 0) > 0 Then
                StrNossoNumero = Right(DtVenc, 4) & Mid(DtVenc, 4, 2) & Left(DtVenc, 2) & VersaoCodBarra & Format(CodigoDocumento, "0000000000000000")
            Else
                StrNossoNumero = VersaoCodBarra & Format(CodigoDocumento, "000000000000000000000000")
            End If
            StrID = "81"
        Else
            If Nvl(Temp.PegaParametro(Bdados, "POSICAO DATA BARRA"), 0) > 0 Then
                DtVenc = DtVenc
                StrNossoNumero = Right(DtVenc, 4) & Mid(DtVenc, 4, 2) & Left(DtVenc, 2) & VersaoCodBarra & Format(CodigoDocumento, "000000000000")
            Else
                StrNossoNumero = VersaoCodBarra & Format(CodigoDocumento, "00000000000000000000")
            End If
            StrID = "86"
        End If
    Else
        StrNossoNumero = MontaNumeroPrefeituraFebraBan(VersaoCodBarra & Correlativo & Parcela & _
                        TipoContribuinte & Contribuinte & Periodo)
    End If
    If Temp.PegaParametro(Bdados, "BARRA COM VALOR") = "NAO" Then
        ValorTotal = 0
        StrID = StrID & "7"
    Else
        StrID = StrID & "6"
    End If
    StrValor = Format((CDbl(ValorTotal) * 100), "00000000000")
    
    StrDigitavel = StrID & StrValor & StrOrgao & StrNossoNumero
    'CALCULA DV GERAL
    '
    StrDV = Codebar.Mod10_Padrao(StrDigitavel)
    
    'StrDV = Codebar.Mod11_Padrao(StrDigitavel)
    
    StrDigitavel = StrID & StrDV & StrValor & StrOrgao & StrNossoNumero 'Linha do CB
    LinhaBarraGerada = StrDigitavel
    'PARTE 1
    Pos = 1: Linha1 = Mid(StrDigitavel, Pos, 11) & " " & Codebar.Mod10_Padrao(Mid(StrDigitavel, 1, 11))
    'PARTE 2
    Pos = Pos + 11: Linha2 = Mid(StrDigitavel, Pos, 11) & " " & Codebar.Mod10_Padrao(Mid(StrDigitavel, Pos, 11))
    'PARTE 3
    Pos = Pos + 11: Linha3 = Mid(StrDigitavel, Pos, 11) & " " & Codebar.Mod10_Padrao(Mid(StrDigitavel, Pos, 11))
    'PARTE 4
    Pos = Pos + 11: Linha4 = Mid(StrDigitavel, Pos, 11) & " " & Codebar.Mod10_Padrao(Mid(StrDigitavel, Pos, 11))
    
    Linha = Linha1 & " " & Linha2 & " " & Linha3 & " " & Linha4
    
    CriaLinhaDigitavel = Linha
    'DtVenc = Format(DtVenc, "yyyy") & Format(DtVenc, "mm") & Format(DtVenc, "dd")
End Function

Function MontaNumeroPrefeituraCgc(CodImposto As String, CodPagamento As String) As String
    Dim Num As String
    
    'Receita = 11 posic. / Cod. Pag. = 10
    MontaNumeroPrefeituraCgc = Trim(Num)

End Function

Function MontaNumeroPrefeituraFebraBan(Numero As String) As String
    Dim Num As String
    
    MontaNumeroPrefeituraFebraBan = Trim(Numero)

End Function

Public Function CriaLinhaDigitavelCBR(ByVal Contribuinte As String, ByVal CodigoImposto As String, ByVal ValorTotal As Double, ByVal Periodo As String, ObjetoPicture As Object, ByVal DtVenc As String, Optional ByVal Parcela As String, Optional CodigoDocumento As String) As String
    On Error Resume Next
    Dim Codebar As New CodBarBB

    Dim Pos As Integer
    Dim Linha1 As String
    Dim Linha2 As String
    Dim Linha3 As String
    Dim Linha4  As String
    Dim Linha As String
    
    Dim IDBanco  As String
    Dim CodigoMoeda  As String
    Dim DVGeral As String
    Dim nossoNumero As String
    Dim LinhaDigitavel As String
    Dim LinhaBarra As String
    Dim ValorDoc  As String
    Dim CampoLivre  As String
    Dim FatorVencimento  As String
    Dim Convenio As String
    Dim DataBase As String
    Dim Aux As String
    
    Dim Campo01 As String
    Dim Campo02 As String
    Dim Campo03 As String
    Dim Campo04 As String
    Dim Campo05 As String
    
    Dim carteira As String
    Dim Agencia As String
    Dim AgenciaComDigito As String
    Dim Conta As String
    
    Dim bcpLinhaDigitavel As String
    Dim bcpVariavelDV As String
    Dim bcpP1A4 As String, bcpP6A9 As String, bcpP10A19 As String, bcpP20A44 As String
'*************************************************************\*******
'    1. LINHA DO CODIGO DE BARRA                                    *
'                                                                   *
'    01 - 03 3   Identificação do banco                             *
'    04 - 04 1   Código da moeda (9 - real)                         *
'    05 - 05 1   Dígito Verificador Geral do Código de Barras       *
'    06 - 09 4   Fator de Vencimento                                *
'    10 - 19 10  Valor do Documento                                 *
'    20 - 44 25  Campo Livre                                        *
'********************************************************************
    LinhaBarraGerada = ""
    LinhaDigitavel = ""
    bcpVariavelDV = ""
    DataBase = "07/10/1997"
    'VERIFICAR CONVENIO SAUDE(PETROLINA)
    Convenio = Format(ConvenioTributo(CodigoImposto), "000000")
    If CDbl(Convenio) = 0 Then Convenio = Temp.PegaParametro(Bdados, "FEBRABAN")
    IDBanco = Format(Temp.PegaParametro(Bdados, "BANCO ARRECADACAO"), "000")
    CodigoMoeda = "9"
    FatorVencimento = Format(DateDiff("d", DataBase, DtVenc), "0000")
    ValorDoc = Format((CDbl(ValorTotal) * 100), "0000000000")
    If IDBanco = 1 Or IDBanco = 237 Or IDBanco = 3 Then 'BB/bradesco/Basa
    ' CODIGO BCP
        Campo01 = ""
        Campo02 = ""
        Campo03 = ""
        Campo04 = ""
        Campo05 = ""
        DVGeral = ""
        bcpP1A4 = ""
        bcpP10A19 = ""
        bcpP6A9 = ""
        bcpP20A44 = ""
        CodigoDocumento = Format(CodigoDocumento, "00000000000")
        nossoNumero = Format(CodigoDocumento, "00000000000")
        Conta = Format(Temp.PegaParametro(Bdados, "CONVENIO CONTA"), "0000000")
        Agencia = Format(Temp.PegaParametro(Bdados, "CONVENIO AGENCIA"), "000")
        AgenciaComDigito = Temp.PegaParametro(Bdados, "AGENCIA")
        AgenciaComDigito = Replace(AgenciaComDigito, "-", "")
        AgenciaComDigito = Format(AgenciaComDigito, "0000")
        carteira = Format(Temp.PegaParametro(Bdados, "CONVENIO CARTEIRA"), "00")
        
        bcpP1A4 = IDBanco & CodigoMoeda
        bcpP6A9 = FatorVencimento
        bcpP10A19 = ValorDoc
        If IDBanco = 3 Then 'BASA
            CodigoDocumento = Format(CodigoDocumento, "0000000")
            If Bdados.AbreTabela("SELECT TOC_NOSSO_NUMERO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_COD_OBRIGACAO=" & CodigoDocumento) Then
                nossoNumero = Format(Bdados.Tabela("TOC_NOSSO_NUMERO"), "0000000")
            End If
            
            'campos livres do banco
            bcpP20A44 = Format(AgenciaComDigito, "0000") & nossoNumero & Format(DtVenc, "DDMMYY") & Format(0, "0000000") & "0"
        
            DVGeral = ""
            'ESSE É O DV DO CÓDIGO DE BARRAS POSICAO 5
            DVGeral = calculoDigitoVerificadorCodigoBarra(bcpP1A4 & bcpP6A9 & bcpP10A19 & bcpP20A44)
            
            '**** ESSE É O CÓDIGO DE BARRAS *****
            LinhaBarraGerada = bcpP1A4 & DVGeral & bcpP6A9 & bcpP10A19 & bcpP20A44
            '*********************************************************************
            
            'AGORA COMEÇA A LINHA DIGITAVEL
            Campo01 = IDBanco & CodigoMoeda
            'Campo01 = Campo01 & "0"
            Campo01 = Campo01 & AgenciaComDigito
            Campo01 = Campo01 & Mid(nossoNumero, 1, 1)
            bcpVariavelDV = Campo01
            Campo01 = Campo01 & Modulo10(bcpVariavelDV)
            
            Campo02 = Mid(nossoNumero, 2, 6)
            Campo02 = Campo02 & Format(DtVenc, "DDMM")
            bcpVariavelDV = Campo02
            Campo02 = Campo02 & Modulo10(bcpVariavelDV)
            
            Campo03 = Format(DtVenc, "YY") & Mid(LinhaBarraGerada, 37, 7) & "0"
            bcpVariavelDV = Campo03
            Campo03 = Campo03 & Modulo10(bcpVariavelDV)
            
            Campo04 = FatorVencimento & ValorDoc
            
        ElseIf IDBanco = 237 Then
            bcpP20A44 = Format(Agencia, "0000") & Format(carteira, "00") & Format(CodigoDocumento, "00000000000") & Format(Conta, "0000000") & "0"
        
            DVGeral = ""
            DVGeral = calculoDigitoVerificadorCodigoBarra(bcpP1A4 & bcpP6A9 & bcpP10A19 & bcpP20A44)
            LinhaBarraGerada = bcpP1A4 & DVGeral & bcpP6A9 & bcpP10A19 & bcpP20A44
        
            Campo01 = IDBanco & CodigoMoeda
            Campo01 = Campo01 & "0"
            Campo01 = Campo01 & "." & Agencia
            bcpVariavelDV = Campo01
            Campo01 = Campo01 & calculoDigitoAutoConferencia(bcpVariavelDV)
            
            Campo02 = Format(carteira, "0") & Left(CodigoDocumento, 4) & "." & Mid(CodigoDocumento, 5, 5)
            bcpVariavelDV = Campo02
            Campo02 = Campo02 & calculoDigitoAutoConferencia(bcpVariavelDV)
            
            Campo03 = Right(CodigoDocumento, 2) & Left(Conta, 3) & "." & Right(Conta, 4) & "0"
            bcpVariavelDV = Campo03
            Campo03 = Campo03 & calculoDigitoAutoConferencia(bcpVariavelDV)
            
            Campo04 = FatorVencimento & ValorDoc
        
'        bcpLinhaDigitavel = Campo01 & " " & Campo02 & " " & Campo03 & " " & DVGeral & " " & Campo04
'        CriaLinhaDigitavelCBR = bcpLinhaDigitavel
  
  
        ElseIf IDBanco = 1 Then
            
            bcpP20A44 = Format(0, "000000") & Format(Convenio, "0000000") & Format(CodigoDocumento, "0000000000") & Format(carteira, "00")
            DVGeral = ""
            DVGeral = calculoDigitoVerificadorCodigoBarra(bcpP1A4 & bcpP6A9 & bcpP10A19 & bcpP20A44)
            LinhaBarraGerada = bcpP1A4 & DVGeral & bcpP6A9 & bcpP10A19 & bcpP20A44
          
            Campo01 = IDBanco & CodigoMoeda
            Campo01 = Campo01 & "0.0000"
            'Campo01 = Campo01 & "." & Agencia & "0"
            bcpVariavelDV = Campo01
            Campo01 = Campo01 & calculoDigitoAutoConferencia(bcpVariavelDV)
            
            Campo02 = Mid(LinhaBarraGerada, 25, 10)
            Campo02 = Left(Campo02, 5) & "." & Right(Campo02, 5)
            bcpVariavelDV = Campo02
            Campo02 = Campo02 & calculoDigitoAutoConferencia(bcpVariavelDV)
            
            Campo03 = Mid(LinhaBarraGerada, 35, 10)
            Campo03 = Left(Campo03, 5) & "." & Right(Campo03, 5)
            bcpVariavelDV = Campo03
            Campo03 = Campo03 & calculoDigitoAutoConferencia(bcpVariavelDV)
            
            Campo04 = FatorVencimento & ValorDoc
        End If
        
        bcpLinhaDigitavel = Campo01 & " " & Campo02 & " " & Campo03 & " " & DVGeral & " " & Campo04 '5 POSICAO 5
        CriaLinhaDigitavelCBR = bcpLinhaDigitavel
    ' FIM CODIGO BCP
    
    Else
    '
    'CODIGO ANTERIOR
    
    If Not Temp.PegaParametro(Bdados, "TIPO TITULO") = "FLOAT" Then 'SERA ANTI-FLOAT
        If Format(Temp.PegaParametro(Bdados, "BANCO ARRECADACAO"), "000") = "001" Then
            nossoNumero = Format(CodigoDocumento, "00000000000000000")
            'Campo Livre 25 POSICOES
            'TENHO QUE COMPLETAR AS 25 POSICOES (GB)
            If Len(Convenio & nossoNumero) = 24 Then
                CampoLivre = Convenio & nossoNumero & "0"
            ElseIf Len(Convenio & nossoNumero) = 23 Then
                CampoLivre = Convenio & nossoNumero & "21" 'AQUI
            ElseIf Len(Convenio & nossoNumero) = 22 Then
                CampoLivre = Convenio & nossoNumero & "012"
            ElseIf Len(Convenio & nossoNumero) = 21 Then
                CampoLivre = Convenio & nossoNumero & "0123"
            End If
        Else 'BRADESCO
'            20 - 23 4 Agência Cedente (Sem o digito verificador, completar com zeros a esquerda quando
'            necessário)
'            24 - 25 2 Carteira
'            26 - 36 11 Número do Nosso Número(Sem o digito verificador)
'            37 - 43 7 Conta do Cedente (Sem o digito verificador, completar com zeros a esquerda quando necessário)
'            44 - 44 1 Zero
            nossoNumero = Format(CodigoDocumento, "00000000000")
            Conta = Format(Temp.PegaParametro(Bdados, "CONVENIO CONTA"), "0000000")
            carteira = Format(Temp.PegaParametro(Bdados, "CONVENIO CARTEIRA"), "00")
            Agencia = Format(Temp.PegaParametro(Bdados, "CONVENIO AGENCIA"), "0000")
            CampoLivre = Agencia & carteira & nossoNumero & Conta & "0"
        End If
    Else
        Aux = Mid(String(25, "0"), 1, 17 - Len(Convenio))
        nossoNumero = Convenio & Format(CodigoDocumento, Aux)
        CampoLivre = Format(nossoNumero & "18", String(25, "0"))
    End If
'   1. LINHA DO CODIGO DE BARRA
    'CALCULA DV GERAL
    DVGeral = CalculaDV(IDBanco & CodigoMoeda & FatorVencimento & ValorDoc & CampoLivre, etmModulo11)
    LinhaBarra = IDBanco & CodigoMoeda & DVGeral & FatorVencimento & ValorDoc & CampoLivre
    
'********************************************************************
'    2. LINHA DIGITAVEL                                             *
'                                                                   *
'Campo  1    Pos 01 a 04  e pos 20 a 24 09                          *
'Dígito verificador Módulo 10           1                           *
'                                                                   *
'Campo  2    Pos 25 a 34                10                          *
'Dígito verificador Módulo 10           1                           *
'                                                                   *
'Campo  3    Pos 35 a 44                10                          *
'Dígito vefificador Módulo 10           1                           *
'                                                                   *
'Campo  4    Pos 05 (DV Geral)   01                                 *
'                                                                   *
'Campo  5    Pos 06 a 09 Fator de venc  04                          *
'                                                                   *
'Pos 10 a 19 (Valor  do  Título)        10                          *
'********************************************************************
    
'   2. LINHA DIGITAVEL
    'CAMPO 01
    Campo01 = IDBanco & CodigoMoeda & Left(CampoLivre, 5)
    Campo01 = Campo01 & CalculaDV(Campo01, etmModulo10)
    Campo01 = Left(Campo01, 5) & "." & Right(Campo01, 5)
    'CAMPO 02
    Campo02 = Mid(CampoLivre, 6, 10)
    Campo02 = Campo02 & CalculaDV(Campo02, etmModulo10)
    Campo02 = Left(Campo02, 5) & "." & Right(Campo02, 6)
    'CAMPO 03
    Campo03 = Mid(CampoLivre, 16, 10)
    Campo03 = Campo03 & CalculaDV(Campo03, etmModulo10)
    Campo03 = Left(Campo03, 5) & "." & Right(Campo03, 6)
    'CAMPO 04
    Campo04 = DVGeral
    'CAMPO 05
    Campo05 = FatorVencimento & ValorDoc
    
    LinhaDigitavel = Campo01 & " " & Campo02 & " " & Campo03 & " " & Campo04 & " " & Campo05
    
    '
    'BCP - ESTE É O RETORNO PADRAO - ANDRE
    CriaLinhaDigitavelCBR = LinhaDigitavel
    
'    Codebar.Imp_Codbar 7, 0, 0&, 1, 50, 0, "Times New Roman", 12, 0, 0, 30, 0, 0, LinhaDigitavel, ObjetoPicture
    'Codebar.Imp_Codbar 7, 0, 0&, 1, 50, 0, "Times New Roman", 12, 0, 0, 30, 0, 0, Left(LinhaBarraGerada, 4) & Mid(LinhaBarraGerada, 6), ObjetoPicture
    Codebar.Imp_Codbar 7, 0, 0&, 1, 50, 0, "Times New Roman", 12, 0, 0, 30, 0, 0, LinhaBarraGerada, ObjetoPicture
'    Codebar.Imprime_Barra LinhaBarraGerada, ObjetoPicture
End If
End Function

Function I25Encode(StringNumber)
   Dim asPattern(), sSTART, sSTOP

   ReDim asPattern(10)

   ' Padrões de Inicio e fim
   sSTART = "NNNN"
   sSTOP = "WNN"

   ' Padrões do código de barras
   asPattern(0) = "NNWWN"
   asPattern(1) = "WNNNW"
   asPattern(2) = "NWNNW"
   asPattern(3) = "WWNNN"
   asPattern(4) = "NNWNW"
   asPattern(5) = "WNWNN"
   asPattern(6) = "NWWNN"
   asPattern(7) = "NNNWW"
   asPattern(8) = "WNNWN"
   asPattern(9) = "NWNWN"

   If (Len(StringNumber) Mod 2) <> 0 Then
      ' O número de caracteres no
      ' argumento devem ser diferentes
      I25Encode = ""
      Exit Function
   End If

      
   If Not IsNumeric(StringNumber) Then
      ' Argumento deve ser numérico
      I25Encode = ""
      Exit Function
   Else
      If (InStr(StringNumber, ".") <> 0) Or _
         (InStr(StringNumber, ",") <> 0) Then
         ' O argumento é numero mais contem
         ' caracteres invalidos para nós
         I25Encode = ""
         Exit Function
      End If
   End If

   Dim sEncodedSTR, sUnit
   Dim iCharRead, sDigit1, sDigit2, I

   sEncodedSTR = ""
   For iCharRead = 1 To Len(StringNumber) Step 2
       sDigit1 = asPattern(Asc( _
                 Mid(StringNumber, iCharRead, 1)) - 48)
       sDigit2 = asPattern(Asc(Mid( _
                 StringNumber, iCharRead + 1, 1)) - 48)

       sUnit = ""

       For I = 1 To 5
           sUnit = sUnit & Mid(sDigit1, I, 1) & _
                   Mid(sDigit2, I, 1)
       Next

       sEncodedSTR = sEncodedSTR & sUnit
   Next
   I25Encode = sSTART & sEncodedSTR & sSTOP
End Function
'BCP - PARA LINHA DIGITAVEL DE ACORDO COM A DOCUMENTACAO BRADESCO
Private Function calculoDigitoVerificadorCodigoBarra(Linha As String)
    'algoritmoModulo11
    'multiplicar cada caractere da direita para esquerda por 9 a 2
    Dim f As Integer 'f=fator padrao 9
    Dim c As Integer 'cada=caractere do campo
    Dim x As Integer 'varivel de iteracao
    Dim r As Integer 'r= resultado
    Dim d As String 'd=digito
    Linha = retiraSeparadores(Linha)
    f = 2
    r = 0
    d = "0"
    For x = 0 To Len(Linha) - 1
         c = Mid(Linha, Len(Linha) - x, 1) ' Len campo = tamanho do campo
         c = c * f
         r = r + c
         If f = 9 Then
            f = 2
         Else
            f = f + 1
         End If
    Next x
    d = r Mod 11 ' de acordo com o documento
    d = 11 - d
    If d = 0 Or d = 1 Or d > 9 Then
        d = 1
    End If
    calculoDigitoVerificadorCodigoBarra = d
End Function
Private Function calculoBasaDigitoAutoConferencia(Campo As String)
    'ESTA FUNCAO VAI TER O MESMO CONCEITO DE MODULO 10
    'PARA NAO MEXER NA QUE JÁ EXISTE
    'algoritmoModulo10
    'multiplicar cada caractere da direita para esquerda por 2 e 1
    Dim f As Integer 'f=fator padrao 2
    Dim c As Integer 'cada=caractere do campo
    Dim x As Integer 'varivel de iteracao
    Dim r As Integer 'r= resultado
    Dim resto As Integer 'resto
    Campo = retiraSeparadores(Campo)
    f = 2
    m = 0
    r = 0
    
    For x = 1 To Len(Campo)
         c = Mid(Campo, x, 1) ' Len campo = tamanho do campo
         c = c * f
         r = r + somarResultadoDV(c)
         If f = 2 Then
            f = 1
         ElseIf f = 1 Then
            f = 2
         End If
    Next x
    If r < 10 Then
        resto = r
    Else
        resto = r Mod 10
    End If
    If resto > 0 Then
        calculoBasaDigitoAutoConferencia = 10 - resto
    Else
        calculoBasaDigitoAutoConferencia = resto
    End If
End Function
Function Modulo10(strNumero As String) As String
    Dim intContador, intMultiplicador, intSelecionado, intResultado, intDigito, intResto As Integer
    Dim strProduto As String
    Dim dv As String
    intResultado = 0
    'inicia o multiplicador
    intMultiplicador = 2
    For intContador = Len(strNumero) To 1 Step -1
        'pega o item do indice
        intSelecionado = Val(Mid(strNumero, intContador, 1))
        'extrai o caracter e multiplica pelo multiplicador, gerando o produto
        strProduto = intSelecionado * intMultiplicador
        For intDigito = 1 To Len(strProduto)
            intResultado = intResultado + Mid(strProduto, intDigito, 1)
        Next
        'se o multiplicador for igual a 2 atribuir valor 1 se for 1 atribui 2
        intMultiplicador = IIf(intMultiplicador = 2, 1, 2)
    Next
    If intResultado >= 10 Then
        intResto = intResultado Mod 10
    Else
        intResto = intResultado
    End If
    If intResto > 0 Then
        intResto = 10 - intResto
    End If
    Modulo10 = intResto
End Function
Function Calculo_DV10(strNumero As String) As String
    'declara As variáveis
    Dim intContador As Integer
    Dim intNumero As Integer
    
    Dim intTotalNumero As Integer
    
    Dim intMultiplicador As Integer
    
    Dim intResto As Integer
    
    ' se nao for um valor numerico sai da função
    If Not IsNumeric(strNumero) Then
      Calculo_DV10 = ""
      Exit Function
    End If
    
    'inicia o multiplicador
    intMultiplicador = 2
    
    'pega cada caracter do numero a partir da direita
    For intContador = Len(strNumero) To 1 Step -1
    
    'extrai o caracter e multiplica pelo multiplicador
    intNumero = Val(Mid(strNumero, intContador, 1)) * intMultiplicador
    
    ' se o resultado for maior que nove soma os algarismos do resultado
    If intNumero > 9 Then
      intNumero = Val(Left(intNumero, 1)) + Val(Right(intNumero, 1))
    End If
    
    'soma o resultado para totalização
    intTotalNumero = intTotalNumero + intNumero
    
    'se o multiplicador for igual a 2 atribuir valor 1 se for 1 atribui 2
    intMultiplicador = IIf(intMultiplicador = 2, 1, 2)
    
    Next
    
        Dim DezenaSuperior As Integer
        If intTotalNumero < 10 Then
            DezenaSuperior = 10
        Else
            DezenaSuperior = 10 * (Val(Left(CStr(intTotalNumero), 1)) + 1)
        End If
        intResto = DezenaSuperior - intTotalNumero
    
    'verifica as exceções ( 0 -> DV=0 )
    Select Case intResto
      Case 0
         Calculo_DV10 = "0"
      Case 1
         Calculo_DV10 = "1"
      Case 10
         Calculo_DV10 = "1"
      Case 11
         Calculo_DV10 = "1"
      Case Else
         Calculo_DV10 = Str(intResto)
    End Select

End Function





Private Function calculoDigitoAutoConferencia(Campo As String)
    'algoritmoModulo10
    'multiplicar cada caractere da direita para esquerda por 2 e 1
    Dim f As Integer 'f=fator padrao 2
    Dim c As Integer 'cada=caractere do campo
    Dim x As Integer 'varivel de iteracao
    Dim r As Integer 'r= resultado
    Dim m As Integer 'multiplicador se o resultado for 11 multiplicar por  2, 21 por 3 ....
    Campo = retiraSeparadores(Campo)
    f = 2
    m = 0
    r = 0
    
    For x = 0 To Len(Campo) - 1
         c = Mid(Campo, Len(Campo) - x, 1) ' Len campo = tamanho do campo
         c = c * f
         r = r + somarResultadoDV(c)
         If f = 2 Then
            f = 1
         ElseIf f = 1 Then
            f = 2
         End If
    Next x
    If Right(r, 1) > 0 Then
        m = 10 * (Left(r, 1) + 1)
    Else
        m = 10 * (Left(r, 1))
    End If
    calculoDigitoAutoConferencia = m - r
End Function
Private Function somarResultadoDV(v As Integer) As Integer 'v= valor
    Dim r As Integer 'resultado
    r = v
    If Len(Trim(v)) = 2 Then
        r = 0
        r = Left(v, 1)
        r = r + Right(v, 1)
    End If
    somarResultadoDV = r
End Function
Private Function retiraSeparadores(Valor As String) As String
    Valor = Replace(Valor, ",", "")
    Valor = Replace(Valor, ".", "")
    Valor = Replace(Valor, "-", "")
    Valor = Replace(Valor, "/", "")
    Valor = Replace(Valor, " ", "")
    retiraSeparadores = Valor
End Function

'BCP - ESTES DOIS METODOS TEM QUE ESTAR IGUAL AO REMESSA BRADESCO EM ARR: nome gerarDigitoConferencia82
Public Function gerarDigitoConferencia(carteira As String, NumeroDocumento) As String '' numeroDocumento aqui, mas a classe ja tem tem ele (nao precisa)
    'carteira formato  (00)
    Dim x As Integer, resultado As Integer, digito As String
    Dim nossoNumero As String, Numero As String
    nossoNumero = Format(carteira, "00") & NumeroDocumento
    Numero = 0
    resultado = 0
    For x = 1 To Len(nossoNumero)
        Numero = Mid(nossoNumero, x, 1)
        Numero = Format(Numero * numeroParaMultiplicacao(x), "00")
        resultado = resultado + Numero
    Next x
    digito = resultado Mod 11 ' de acordo com o documento
    If digito = 1 Then
        digito = "P"
    ElseIf digito = 0 Then
        digito = 0
    Else
        digito = 11 - digito
    End If
    
    gerarDigitoConferencia = digito
End Function
Private Function numeroParaMultiplicacao(Posicao As Integer) As Integer
        Dim Numero As Integer
        If Posicao = 1 Or Posicao = 7 Or Posicao = 13 Then
            Numero = 2
        ElseIf Posicao = 2 Or Posicao = 8 Then
            Numero = 7
        ElseIf Posicao = 3 Or Posicao = 9 Then
            Numero = 6
        ElseIf Posicao = 4 Or Posicao = 10 Then
            Numero = 5
        ElseIf Posicao = 5 Or Posicao = 11 Then
            Numero = 4
        ElseIf Posicao = 6 Or Posicao = 12 Then
            Numero = 3
        End If
        numeroParaMultiplicacao = Numero
End Function

