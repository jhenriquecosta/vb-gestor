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
        If Len(StrOrgao) = 4 Then
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
    Dim NossoNumero As String
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
    
    Dim Carteira As String
    Dim Agencia As String
    Dim Conta As String
'********************************************************************
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
    DataBase = "07/10/1997"
    'VERIFICAR CONVENIO SAUDE(PETROLINA)
    Convenio = Format(ConvenioTributo(CodigoImposto), "000000")
    If CDbl(Convenio) = 0 Then Convenio = Temp.PegaParametro(Bdados, "FEBRABAN")
    IDBanco = Format(Temp.PegaParametro(Bdados, "BANCO ARRECADACAO"), "000")
    CodigoMoeda = "9"
    FatorVencimento = Format(DateDiff("d", DataBase, DtVenc), "0000")
    ValorDoc = Format((CDbl(ValorTotal) * 100), "0000000000")
    
    If Not Temp.PegaParametro(Bdados, "TIPO TITULO") = "FLOAT" Or Format(Temp.PegaParametro(Bdados, "BANCO ARRECADACAO"), "000") = "237" Then 'SERA ANTI-FLOAT
        If Format(Temp.PegaParametro(Bdados, "BANCO ARRECADACAO"), "000") = "001" Then
            NossoNumero = Format(CodigoDocumento, "00000000000000000")
            'Campo Livre 25 POSICOES
            'TENHO QUE COMPLETAR AS 25 POSICOES (GB)
            If Len(Convenio & NossoNumero) = 24 Then
                CampoLivre = Convenio & NossoNumero & "0"
            ElseIf Len(Convenio & NossoNumero) = 23 Then
                CampoLivre = Convenio & NossoNumero & "21" 'AQUI
            ElseIf Len(Convenio & NossoNumero) = 22 Then
                CampoLivre = Convenio & NossoNumero & "012"
            ElseIf Len(Convenio & NossoNumero) = 21 Then
                CampoLivre = Convenio & NossoNumero & "0123"
            End If
        Else 'BRADESCO
'            20 - 23 4 Agência Cedente (Sem o digito verificador, completar com zeros a esquerda quando
'            necessário)
'            24 - 25 2 Carteira
'            26 - 36 11 Número do Nosso Número(Sem o digito verificador)
'            37 - 43 7 Conta do Cedente (Sem o digito verificador, completar com zeros a esquerda quando necessário)
'            44 - 44 1 Zero
            NossoNumero = Format(CodigoDocumento, "00000000000")
            Conta = Format(Temp.PegaParametro(Bdados, "CONVENIO CONTA"), "0000000")
            Carteira = Format(Temp.PegaParametro(Bdados, "CONVENIO CARTEIRA"), "00")
            Agencia = Format(Temp.PegaParametro(Bdados, "CONVENIO AGENCIA"), "0000")
            CampoLivre = Agencia & Carteira & NossoNumero & Conta & "0"
        End If
    Else
        Aux = Mid(String(25, "0"), 1, 17 - Len(Convenio))
        NossoNumero = Convenio & Format(CodigoDocumento, Aux)
        CampoLivre = Format(NossoNumero & "18", String(25, "0"))
    End If
'   1. LINHA DO CODIGO DE BARRA
    'CALCULA DV GERAL
    DVGeral = CalculaDV(IDBanco & CodigoMoeda & FatorVencimento & ValorDoc & CampoLivre, etmModulo11)
    LinhaBarra = IDBanco & CodigoMoeda & DVGeral & FatorVencimento & ValorDoc & CampoLivre
    LinhaBarraGerada = LinhaBarra
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
    CriaLinhaDigitavelCBR = LinhaDigitavel
    
'    Codebar.Imp_Codbar 7, 0, 0&, 1, 50, 0, "Times New Roman", 12, 0, 0, 30, 0, 0, LinhaDigitavel, ObjetoPicture
    'Codebar.Imp_Codbar 7, 0, 0&, 1, 50, 0, "Times New Roman", 12, 0, 0, 30, 0, 0, Left(LinhaBarraGerada, 4) & Mid(LinhaBarraGerada, 6), ObjetoPicture
    Codebar.Imp_Codbar 7, 0, 0&, 1, 50, 0, "Times New Roman", 12, 0, 0, 30, 0, 0, LinhaBarraGerada, ObjetoPicture
'    Codebar.Imprime_Barra LinhaBarraGerada, ObjetoPicture
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

