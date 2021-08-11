Attribute VB_Name = "Module1"
Option Explicit
Global teve_alteracao_contador As Integer
Global teve_alteracao_declaracao As Boolean
Global gformaAnte As Integer
Global vim_trata_epp As Integer
Global vim_de_onde As String
Global G_opt_fat1 As Integer
Global G_opt_fat2 As Integer
Global G_opt_fat3 As Integer
Global Contadores As Recordset
Global g_nro_contador As String
Global msgerro As String
Global rota As Integer
Global primeira_vez As Integer
Global busca_total_contador As Integer
Global quem_me_chamou As String
Global retorno As String
Global restauracontr As Recordset
Global reta As Integer
Global dbgim As Database 'variavel para receber o banco de dados
Global PESQ_SNAP As Recordset
Global strSQLA As String
Global Const NOME_SISTEMA_DEFAULT = "DIEF - SEFAZ - MA"
Global Const NOMBRE_SISTEMA_DEFAULT = "DIEF - SEFAZ - MA"
Global out3 As Integer

'Globales del usuario de Oracle
Global Const PRIMERACORRIDA = 0
Global Const DEC_SEP_DEFAULT = ","
Global Const DBFILE_DEFAULT = "Dief.MDB"
'Global Const HOSTNAME_DEFAULT = "VB4DUMMY"
'Hasta aquí

Global Const BASE9 = 9

Global g_nro_inscricao As String
Global teve_alteracao As Integer

Global TabUf(27) As String * 2
Global g_ano_inicio_operacao As Integer
Global g_periodo_inicial As String
Global g_num_inscricao As String
Global g_num_cgc_cpf As String
Global g_tipo_decla As Integer
Global g_tipo_contribuinte As Byte
Global g_aliquota_Maior17 As Byte
Global gformulario As Integer

   Global BuscaContr As Integer
   Global BuscaContad As Integer
   
   Global Navega_Ficha(0 To 1, 0 To 4) As Integer
' ***** estava no criticas
Global resposta As Integer
Global tab_temp1 As Recordset
Global Query1 As String
Global linhas_impressas As Integer

Global a, c, X
Global b As Integer
Global tipo As String

Global status As Integer
Global condicao As String
Global TabTemp As Recordset

Global g_condicao As Byte
Global query As String
Global seleccion As Recordset
Global g_tot_val_contabil As Double
Global g_tot_val_contabil_saida As Double
Global g_tot_val_para_faturamento As Double 'valores Contabeis de Vendas - P/o Estado
                                                                          ' p/outros Estados
                                                                          ' p/o Exterior
   
Global g_tot_val_imp_debitado As Double
Global g_tot_val_imp_cred_entradas As Double
Global g_tot_base_calc As Double
Global g_tot_imp_creditado As Double
Global g_tot_isentas As Double
Global g_tot_outras As Double

Global ArrayEntra() As Variant
Global apurICMS() As Variant
Global ArrayDetDesp As Variant

Global ArrayFatura() As Variant
Global ArraySaidas() As Variant
Global CApurICMS() As Variant
Global ArrayRec() As Variant
Global VALOR200 As String
Global VALOR214 As String
Global VALOR215 As String
Global VALOR216 As String
Global VALOR300 As String
Global VALOR301 As String
Global VALOR302 As String
Global VALOR303 As String
Global VALOR400 As String
Global VALOR424 As String

Global VALOR416 As String
Global VALOR417 As String
Global VALOR418 As String
Global VALOR419 As String
Global VALOR420 As String
Global VALOR421 As String
Global VALOR422 As String
Global VALOR423 As String
Global MaxCol As Integer
Global MaxRow As Integer
Global MaxCol_s As Integer
Global MaxRow_s As Integer
Global MaxCol_ICMS As Integer
Global MaxRow_ICMS As Integer
Global MaxCol_rec As Integer
Global MaxRow_rec As Integer
Global MaxCol_fat As Integer
Global MaxRow_fat As Integer
Global MaxCol_DetDesp As Integer
Global MaxRow_DetDesp As Integer

Function verifica_preenchimento_dados_complementares(txt400 As String, txt416 As String, txt417 As String, txt418 As String, txt419 As String, txt420 As String, txt421 As String, txt422 As String, txt423 As String, txt424 As String, grid_DetDesp() As Variant) As Integer
verifica_preenchimento_dados_complementares = True
Dim coluna As Integer
Dim linha As Integer


If Trim(txt400) = "" And Trim(txt424) = "" And Trim(txt416) = "" And Trim(txt417) = "" And Trim(txt418) = "" And Trim(txt419) = "" Then
  If Trim(txt420) = "" And Trim(txt421) = "" And Trim(txt422) = "" And Trim(txt423) = "" Then
    For coluna = 1 To 1
          For linha = 0 To 13
             If Trim(grid_DetDesp(coluna, linha)) <> "" Then
                 Exit For
             End If
          Next linha
       Next coluna
    If linha = 14 Then
       verifica_preenchimento_dados_complementares = False
    End If
 End If
End If

End Function

Sub cargar_Array_ApurICMS(inscricao As String, periodo As String, Substitutiva As Integer)
 
 Dim i As Long
   
 query = "Select *  From tab_detalhe_declaracao " _
 & "Where tdd_num_inscricao  = '" & inscricao & "'" _
 & " and tdd_periodo_inicial = '" & periodo & "'" _
 & " and tdd_substitutiva = " & Substitutiva
   query = query & " and tdd_item between 200 and 221 "
  
  ReDim apurICMS(0 To 1, 0 To 17)
  ReDim Preserve apurICMS(0 To 1, 0 To 17)
  MaxCol_ICMS = 2
  MaxRow_ICMS = 18
  Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
  If seleccion.BOF Or seleccion.EOF Then
       seleccion.Close
       Call monta_cabecalho_ApurICMS
        Exit Sub
  End If
  i = 0
  For i = 0 To 17
     apurICMS(1, i) = "0,00"
  Next i
    
  seleccion.MoveFirst
    
   Call monta_cabecalho_ApurICMS
    i = 0
    While Not seleccion.EOF
      Select Case seleccion("tdd_item")
       Case Is = 200
        VALOR200 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
          'msk_Fat_frm2.Text = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 201   '  coluna 1
         apurICMS(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 202
         apurICMS(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 203
         apurICMS(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 204
         apurICMS(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 205
         apurICMS(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 206
         apurICMS(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 207
         apurICMS(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 208
         apurICMS(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 209
         apurICMS(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 210
         apurICMS(1, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 211
         apurICMS(1, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
        Case Is = 212
         apurICMS(1, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 213
         apurICMS(1, 12) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      Case Is = 214
         apurICMS(1, 13) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 215
        apurICMS(1, 14) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      Case Is = 216
         apurICMS(1, 15) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      Case Is = 217
          apurICMS(1, 16) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      Case Is = 218
          apurICMS(1, 17) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      Case Is = 219
         VALOR214 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      Case Is = 220
         VALOR215 = Format(seleccion("tdd_valor"), "###,###,##0.0")
      Case Is = 221
         VALOR216 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      End Select
       seleccion.MoveNext
       i = i + 1
    Wend
       
    MaxCol_ICMS = 2
    MaxRow_ICMS = 18
    seleccion.Close
End Sub

Sub cargar_Array_Entrada(inscricao As String, periodo As String, Substitutiva As Integer)
 
 Dim estado As String
 Dim i As Long
   
  query = "Select *  From tab_detalhe_declaracao " _
 & "Where tdd_num_inscricao  = '" & inscricao & "'" _
 & " and tdd_periodo_inicial = '" & periodo & "'" _
 & " and tdd_substitutiva = " & Substitutiva
  
 query = query & " and tdd_item between 01 and 60 "
  
 ReDim ArrayEntra(0 To 5, 0 To 12)
 ReDim Preserve ArrayEntra(0 To 5, 0 To 12)
 MaxCol = 6
 MaxRow = 13
  
 Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
 If seleccion.BOF Or seleccion.EOF Then
        seleccion.Close
        monta_cabecalho_Entrada
        Exit Sub
 End If
  
 seleccion.MoveFirst
 monta_cabecalho_Entrada
 i = 0
 While Not seleccion.EOF
      Select Case seleccion("tdd_item")
       Case Is = 1   '  coluna 1
         ArrayEntra(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 2
         ArrayEntra(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 3
         ArrayEntra(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 4
         ArrayEntra(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 5
         ArrayEntra(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 6
         ArrayEntra(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 7
         ArrayEntra(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 8
         ArrayEntra(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 9
         ArrayEntra(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 10
         ArrayEntra(1, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 11
         ArrayEntra(1, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 12
         ArrayEntra(1, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
        Case Is = 13
         ArrayEntra(2, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 14
         ArrayEntra(2, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 15
         ArrayEntra(2, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 16
         ArrayEntra(2, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 17
         ArrayEntra(2, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 18
         ArrayEntra(2, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 19
         ArrayEntra(2, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 20
         ArrayEntra(2, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 21
         ArrayEntra(2, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 22
         ArrayEntra(2, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 23
         ArrayEntra(2, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 24
         ArrayEntra(2, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
     'coluna 3
       Case Is = 25
         ArrayEntra(3, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 26
         ArrayEntra(3, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 27
         ArrayEntra(3, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 28
         ArrayEntra(3, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 29
         ArrayEntra(3, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 30
         ArrayEntra(3, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 31
         ArrayEntra(3, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 32
         ArrayEntra(3, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 33
         ArrayEntra(3, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 34
         ArrayEntra(3, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 35
         ArrayEntra(3, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 36
        ArrayEntra(3, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
     'coluna 4
       Case Is = 37
         ArrayEntra(4, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 38
         ArrayEntra(4, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 39
         ArrayEntra(4, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 40
         ArrayEntra(4, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 41
         ArrayEntra(4, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 42
         ArrayEntra(4, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 43
         ArrayEntra(4, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 44
         ArrayEntra(4, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 45
         ArrayEntra(4, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 46
         ArrayEntra(4, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 47
         ArrayEntra(4, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 48
         ArrayEntra(4, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
     'coluna 5
       Case Is = 49
         ArrayEntra(5, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 50
         ArrayEntra(5, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 51
         ArrayEntra(5, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 52
         ArrayEntra(5, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 53
         ArrayEntra(5, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 54
         ArrayEntra(5, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 55
         ArrayEntra(5, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 56
         ArrayEntra(5, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 57
         ArrayEntra(5, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 58
         ArrayEntra(5, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 59
         ArrayEntra(5, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 60
         ArrayEntra(5, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      End Select
      seleccion.MoveNext
      i = i + 1
    Wend
    'grid_dem_entradas.Refresh
    MaxCol = 6
    MaxRow = 13
    seleccion.Close
End Sub

Sub cargar_Array_saidas(inscricao As String, periodo As String, Substitutiva As Integer)

 Dim estado As String
 Dim i As Long
   
  query = "Select *  From tab_detalhe_declaracao " _
 & "Where tdd_num_inscricao  = '" & inscricao & "'" _
 & " and tdd_periodo_inicial = '" & periodo & "'" _
 & " and tdd_substitutiva = " & Substitutiva
  
 query = query & " and tdd_item between 100 and 159 "
  
 ReDim ArraySaidas(0 To 5, 0 To 12)
 ReDim Preserve ArraySaidas(0 To 5, 0 To 12)
 MaxCol_s = 6
 MaxRow_s = 12
 Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
 If seleccion.BOF Or seleccion.EOF Then
    seleccion.Close
    monta_Arraysaidas
    Exit Sub
 End If
 seleccion.MoveFirst
    
 monta_Arraysaidas
 i = 0
 While Not seleccion.EOF
      Select Case seleccion("tdd_item")
       Case Is = 100   '  coluna 1
         ArraySaidas(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 101
         ArraySaidas(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 102
         ArraySaidas(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 103
         ArraySaidas(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 104
         ArraySaidas(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 105
         ArraySaidas(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 106
         ArraySaidas(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 107
         ArraySaidas(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 108
         ArraySaidas(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 109
         ArraySaidas(1, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 110
         ArraySaidas(1, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 111
         ArraySaidas(1, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         'coluna 2
       Case Is = 112
         ArraySaidas(2, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 113
         ArraySaidas(2, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 114
         ArraySaidas(2, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 115
         ArraySaidas(2, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 116
         ArraySaidas(2, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 117
         ArraySaidas(2, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 118
         ArraySaidas(2, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 119
         ArraySaidas(2, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 120
         ArraySaidas(2, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 121
         ArraySaidas(2, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 122
         ArraySaidas(2, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 123
         ArraySaidas(2, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
     'coluna 3
       Case Is = 124
         ArraySaidas(3, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 125
         ArraySaidas(3, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 126
         ArraySaidas(3, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 127
         ArraySaidas(3, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 128
         ArraySaidas(3, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 129
         ArraySaidas(3, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 130
         ArraySaidas(3, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 131
         ArraySaidas(3, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 132
         ArraySaidas(3, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 133
         ArraySaidas(3, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 134
         ArraySaidas(3, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 135
         ArraySaidas(3, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
     'coluna 4
       Case Is = 136
         ArraySaidas(4, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 137
         ArraySaidas(4, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 138
         ArraySaidas(4, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 139
         ArraySaidas(4, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 140
         ArraySaidas(4, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 141
         ArraySaidas(4, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 142
         ArraySaidas(4, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 143
         ArraySaidas(4, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 144
         ArraySaidas(4, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 145
         ArraySaidas(4, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 146
         ArraySaidas(4, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 147
         ArraySaidas(4, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
     'coluna 5
       Case Is = 148
         ArraySaidas(5, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 149
         ArraySaidas(5, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 150
         ArraySaidas(5, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 151
         ArraySaidas(5, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 152
         ArraySaidas(5, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 153
         ArraySaidas(5, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 154
         ArraySaidas(5, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 155
         ArraySaidas(5, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 156
         ArraySaidas(5, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 157
         ArraySaidas(5, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 158
         ArraySaidas(5, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 159
        ArraySaidas(5, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
      End Select
        seleccion.MoveNext
        i = i + 1
    Wend
    seleccion.Close
End Sub
Function converte_tipo_logradouro(tipo As String) As String
converte_tipo_logradouro = ""

Select Case Trim(tipo)
   Case "AL"
  converte_tipo_logradouro = "Alameda"
   Case "AV"
  converte_tipo_logradouro = "Avenida"
   Case "CAM"
  converte_tipo_logradouro = "Caminho"
   Case "CON"
  converte_tipo_logradouro = "Condominio"
   Case "ETR"
  converte_tipo_logradouro = "Estrada"
   Case "ILH"
  converte_tipo_logradouro = "Ilha"
   Case "LAD"
  converte_tipo_logradouro = "Ladeira"
   Case "LOT"
  converte_tipo_logradouro = "Loteamento"
   Case "LUG"
  converte_tipo_logradouro = "Lugarejo"
   Case "MRO"
  converte_tipo_logradouro = "Morro"
   Case "PCA"
  converte_tipo_logradouro = "Praça"
   Case "PR"
  converte_tipo_logradouro = "Praia"
   Case "PRQ"
  converte_tipo_logradouro = "Parque"
   Case "R"
  converte_tipo_logradouro = "Rua"
   Case "RDV"
   converte_tipo_logradouro = "Rodovia"
   Case "TR"
   converte_tipo_logradouro = "Travessa"
   Case "VL"
   converte_tipo_logradouro = "Vila"
End Select

End Function

Function critica_correio_eletronico(correio As String) As Integer

critica_correio_eletronico = True
Dim i As Integer
Dim achou As Integer
Dim local_do_Asteristico As Integer
Dim local_do_ponto As Integer

achou = 0

 For i = 1 To Len(correio)
      If Mid$(correio, i, 1) = "@" Then
        achou = 1
      End If
 Next i
If achou = 0 Then
  critica_correio_eletronico = False
  Exit Function
End If

local_do_Asteristico = InStr(Trim$(correio), "@")
local_do_ponto = InStr(Trim$(correio), ".")

If local_do_ponto < local_do_Asteristico Then
  critica_correio_eletronico = False
  Exit Function
End If


End Function
Function critica_valor_com_caracteres(VARIAVEL As String) As Integer
critica_valor_com_caracteres = True
Dim i As Integer
Dim achou As Integer
achou = 0

 For i = 1 To Len(VARIAVEL)
      If Mid$(VARIAVEL, i, 1) = "," Then
        achou = achou + 1
      End If
  Next i

If achou > 1 Then
  critica_valor_com_caracteres = False
End If

End Function

Function critica_telefone(ddd As String, telefone As String) As Integer
critica_telefone = 0

If Trim(ddd) <> "" Then
  If Left(ddd, 1) <> "0" Then
     critica_telefone = 3
      Exit Function
  End If
End If

 If Trim(ddd) <> "" And Len(Trim(ddd)) >= 3 Then
   If Trim(telefone) = "" Or Len(Trim(telefone)) < 6 Then
      critica_telefone = 1
      Exit Function
   End If
 End If
 
 If Trim(ddd) = "" Then
   If Trim(telefone) <> "" And Len(Trim(telefone)) >= 6 Then
      critica_telefone = 2
      Exit Function
   End If
 End If
 
 
 
End Function

Function enviar_mensagem(mensagem As String) As Integer


 enviar_mensagem = MsgBox(mensagem, 33 + vbDefaultButton2, NOMBRE_SISTEMA_DEFAULT)
 
 
End Function

Sub limpa_baloes()

MDIdec.lbl_cadastro.Visible = False
MDIdec.lbl_declaracao.Visible = False
MDIdec.lbl_Diretorio.Visible = False
MDIdec.lbl_reparar.Visible = False
MDIdec.lbl_ajuda.Visible = False
MDIdec.lbl_sair.Visible = False

End Sub

Sub monta_cabecalho_FAtu(periodo As String)

   If status = 2 Then
      If Mid$(Trim(periodo), 5, 2) = "03" Then
         ArrayFatura(0, 0) = "Janeiro  "
         ArrayFatura(0, 1) = "Fevereiro"
         ArrayFatura(0, 2) = "Março    "
         ArrayFatura(0, 3) = "Total    "
      ElseIf Mid$(Trim(periodo), 5, 2) = "06" Then
         ArrayFatura(0, 0) = "Abril    "
         ArrayFatura(0, 1) = "Maio     "
         ArrayFatura(0, 2) = "Junho    "
         ArrayFatura(0, 3) = "Total    "
    ElseIf Mid$(Trim(periodo), 5, 2) = "09" Then
         ArrayFatura(0, 0) = "Julho    "
         ArrayFatura(0, 1) = "Agosto   "
         ArrayFatura(0, 2) = "Setembro "
         ArrayFatura(0, 3) = "Total    "
         
    ElseIf Mid$(Trim(periodo), 5, 2) = "12" Then
         ArrayFatura(0, 0) = "Outubro  "
         ArrayFatura(0, 1) = "Novembro "
         ArrayFatura(0, 2) = "Dezembro "
         ArrayFatura(0, 3) = "Total    "
    End If
  
   End If
   
   
End Sub
Sub cargar_Array_faturamento(inscricao As String, periodo As String, Substitutiva As Integer)

 Dim estado As String
 
  Dim i As Long
   
  query = "Select *  From tab_detalhe_declaracao " _
 & "Where tdd_num_inscricao  = '" & inscricao & "'" _
 & " and tdd_periodo_inicial = '" & periodo & "'" _
 & " and tdd_substitutiva = " & Substitutiva

 query = query & " and tdd_item between 223 and 242 "

  ReDim ArrayFatura(0 To 3, 0 To 3)
  ReDim Preserve ArrayFatura(0 To 3, 0 To 3)
  
   MaxCol_fat = 4
   MaxRow_fat = 4
   Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
   If seleccion.BOF Or seleccion.EOF Then
        seleccion.Close
        Exit Sub
    End If
    
   For i = 0 To 3
       ArrayFatura(1, i) = "0,00"
    Next i
   For i = 0 To 3
       ArrayFatura(2, i) = "0,00"
    Next i
   For i = 0 To 3
       ArrayFatura(3, i) = "0,00"
    Next i

      Call monta_cabecalho_FAtu(periodo)

    seleccion.MoveFirst
    
    i = 0
    While Not seleccion.EOF
      Select Case seleccion("tdd_item")
'       Case Is = 217   '  coluna 1
'         ArrayFatura(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'       Case Is = 218
'         ArrayFatura(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'       Case Is = 219
'         ArrayFatura(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'       Case Is = 220
'         ArrayFatura(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'       Case Is = 221
'         ArrayFatura(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'       Case Is = 222
'         ArrayFatura(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")

         'coluna 2
       Case Is = 223
         ArrayFatura(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 224
         ArrayFatura(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 225
         ArrayFatura(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'       Case Is = 226
'         ArrayFatura(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'       Case Is = 227
''         ArrayFatura(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
 '      Case Is = 228
 '        ArrayFatura(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 229
         ArrayFatura(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
     'coluna 3
       Case Is = 230
         ArrayFatura(2, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 231
         ArrayFatura(2, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 232
         ArrayFatura(2, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'       Case Is = 233
'         ArrayFatura(2, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'       Case Is = 234
'         ArrayFatura(2, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
'       Case Is = 235
'         ArrayFatura(2, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
     'coluna 4
       Case Is = 236
         ArrayFatura(3, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 237
         ArrayFatura(3, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 238
         ArrayFatura(3, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
 '      Case Is = 239
 '        ArrayFatura(3, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
 '      Case Is = 240
 '        ArrayFatura(3, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
 '      Case Is = 241
 '        ArrayFatura(3, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 242
         ArrayFatura(3, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       End Select
        seleccion.MoveNext
        i = i + 1
    Wend
    seleccion.Close
    
 
End Sub
Sub cargar_Array_recolhimento(inscricao As String, periodo As String, Substitutiva As Integer)
  Dim i As Long
   
 query = "Select *  From tab_detalhe_declaracao " _
 & "Where tdd_num_inscricao  = '" & inscricao & "'" _
 & " and tdd_periodo_inicial = '" & periodo & "'" _
 & " and tdd_substitutiva = " & Substitutiva
 query = query & " and tdd_item between 300 and 325"
 
 ReDim ArrayRec(0 To 2, 0 To 10)
 ReDim Preserve ArrayRec(0 To 2, 0 To 10)
  
 MaxCol_rec = 3
 MaxRow_rec = 11
  
  Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
    If seleccion.BOF Or seleccion.EOF Then
        seleccion.Close
        monta_cabecalho_recolhimento
        Exit Sub
    End If

    seleccion.MoveFirst
    
    monta_cabecalho_recolhimento
    i = 0
    While Not seleccion.EOF
      Select Case seleccion("tdd_item")
       Case Is = 300
            VALOR300 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 301
           VALOR301 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 302
           VALOR302 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 303
           VALOR303 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 304   '  coluna 1
         ArrayRec(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 305
         ArrayRec(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 306
         ArrayRec(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 307
         ArrayRec(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 308
         ArrayRec(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 309
         ArrayRec(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 310
         ArrayRec(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 311
         ArrayRec(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 312
         ArrayRec(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 313
         ArrayRec(1, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 314
         ArrayRec(1, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 315
         ArrayRec(2, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 316
         ArrayRec(2, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 317
         ArrayRec(2, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 318
         ArrayRec(2, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 319
         ArrayRec(2, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 320
         ArrayRec(2, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 321
         ArrayRec(2, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 322
         ArrayRec(2, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 323
         ArrayRec(2, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 324
         ArrayRec(2, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 325
         ArrayRec(2, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
         
      End Select
        seleccion.MoveNext
        i = i + 1
    Wend
    seleccion.Close
    
    MaxCol_rec = 2
    MaxRow_rec = 11


End Sub

Sub cargar_Array_DetDesp(inscricao As String, periodo As String, Substitutiva As Integer)
 
 Dim i As Long
 
 query = "Select *  From tab_detalhe_declaracao " _
 & "Where tdd_num_inscricao  = '" & inscricao & "'" _
 & " and tdd_periodo_inicial = '" & periodo & "'" _
 & " and tdd_substitutiva = " & Substitutiva
  
 query = query & " and tdd_item between 400 and 424 "
  
  ReDim ArrayDetDesp(0 To 1, 0 To 14)
  ReDim Preserve ArrayDetDesp(0 To 1, 0 To 14)
    MaxCol_DetDesp = 2
    MaxRow_DetDesp = 15

    Set seleccion = dbgim.OpenRecordset(query, dbOpenSnapshot)
    If seleccion.BOF Or seleccion.EOF Then
        seleccion.Close
        monta_cabecalho_DetDesp
        Exit Sub
    End If
    
'    For i = 0 To 14
'       ArrayDetDesp(1, i) = "0,00"
'    Next i

    seleccion.MoveFirst
    
    monta_cabecalho_DetDesp
    i = 0
    While Not seleccion.EOF
      Select Case seleccion("tdd_item")
       Case Is = 400
          VALOR400 = seleccion("tdd_valor")
       Case Is = 401   '  coluna 1
         ArrayDetDesp(1, 0) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 402
         ArrayDetDesp(1, 1) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 403
         ArrayDetDesp(1, 2) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 404
         ArrayDetDesp(1, 3) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 405
         ArrayDetDesp(1, 4) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 406
         ArrayDetDesp(1, 5) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 407
         ArrayDetDesp(1, 6) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 408
         ArrayDetDesp(1, 7) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 409
         ArrayDetDesp(1, 8) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 410
         ArrayDetDesp(1, 9) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 411
         ArrayDetDesp(1, 10) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 412
         ArrayDetDesp(1, 11) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 413
         ArrayDetDesp(1, 12) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 414
         ArrayDetDesp(1, 13) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 415
         ArrayDetDesp(1, 14) = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 416
          VALOR416 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 417
          VALOR417 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 418
          VALOR418 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 419
          VALOR419 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 420
          VALOR420 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 421
          VALOR421 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 422
          VALOR422 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 423
          VALOR423 = Format(seleccion("tdd_valor"), "###,###,##0.#0")
       Case Is = 424
          VALOR424 = Format(seleccion("tdd_valor"), "###,###,##0.#0")

      End Select
       seleccion.MoveNext
        i = i + 1
    Wend
      
    MaxCol_DetDesp = 2
    MaxRow_DetDesp = 13
    seleccion.Close
End Sub

Function critica_livro_registro_inicial(txt416 As String, txt417 As String, txt418 As String, txt419 As String) As Integer
Dim total As Double
 critica_livro_registro_inicial = True
If Trim(txt416) = "" Then
  txt416 = "0"
End If
If Trim(txt417) = "" Then
  txt417 = "0"
End If
If Trim(txt418) = "" Then
  txt418 = "0"
End If
If Trim(txt419) = "" Then
  txt419 = "0"
End If

total = CCur(CDbl(txt416)) + CCur(CDbl(txt417)) + CCur(CDbl(txt418))
If CCur(CDbl(total)) <> CCur(CDbl(txt419)) Then
   critica_livro_registro_inicial = False
End If


End Function
Function critica_livro_registro_final(txt420 As String, txt421 As String, txt422 As String, txt423 As String) As Integer
Dim total As Double
 critica_livro_registro_final = True
If Trim(txt420) = "" Then
  txt420 = "0"
End If
If Trim(txt421) = "" Then
  txt421 = "0"
End If
If Trim(txt422) = "" Then
  txt422 = "0"
End If
If Trim(txt423) = "" Then
  txt423 = "0"
End If

total = CCur(CDbl(txt420)) + CCur(CDbl(txt421)) + CCur(CDbl(txt422))
If CCur(CDbl(total)) <> CCur(CDbl(txt423)) Then
   critica_livro_registro_final = False
End If


End Function
Function critica_detalhe_despesas(grid_DetDesp() As Variant) As Integer
Dim coluna As Integer
Dim linha As Integer
Dim total As Double
Dim valor As Double
critica_detalhe_despesas = True
total = 0#
     For coluna = 1 To 1
          For linha = 0 To 13
             If Trim(grid_DetDesp(coluna, linha)) <> "" Then
               total = total + CCur(CDbl(grid_DetDesp(coluna, linha)))
             End If
          Next linha
       Next coluna

valor = 0#
If Trim(grid_DetDesp(1, 14)) <> "" Then
   valor = Trim(grid_DetDesp(1, 14))
End If

If CCur(CDbl(total)) <> valor Then
   critica_detalhe_despesas = False
End If


End Function

Function critica_detalhe_demonstrativos(GridArray() As Variant, linha_erro As Integer, cod_erro As Integer, demonstrativo As String, msgt As String) As Integer
'Function critica_detalhe_demonstrativos(GridArray() As Variant, linha_erro As Integer, cod_erro As Integer, demonstrativo As String, txt012 As String, txt024 As String, txt036 As String, txt048 As String, txt060 As String, msgt As String) As Integer

Dim linha As Integer
Dim coluna As Integer
On Error GoTo erro_xxx
g_tot_val_contabil = 0#
g_tot_base_calc = 0#
g_tot_imp_creditado = 0#
g_tot_isentas = 0#
g_tot_outras = 0#
g_tot_val_contabil_saida = 0#
g_tot_val_para_faturamento = 0#
g_tot_val_imp_debitado = 0#
g_tot_val_imp_cred_entradas = 0#
Dim VALOR_1 As String
Dim VALOR_2 As String
Dim VALOR_3 As String
Dim VALOR_4 As String
Dim VALOR_5 As String

critica_detalhe_demonstrativos = True

'critica dos valores contabeis  definicao - valores contabeis = base de calculo + isentas + outras
    For linha = 0 To 11
          linha_erro = linha + 1
          cod_erro = linha + 1
          If Trim(GridArray(1, linha)) = "" Then
              VALOR_1 = "0"
          Else
              VALOR_1 = Trim(GridArray(1, linha))
          End If
          If Trim(GridArray(2, linha)) = "" Then
              VALOR_2 = "0"
          Else
              VALOR_2 = Trim(GridArray(2, linha))
          End If
          If Trim(GridArray(4, linha)) = "" Then
              VALOR_4 = "0"
          Else
              VALOR_4 = Trim(GridArray(4, linha))
          End If
          If Trim(GridArray(5, linha)) = "" Then
              VALOR_5 = "0"
          Else
              VALOR_5 = Trim(GridArray(5, linha))
          End If
          
         critica_detalhe_demonstrativos = crit_val_contabeis(CDbl(VALOR_1), CDbl(VALOR_2), CDbl(VALOR_4), CDbl(VALOR_5))
         If Not critica_detalhe_demonstrativos Then Exit Function
    Next linha
   
    linha_erro = 12
    cod_erro = 12
    critica_detalhe_demonstrativos = crit_val_contabeis(CDbl(VALOR_1), CDbl(VALOR_2), CDbl(VALOR_4), CDbl(VALOR_5))
    If Not critica_detalhe_demonstrativos Then Exit Function
    
    'critica dos valores do Imposto  de 0 a 25% da base de calculo
Dim vinte_e_cinco_por_cento As Double

    For linha = 0 To 10
        linha_erro = linha + 1
          If Trim(GridArray(2, linha)) = "" Then
              VALOR_2 = "0"
          Else
              VALOR_2 = Trim(GridArray(2, linha))
          End If
          If Trim(GridArray(3, linha)) = "" Then
              VALOR_3 = "0"
          Else
              VALOR_3 = Trim(GridArray(3, linha))
          End If
         vinte_e_cinco_por_cento = VALOR_2 * 0.25
         If VALOR_3 > vinte_e_cinco_por_cento Then
            cod_erro = 15
            critica_detalhe_demonstrativos = False
            Exit Function
         End If
    Next linha
    
'critica de totais
For coluna = 1 To 5
    For linha = 0 To 10
        If Trim(GridArray(coluna, linha)) = "" Then
            VALOR_1 = "0"
        Else
            VALOR_1 = Trim(GridArray(coluna, linha))
        End If
        Select Case coluna
           Case Is = 1
             g_tot_val_contabil = g_tot_val_contabil + CDbl(VALOR_1)
             If linha = 0 Then
                g_tot_val_para_faturamento = g_tot_val_para_faturamento + CDbl(VALOR_1)
             ElseIf linha = 4 Then
                g_tot_val_para_faturamento = g_tot_val_para_faturamento + CDbl(VALOR_1)
             ElseIf linha = 8 Then
                g_tot_val_para_faturamento = g_tot_val_para_faturamento + CDbl(VALOR_1)
             End If
           Case Is = 2
             g_tot_base_calc = g_tot_base_calc + CDbl(VALOR_1)
           Case Is = 3
             g_tot_imp_creditado = g_tot_imp_creditado + CDbl(VALOR_1)
           Case Is = 4
             g_tot_isentas = g_tot_isentas + CDbl(VALOR_1)
           Case Is = 5
             g_tot_outras = g_tot_outras + CDbl(VALOR_1)
         End Select
         
    Next linha
 Next coluna
 If demonstrativo = "S" Then
    
    g_tot_val_contabil_saida = g_tot_val_contabil
    g_tot_val_imp_debitado = g_tot_imp_creditado
 Else
    g_tot_val_imp_cred_entradas = g_tot_imp_creditado
 End If
    
 If Trim(GridArray(1, 11)) = "" Then
    VALOR_1 = "0"
 Else
    VALOR_1 = Trim(GridArray(1, 11))
 End If
msgt = ""
If g_tot_val_contabil <> CCur(CDbl(VALOR_1)) Then
   critica_detalhe_demonstrativos = False
   linha_erro = 13
   cod_erro = 13
   If demonstrativo = "S" Then
      msgt = "Erro de Totalização dos Valores Contabeis no Demonstrativo das Operações (Saídas)."
   Else
      msgt = "Erro de Totalização dos Valores Contabeis no Demonstrativo das Operações (Entradas)."
   End If
   Exit Function
End If

 If Trim(GridArray(2, 11)) = "" Then
    VALOR_1 = "0"
 Else
    VALOR_1 = Trim(GridArray(2, 11))
 End If
 If g_tot_base_calc <> CCur(CDbl(VALOR_1)) Then
    critica_detalhe_demonstrativos = False
    cod_erro = 13
    linha_erro = 13
    If demonstrativo = "S" Then
       msgt = "Erro de Totalização da Base de Cálculo no Demonstrativo das Operações (Saídas)."
    Else
       msgt = "Erro de Totalização da Base de Cálculo no Demonstrativo das Operações (Entradas)."
    End If
    Exit Function
 End If
' If Len(Trim(txt036)) < 3 Then
 If Trim(GridArray(3, 11)) = "" Then
    VALOR_1 = "0"
 Else
    VALOR_1 = Trim(GridArray(3, 11))
 End If
If g_tot_imp_creditado <> CCur(CDbl(VALOR_1)) Then
  critica_detalhe_demonstrativos = False
  linha_erro = 13
  cod_erro = 13
 If demonstrativo = "S" Then
    msgt = "Erro de Totalização do Imposto Debitado no Demonstrativo das Operações (Saídas)."
 Else
   msgt = "Erro de Totalização do Imposto Creditado no Demonstrativo das Operações (Entradas)."
 End If
  Exit Function
End If
If Trim(GridArray(4, 11)) = "" Then
'If Trim(txt048) = "" Then
    VALOR_1 = "0"
 Else
    VALOR_1 = Trim(GridArray(4, 11))
 End If
If g_tot_isentas <> CCur(CDbl(VALOR_1)) Then
   critica_detalhe_demonstrativos = False
   linha_erro = 13
   cod_erro = 13
  If demonstrativo = "S" Then
     msgt = "Erro de Totalização da Isentas não Tributadas no Demonstrativo das Operações (Saídas)."
  Else
     msgt = "Erro de Totalização da Isentas não Tributadas no Demonstrativo das Operações (Entradas)."
  End If
  Exit Function
End If
If Trim(GridArray(5, 11)) = "" Then
'If Trim(txt060) = "" Then
    VALOR_1 = "0"
 Else
    VALOR_1 = Trim(GridArray(5, 11))
 End If
If g_tot_outras <> CCur(CDbl(VALOR_1)) Then
   critica_detalhe_demonstrativos = False
   linha_erro = 13
   cod_erro = 13
  If demonstrativo = "S" Then
     msgt = "Erro de Totalização de Outras no Demonstrativo das Operações (Saídas)."
  Else
     msgt = "Erro de Totalização de Outras no Demonstrativo das Operações (Entradas)."
  End If
   Exit Function
End If

 
If Not critica_detalhe_demonstrativos Then Exit Function

erro_xxx:
If Err <> 0 Then
     MsgBox "Error Message :" & Str(Err) & " " & Error, MB_ICONEXCLAMATION, NOMBRE_SISTEMA_DEFAULT
    Resume Next
End If
End Function
Function formata_p_impressao_2(VARIAVEL As Variant, tam As Integer) As String

  Select Case Len(Trim(VARIAVEL))
       Case Is = 1
          formata_p_impressao_2 = String$((tam - 1), " ") & Trim(VARIAVEL)
      Case Is = 2
         formata_p_impressao_2 = String$((tam - 2), " ") & Trim(VARIAVEL)
       Case Is = 3
          formata_p_impressao_2 = String$((tam - 3), " ") & Trim(VARIAVEL)
      Case Is = 4
         formata_p_impressao_2 = String$((tam - 4), " ") & Trim(VARIAVEL)
       Case Is = 5
          formata_p_impressao_2 = String$((tam - 5), " ") & Trim(VARIAVEL)
      Case Is = 6
         formata_p_impressao_2 = String$((tam - 6), " ") & Trim(VARIAVEL)
       Case Is = 7
          formata_p_impressao_2 = String$((tam - 7), " ") & Trim(VARIAVEL)
      Case Is = 8
         formata_p_impressao_2 = String$((tam - 8), " ") & Trim(VARIAVEL)
      Case Is = 9
         formata_p_impressao_2 = String$((tam - 9), " ") & Trim(VARIAVEL)
      Case Is = 10
         formata_p_impressao_2 = String$((tam - 10), " ") & Trim(VARIAVEL)
       Case Is = 11
         formata_p_impressao_2 = String$((tam - 11), " ") & Trim(VARIAVEL)
       Case Is = 12
         formata_p_impressao_2 = String$((tam - 12), " ") & Trim(VARIAVEL)
       Case Is = 13
         formata_p_impressao_2 = String$((tam - 13), " ") & Trim(VARIAVEL)
       Case Is = 14
         formata_p_impressao_2 = String$((tam - 14), " ") & Trim(VARIAVEL)
       Case Is = 15
         formata_p_impressao_2 = String$((tam - 15), " ") & Trim(VARIAVEL)
       Case Is = 16
         formata_p_impressao_2 = String$((tam - 16), " ") & Trim(VARIAVEL)
       Case Is = 17
         formata_p_impressao_2 = String$((tam - 17), " ") & Trim(VARIAVEL)
       Case Is = 18
         formata_p_impressao_2 = String$((tam - 18), " ") & Trim(VARIAVEL)
       Case Is = 19
         formata_p_impressao_2 = String$((tam - 19), " ") & Trim(VARIAVEL)
       Case Is = 20
         formata_p_impressao_2 = String$((tam - 20), " ") & Trim(VARIAVEL)

    End Select
       
End Function
Function critica_demonstrativos_EPP_micro(GridArray() As Variant) As Integer
Dim coluna As Integer
Dim linha As Integer
Dim total As Double
On Error GoTo erro_xx
critica_demonstrativos_EPP_micro = True
total = 0#
For coluna = 1 To 1
   For linha = 0 To 10
      If Trim(GridArray(coluna, linha)) <> "" Then
         total = total + CCur(CDbl(GridArray(coluna, linha)))
      End If
   Next linha
Next coluna

If Trim(GridArray(1, 11)) <> "" Then
   If CCur(CDbl(total)) <> CCur(CDbl(GridArray(1, 11))) Then
      critica_demonstrativos_EPP_micro = False
   End If
Else
  If total > 0 Then
      critica_demonstrativos_EPP_micro = False
  End If
End If

erro_xx:
If Err <> 0 Then
     MsgBox "Error Message :" & Str(Err) & " " & Error, MB_ICONEXCLAMATION, NOMBRE_SISTEMA_DEFAULT
    Resume Next
End If
End Function

Function critica_apuracao_icms_normal(FATURAMENTO As String, grid_ApurICMS() As Variant, tipo_erro As Integer) As Integer
Dim total_debito As Double
Dim total_credito As Double
Dim tot_deb_menos_credito As Double
Dim tot_cred_menos_debito As Double
Dim sal_devedor_menos_deducoes As Double
Dim coluna As Integer
Dim linha As Integer
Dim faturamento_wk As Double

critica_apuracao_icms_normal = True
If Trim(FATURAMENTO) = "" Then
  faturamento_wk = 0#
Else
  faturamento_wk = CDbl(FATURAMENTO)
End If

If faturamento_wk < CCur(CDbl(g_tot_val_para_faturamento)) Then
  critica_apuracao_icms_normal = False
  tipo_erro = 1  'erro de faturamento
  Exit Function
End If

total_debito = 0#
For coluna = 1 To 1
  For linha = 0 To 2
    If Trim(grid_ApurICMS(coluna, linha)) <> "" And Trim(grid_ApurICMS(coluna, linha)) <> "0,00" Then
       total_debito = total_debito + CDbl(grid_ApurICMS(coluna, linha))
    End If
  Next linha
Next coluna
 
If CCur(CDbl(grid_ApurICMS(1, 3))) <> CCur(CDbl(total_debito)) Then
  critica_apuracao_icms_normal = False
  tipo_erro = 2  'erro de total de debito
  Exit Function
End If

total_credito = 0#
For coluna = 1 To 1
  For linha = 4 To 5
    If Trim(grid_ApurICMS(coluna, linha)) <> "" And Trim(grid_ApurICMS(coluna, linha)) <> "0,00" Then
      total_credito = total_credito + CDbl(grid_ApurICMS(coluna, linha))
    End If
  Next linha
Next coluna
 
For coluna = 1 To 1
  For linha = 11 To 12
    If Trim(grid_ApurICMS(coluna, linha)) <> "" And Trim(grid_ApurICMS(coluna, linha)) <> "0,00" Then
      total_credito = total_credito + CDbl(grid_ApurICMS(coluna, linha))
    End If
  Next linha
Next coluna

If CCur(CDbl(grid_ApurICMS(1, 13))) <> CCur(CDbl(total_credito)) Then
  critica_apuracao_icms_normal = False
  tipo_erro = 3  'erro de total de credito
  Exit Function
End If

End Function

Function critica_apuracao_icms_micro(grid_fatura() As Variant, linha_erro As Integer, tipo_erro As Integer) As Integer
Dim total_faturamento As Double
Dim total_imposto As Double
Dim aliquota_minima As Double
Dim aliquota_maxima As Double
Dim fatu_vezes_aliquota As Double
Dim folga As Double
Dim folga_minima As Double
Dim folga_maxima As Double
Dim fatura As Double
Dim imposto As Double
Dim ALIQUOTA As Double
Dim linha As Integer
Dim coluna As Integer
aliquota_minima = 0.3
aliquota_maxima = 1
folga_minima = -1#
folga_maxima = 1#

critica_apuracao_icms_micro = True
total_imposto = 0#
total_faturamento = 0#
   For coluna = 1 To 3
      For linha = 0 To 2
         linha_erro = linha + 1
            Select Case coluna
               Case Is = 1
                    If Trim(grid_fatura(coluna, linha)) <> "" Then
                      total_faturamento = total_faturamento + CCur(CDbl(grid_fatura(coluna, linha)))
                    ElseIf Trim(grid_fatura((coluna + 1), linha)) <> "" And Val(Trim(grid_fatura((coluna + 1), linha))) <> 0 Then
                         critica_apuracao_icms_micro = False
                         tipo_erro = 15                'erro de faturamento de microempresa
                         Exit Function
                    ElseIf Trim(grid_fatura((coluna + 2), linha)) <> "" And Val(Trim(grid_fatura((coluna + 2), linha))) <> 0 Then
                         critica_apuracao_icms_micro = False
                         tipo_erro = 15                'erro de faturamento de microempresa
                         Exit Function
                    End If
               Case Is = 2
               
                If Trim(grid_fatura(coluna, linha)) <> "" And Trim(grid_fatura(coluna, linha)) <> "0,00" Then
                   If CCur(CDbl(grid_fatura(coluna, linha))) < aliquota_minima Then
                      critica_apuracao_icms_micro = False
                      tipo_erro = 13                'erro de aliquota de microempresa
                      Exit Function
                   End If
                   If CCur(CDbl(grid_fatura(coluna, linha))) > aliquota_maxima Then
                      critica_apuracao_icms_micro = False
                      tipo_erro = 13                'erro de aliquotade microempresa
                      Exit Function
                   End If
                Else
                   If Trim(grid_fatura((coluna - 1), linha)) <> "" And Val(Trim(grid_fatura((coluna - 1), linha))) <> 0 Then
                       critica_apuracao_icms_micro = False
                       tipo_erro = 13                'erro de aliquotade microempresa
                       Exit Function
                   ElseIf Trim(grid_fatura((coluna + 1), linha)) <> "" And Val(Trim(grid_fatura((coluna + 1), linha))) <> 0 Then
                       critica_apuracao_icms_micro = False
                       tipo_erro = 13                'erro de aliquotade microempresa
                       Exit Function
                       
                   End If
                End If
            Case Is = 3
                If Trim(grid_fatura((coluna), linha)) <> "" Then
                   total_imposto = total_imposto + CCur(CDbl(grid_fatura(coluna, linha)))
                ElseIf Trim(grid_fatura((coluna - 1), linha)) <> "" And Val(Trim(grid_fatura((coluna - 1), linha))) <> 0 Then
                     critica_apuracao_icms_micro = False
                     tipo_erro = 16                'erro de imposto de microempresa
                     Exit Function
                ElseIf Trim(grid_fatura((coluna - 2), linha)) <> "" And Val(Trim(grid_fatura((coluna - 2), linha))) <> 0 Then
                     critica_apuracao_icms_micro = False
                     tipo_erro = 16                'erro de imposto de microempresa
                     Exit Function
                End If
            End Select
      Next linha
   Next coluna
   
If CCur(CDbl(total_faturamento)) <> CCur(CDbl(grid_fatura(1, 3))) Then
   critica_apuracao_icms_micro = False
   tipo_erro = 11                   '"erro total de faturamento de microempresa"
   Exit Function
End If

If CCur(CDbl(total_imposto)) <> CCur(CDbl(grid_fatura(3, 3))) Then
   critica_apuracao_icms_micro = False
   tipo_erro = 12                   '"erro total de imposto de microempresa"
   Exit Function
End If

 'For Coluna = 1 To 3
    For linha = 0 To 2
       linha_erro = linha + 1
       fatura = 0#
       If Trim(grid_fatura(1, linha)) <> "" Then
          fatura = Trim(grid_fatura(1, linha))
       End If
       ALIQUOTA = 0#
       If Trim(grid_fatura(2, linha)) <> "" Then
          ALIQUOTA = Trim(grid_fatura(2, linha))
       End If
       imposto = 0#
       If Trim(grid_fatura(3, linha)) <> "" Then
          imposto = Trim(grid_fatura(3, linha))
       End If
       fatu_vezes_aliquota = (CCur(fatura) * CCur(ALIQUOTA)) / 100
       If CCur(CDbl(fatu_vezes_aliquota)) <> CCur(imposto) Then
           folga = CCur(CDbl(fatu_vezes_aliquota)) - CCur(imposto)
            If folga > 0 Then
               If folga > folga_maxima Then
                  critica_apuracao_icms_micro = False
                  tipo_erro = 14                   'erro de faturamento * aliquota diferente do imp com folga
                  Exit Function
                End If
             Else
                If folga < 0 Then
                   If folga < folga_minima Then
                      critica_apuracao_icms_micro = False
                      tipo_erro = 14                   'erro de faturamento * aliquota diferente do imp com folga
                      Exit Function
                    End If
                 End If
             End If
        End If
        
    Next linha
' Next Coluna

End Function
Function critica_apuracao_icms_EPP(FATURAMENTO As String, ALIQUOTA As String, imposto As String, tipo_erro As Integer) As Integer
Dim aliquota_minima As Double
Dim aliquota_maxima As Double
Dim fatu_vezes_aliquota As Double
Dim folga As Double
Dim folga_minima As Double
Dim folga_maxima As Double

critica_apuracao_icms_EPP = True
aliquota_minima = 1.5
aliquota_maxima = 2
folga_minima = -1#
folga_maxima = 1#

If Trim(FATURAMENTO) = "" Then
   FATURAMENTO = "0"
End If
If Trim(ALIQUOTA) = "" Then
   ALIQUOTA = "0"
End If
If Trim(imposto) = "" Then
   imposto = "0"
End If


End Function
Sub monta_cabecalho_DetDesp()
   
   ArrayDetDesp(0, 0) = "Pró-Labore                    :"
   ArrayDetDesp(0, 1) = "Comissões, Salários,Ordenados :"
   ArrayDetDesp(0, 2) = "Combustíveis e Lubrificantes  :"
   ArrayDetDesp(0, 3) = "Encargos Sociais              :"
   ArrayDetDesp(0, 4) = "Tributos Federais             :"
   ArrayDetDesp(0, 5) = "Tributos Estaduais            :"
   ArrayDetDesp(0, 6) = "Tributos Municipais           :"
   ArrayDetDesp(0, 7) = "Água, Luz, Telefone           :"
   ArrayDetDesp(0, 8) = "Aluguéis                      :"
   ArrayDetDesp(0, 9) = "Serviços Profissionais        :"
   ArrayDetDesp(0, 10) = "Seguros                       :"
   ArrayDetDesp(0, 11) = "Fretes e carretos             :"
   ArrayDetDesp(0, 12) = "Despesas Financeiras          :"
   ArrayDetDesp(0, 13) = "Outras Despesas               :"
   ArrayDetDesp(0, 14) = "TOTAL                         :"

End Sub

Sub monta_cabecalho_recolhimento()

          ArrayRec(0, 0) = "Normal                                   :"
          ArrayRec(0, 1) = "Empresa de Pequeno Porte                 :"
          ArrayRec(0, 2) = "Micro Empresa                            :"
          ArrayRec(0, 3) = "Substituicao de Saidas                   :"
          ArrayRec(0, 4) = "Substituicao de Entradas                 :"
          ArrayRec(0, 5) = "Antecipacao Parcial                      :"
          ArrayRec(0, 6) = "Diferencial de Alíquota                  :"
          ArrayRec(0, 7) = "Importações                              :"
'          ArrayRec(0, 8) = "Simplificado                             :"
'          ArrayRec(0, 9) = "Incentivado                              :"
          ArrayRec(0, 8) = "Total                                    :"

End Sub
Sub monta_cabecalho_Entrada()

          ArrayEntra(0, 0) = "Compras no Estado           :"
          ArrayEntra(0, 1) = "Transferencias no Estado    :"
          ArrayEntra(0, 2) = "Devoluções no Estado        :"
          ArrayEntra(0, 3) = "Outras no Estado            :"
          ArrayEntra(0, 4) = "Compras de outros Estados   :"
          ArrayEntra(0, 5) = "Transf. de outros Estados   :"
          ArrayEntra(0, 6) = "Devoluções de outros Estados:"
          ArrayEntra(0, 7) = "Outras de outros Estados    :"
          ArrayEntra(0, 8) = "Compras no Exterior         :"
          ArrayEntra(0, 9) = "Devoluções do Exterior      :"
          ArrayEntra(0, 10) = "Outras entradas do Exterior :"
          ArrayEntra(0, 11) = "Totais de Entradas          :"
          ArrayEntra(0, 12) = "                             "
          
End Sub
Sub monta_cabecalho_ApurICMS()

   'ReDim CApurICMS(0 To 5, 0 To 12)
   apurICMS(0, 0) = "Débito do Imposto:  Por Saídas ou Prestações        :"
   apurICMS(0, 1) = "                    Outros Débitos                  :"
   apurICMS(0, 2) = "                    Estorno de Créditos             :"
   apurICMS(0, 3) = "                    Total dos Débitos               :"
   apurICMS(0, 4) = "Crédito do Imposto: Por Entradas ou Aquisições      :"
   apurICMS(0, 5) = "                    = Outros Créditos               :"
   apurICMS(0, 6) = "                    + Restituição                   :"
   apurICMS(0, 7) = "                    + Crédito Presumido             :"
   apurICMS(0, 8) = "                    + Antecip.Parcial Interestadual :"
   apurICMS(0, 9) = "                    + Antecipação Total             :"
  apurICMS(0, 10) = "                    + Diferencial de Aliquota/Outros:"
  apurICMS(0, 11) = "                    Estorno de Débitos              :"
  apurICMS(0, 12) = "                    Saldo credor do período anterior:"
  apurICMS(0, 13) = "                    Total dos Créditos              :"
  apurICMS(0, 14) = "Apuração do Imposto:Saldo Devedor                   :"
  apurICMS(0, 15) = "                    Deduções                        :"
  apurICMS(0, 16) = "                    Imposto a Recolher              :"
  apurICMS(0, 17) = "                    Saldo credor a transferir       :"



   
End Sub
Sub monta_Arraysaidas()

          ArraySaidas(0, 0) = "Vendas para o Estado        :"
          ArraySaidas(0, 1) = "Transferencias para o Estado:"
          ArraySaidas(0, 2) = "Devoluções para o Estado    :"
          ArraySaidas(0, 3) = "Outras para o Estado        :"
          ArraySaidas(0, 4) = "Vendas para outros Estados  :"
          ArraySaidas(0, 5) = "Transf. para outros Estados :"
          ArraySaidas(0, 6) = "Devoluções p/outros Estados :"
          ArraySaidas(0, 7) = "Outras para outros Estados  :"
          ArraySaidas(0, 8) = "Vendas para o Exterior      :"
          ArraySaidas(0, 9) = "Devoluções para o Exterior  :"
         ArraySaidas(0, 10) = "Outras para o Exterior      :"
         ArraySaidas(0, 11) = "Totais de Saídas            :"

End Sub
Sub LIMPA_VARIAVEIS_IMPRESSAO()

 VALOR200 = ""
 VALOR214 = ""
 VALOR215 = ""
 VALOR216 = ""
 VALOR300 = ""
 VALOR301 = ""
 VALOR302 = ""
 VALOR303 = ""
 VALOR400 = ""

 VALOR416 = ""
 VALOR417 = ""
 VALOR418 = ""
 VALOR419 = ""
 VALOR420 = ""
 VALOR421 = ""
 VALOR422 = ""
 VALOR423 = ""
VALOR424 = ""
End Sub
Sub imprime_recolhimento_periodo(inscricao As String, Razao_social As String, periodo As String, Situacao As String, data_ger As String, CRC As String, ufCRC As String, Substitutiva As Integer)
Dim b As Integer
Printer.FontBold = True
Printer.FontSize = 9

Printer.Print Space(40) & "RECOLHIMENTO DO PERÍODO"
Printer.Print
Printer.FontBold = False
Printer.FontSize = 7
Printer.Print "                      " & Space(6) & "Espécie de Recolhimento " & Space(7) & "             " & Space(1) & "  Recolhidos             A Recolher                "
Printer.Print String(130, "=")

linhas_impressas = linhas_impressas + 4

X = 0
For a = 0 To 10
    If Len(Trim(ArrayRec(1, a))) < 21 Then
      ArrayRec(1, a) = Space(21 - Len(ArrayRec(1, a))) + ArrayRec(1, a)
    End If
    
    If Len(Trim(ArrayRec(2, a))) < 21 Then
      ArrayRec(2, a) = Space(21 - Len(ArrayRec(2, a))) + ArrayRec(2, a)
    End If
    
   If Trim(ArrayRec(1, a)) <> "" Or Trim(ArrayRec(2, a)) <> "" Then
        X = X + 1
        Printer.Print Space(15) & ArrayRec(0, a) & Space(5) & ArrayRec(1, a) & Space(5) & ArrayRec(2, a)
        linhas_impressas = linhas_impressas + 1

   End If
Next
If X = 0 Then
     Printer.Print Space(60) & "SEM MOVIMENTOS"
     linhas_impressas = linhas_impressas + 1

End If

End Sub

Sub imprime_faturamento_micro()
Dim b As Integer

Printer.FontBold = True
Printer.FontSize = 9


 Printer.Print Space(20) & "APURAÇÃO PARA CONTRIBUINTES ESPECIAIS DO ICMS - MICROEMPRESA"

Printer.FontBold = False
Printer.FontSize = 7

Printer.Print
Printer.Print Space(30) & "Mês      Faturamento do Mês(Receita Bruta)     Alíquota(%)                           Imposto     " 'EPP c/aliq.<17

linhas_impressas = linhas_impressas + 3

X = 0
For a = 0 To 3
    If Len(Trim(ArrayFatura(1, a))) < 20 Then
      ArrayFatura(1, a) = Space(20 - Len(ArrayFatura(1, a))) + ArrayFatura(1, a)
    End If
    If Len(Trim(ArrayFatura(3, a))) < 20 Then
      ArrayFatura(3, a) = Space(20 - Len(ArrayFatura(3, a))) + ArrayFatura(3, a)
    End If

    If a < 3 Then
        If Trim(ArrayFatura(1, a)) <> "" Then
            X = X + 1
            Printer.Print Space(30) & ArrayFatura(0, a) & Space(5) & ArrayFatura(1, a) & Space(17) & ArrayFatura(2, a) & Space(17) & ArrayFatura(3, a)
            linhas_impressas = linhas_impressas + 1
        End If
    Else
        X = X + 1
        Printer.Print Space(30) & ArrayFatura(0, a) & Space(5) & ArrayFatura(1, a) & Space(16) & Space(22) & ArrayFatura(3, a)
        linhas_impressas = linhas_impressas + 1
    End If
Next
If X = 0 Then
     Printer.Print Space(60) & "SEM MOVIMENTOS"
     linhas_impressas = linhas_impressas + 1

End If

End Sub
Sub imprime_entrada_ativo_fixo(inscricao As String, Razao_social As String, periodo As String, Situacao As String, data_ger As String, CRC As String, ufCRC As String, Substitutiva As Integer)


If status <> 2 Then
    Printer.FontBold = True
    Printer.FontSize = 9
    Printer.Print Space(35) & "ENTRADA DE ATIVO FIXO OU CONSUMO"
    Printer.Print
    linhas_impressas = linhas_impressas + 2
    Printer.FontBold = False
    Printer.FontSize = 7
    
    If Trim(VALOR300) = "" And Trim(VALOR301) = "" And Trim(VALOR302) = "" And Trim(VALOR303) = "" Then
       Printer.Print Space(60) & "SEM MOVIMENTOS"
       linhas_impressas = linhas_impressas + 1
    Else
      VALOR300 = Space(20 - Len(Trim(VALOR300))) & Trim(VALOR300)
      VALOR301 = Space(20 - Len(Trim(VALOR301))) & Trim(VALOR301)
      VALOR302 = Space(20 - Len(Trim(VALOR302))) & Trim(VALOR302)
      VALOR303 = Space(20 - Len(Trim(VALOR303))) & Trim(VALOR303)
      Printer.Print Space(40) & "Do Estado                                : " & Space(5) & VALOR300
      Printer.Print Space(40) & "De Outros Estados                        : " & Space(5) & VALOR301
      Printer.Print Space(40) & "Do Exterior                              : " & Space(5) & VALOR302
      Printer.Print Space(40) & "Total                                    : " & Space(5) & VALOR303
      linhas_impressas = linhas_impressas + 4
    End If
    
    Printer.FontBold = False
    Printer.Print String(130, "-")
    Printer.Print
    linhas_impressas = linhas_impressas + 1

End If

End Sub
Sub imprime_dem_saidas()
Dim b As Integer
Printer.FontBold = True
Printer.FontSize = 9

Printer.Print Space(32) & "DEMONSTRATIVO DAS OPERAÇÕES - SAÍDAS"

Printer.FontBold = False
Printer.FontSize = 7

Printer.Print "Natureza da Operação  " & Space(6) & "    Valores Contábeis" & Space(7) & " Base Calculo" & Space(3) & " Imposto Debitado        Isentas                   Outras"
Printer.Print Space(92) & "   não Tributadas"

Printer.Print String(130, "=")
X = 0
linhas_impressas = linhas_impressas + 3
c = 1
For a = 0 To 11
   For b = 1 To 5
    If Len(Trim(ArraySaidas(b, a))) < 20 Then
      ArraySaidas(b, a) = Space(20 - Len(Trim(ArraySaidas(b, a)))) + Trim(ArraySaidas(b, a))
    End If
   Next
   If Trim(ArraySaidas(1, a)) <> "" Or Trim(ArraySaidas(2, a)) <> "" Or Trim(ArraySaidas(3, a)) <> "" Or Trim(ArraySaidas(4, a)) <> "" Or Trim(ArraySaidas(5, a)) <> "" Then
        X = X + 1
'        Printer.Print ArraySaidas(0, a) & Space(1) & ArraySaidas(1, a) & Space(c) & ArraySaidas(2, a) & Space(c) & ArraySaidas(3, a) & Space(c) & ArraySaidas(4, a) & Space(c) & ArraySaidas(5, a)
        Printer.Print ArraySaidas(0, a) & ArraySaidas(1, a) & ArraySaidas(2, a) & ArraySaidas(3, a) & ArraySaidas(4, a) & ArraySaidas(5, a)
        linhas_impressas = linhas_impressas + 1

   End If
Next
If X = 0 Then
     Printer.Print Space(60) & "SEM MOVIMENTOS"
     linhas_impressas = linhas_impressas + 1

End If

End Sub
Sub imprime_dem_entradas()
Dim b As Integer
Printer.FontBold = True
Printer.FontSize = 9

'Imprime operações de Entrada

Printer.Print Space(30) & "DEMONSTRATIVO DAS OPERAÇÕES - ENTRADAS"

Printer.FontBold = False
Printer.FontSize = 7

Printer.Print "Natureza da Operação  " & Space(6) & "    Valores Contábeis" & Space(7) & " Base Calculo" & Space(1) & "  Imposto Creditado        Isentas                   Outras"
Printer.Print Space(92) & "   não Tributadas"

Printer.Print String(130, "=")

linhas_impressas = linhas_impressas + 3
c = 1
X = 0
For a = 0 To 12
   For b = 1 To 5
    If Len(Trim(ArrayEntra(b, a))) < 20 Then
      ArrayEntra(b, a) = Space(20 - Len(Trim(ArrayEntra(b, a)))) + Trim(ArrayEntra(b, a))
    End If
   Next
'   If Trim(ArrayEntra(1, a)) <> "" And Trim(ArrayEntra(2, a)) <> "" And Trim(ArrayEntra(3, a)) <> "" And Trim(ArrayEntra(4, a)) <> "" And Trim(ArrayEntra(5, a)) <> "" Then
   If Trim(ArrayEntra(1, a)) <> "" Or Trim(ArrayEntra(2, a)) <> "" Or Trim(ArrayEntra(3, a)) <> "" Or Trim(ArrayEntra(4, a)) <> "" Or Trim(ArrayEntra(5, a)) <> "" Then
      X = X + 1
'      Printer.Print ArrayEntra(0, a) & Space(1) & ArrayEntra(1, a) & Space(c) & ArrayEntra(2, a) & Space(c) & ArrayEntra(3, a) & Space(c) & ArrayEntra(4, a) & Space(c) & ArrayEntra(5, a)
      Printer.Print ArrayEntra(0, a) & ArrayEntra(1, a) & ArrayEntra(2, a) & ArrayEntra(3, a) & ArrayEntra(4, a) & ArrayEntra(5, a)
      linhas_impressas = linhas_impressas + 1

   End If
   
Next
If X = 0 Then
     Printer.Print Space(60) & "SEM MOVIMENTOS"
     linhas_impressas = linhas_impressas + 1

End If

End Sub
Sub imprime_dados_complementares(inscricao As String, Razao_social As String, periodo As String, Situacao As String, data_ger As String, CRC As String, ufCRC As String, Substitutiva As Integer)
Dim b As Integer
Printer.FontBold = True
Printer.FontSize = 9


Printer.Print Space(30) & "DADOS DO EXERCÍCIO ANTERIOR"
Printer.Print

Printer.FontBold = False
Printer.FontSize = 7
 VALOR424 = Space(20 - Len(Trim(VALOR424))) & Trim(VALOR424)
 VALOR400 = Space(20 - Len(Trim(VALOR400))) & Trim(VALOR400)

Printer.Print Space(40) & "Numero de Empregados               : " & VALOR400
Printer.Print Space(40) & "Saldo de caixa mais Bancos em 31/12: " & VALOR424

'Imprime Despesas
Printer.Print
linhas_impressas = linhas_impressas + 5



Printer.FontBold = True
Printer.FontSize = 9
Printer.Print Space(30) & "DETALHAMENTO DAS DESPESAS"
Printer.Print
linhas_impressas = linhas_impressas + 2

Printer.FontBold = False
Printer.FontSize = 7

X = 0

For a = 0 To 14
    If Len(Trim(ArrayDetDesp(1, a))) < 21 Then
      ArrayDetDesp(1, a) = Space(20 - Len(ArrayDetDesp(1, a))) + ArrayDetDesp(1, a)
    End If
     If Trim(ArrayDetDesp(1, a)) <> "" Then
        X = X + 1
        Printer.Print Space(40) & ArrayDetDesp(0, a) & Space(6) & ArrayDetDesp(1, a)
        linhas_impressas = linhas_impressas + 1

     End If
Next
If X = 0 Then
     Printer.Print Space(60) & "SEM MOVIMENTOS"
     linhas_impressas = linhas_impressas + 1

End If

Printer.FontBold = False
Printer.Print String(130, "-")
Printer.Print
linhas_impressas = linhas_impressas + 2

If linhas_impressas >= 65 Then
   Printer.NewPage
   Call imprime_cabecalho_declaracao(inscricao, Razao_social, periodo, Situacao, data_ger, CRC, ufCRC, Substitutiva)

End If

Printer.FontBold = True
Printer.FontSize = 9

Printer.Print Space(30) & "RESUMO DO LIVRO DE REGISTRO DE INVENTARIO"
Printer.Print

Printer.FontBold = False
Printer.FontSize = 7

 VALOR416 = Space(20 - Len(Trim(VALOR416))) & Trim(VALOR416)
 VALOR417 = Space(20 - Len(Trim(VALOR417))) & Trim(VALOR417)
 VALOR418 = Space(20 - Len(Trim(VALOR418))) & Trim(VALOR418)
 VALOR419 = Space(20 - Len(Trim(VALOR419))) & Trim(VALOR419)
 VALOR420 = Space(20 - Len(Trim(VALOR420))) & Trim(VALOR420)
 VALOR421 = Space(20 - Len(Trim(VALOR421))) & Trim(VALOR421)
 VALOR422 = Space(20 - Len(Trim(VALOR422))) & Trim(VALOR422)
 VALOR423 = Space(20 - Len(Trim(VALOR423))) & Trim(VALOR423)
 
Printer.Print Space(40) & "Estoque Inicial"
Printer.Print Space(40) & "Tributadas                    :" & Space(6) & VALOR416
Printer.Print Space(40) & "Não Tributadas                :" & Space(6) & VALOR417
Printer.Print Space(40) & "Outras                        :" & Space(6) & VALOR418
Printer.Print Space(40) & "Total                         :" & Space(6) & VALOR419
Printer.Print
Printer.Print Space(40) & "Estoque Final"
Printer.Print Space(40) & "Tributadas                    :" & Space(6) & VALOR420
Printer.Print Space(40) & "Não Tributadas                :" & Space(6) & VALOR421
Printer.Print Space(40) & "Outras                        :" & Space(6) & VALOR422
Printer.Print Space(40) & "Total                         :" & Space(6) & VALOR423
Printer.Print

End Sub
Sub imprime_cabecalho_declaracao(inscricao As String, Razao_social As String, periodo As String, Situacao As String, data As String, CRC As String, ufCRC As String, Substitutiva As Integer)
  
Printer.FontName = "Courier new"
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print ;
Printer.Print "ESTADO DO MARANHÃO"; Spc(44); "DATA: "; Date
Printer.Print "SECRETARIA DE ESTADO DA FAZENDA"; Spc(31); "Pagina: "; Format$(Printer.Page, "000")
Printer.Print
Printer.Print Space(18) & "DIEF - Declaração de Informações Econômico-Fiscais"
Printer.FontSize = 9
Printer.Print "Referência:"
Printer.FontBold = True
Printer.Print "Inscrição Estadual: " & Format(inscricao, "@@.@@@.@@@-@"); " - " & Trim(Razao_social)
Printer.FontBold = False
Printer.Print "Período           : " & Mid(periodo, 5, 2) & "/" & Mid(periodo, 1, 4);
If status = 2 Then
 Select Case Mid(periodo, 5, 2)
   Case Is = "03"
     Printer.Print Space(4) & "1o. Trimestre";
   Case Is = "06"
     Printer.Print Space(4) & "2o. Trimestre";
   Case Is = "09"
     Printer.Print Space(4) & "3o. Trimestre";
   Case Is = "12"
     Printer.Print Space(4) & "4o. Trimestre";
 End Select
End If

If Substitutiva = 0 Then
   tipo = "Não"
Else
   tipo = "Sim"
End If
Printer.Print Space(15) & "Substitutiva: " & tipo
Printer.Print "Regime de Pagamento:" & condicao
If Situacao = "I" Then
   Printer.Print "Situação          : Declaração Incompleta"
ElseIf Situacao = "F" Then
   Printer.Print "Situação          : Declaração Finalizada"
ElseIf Situacao = "G" Then
   Printer.Print "Situação          : Declaração Gerada"
End If
Printer.Print

linhas_impressas = 10

End Sub
Sub imprime_apuracao_icms()
Dim b As Integer

Printer.FontBold = True
Printer.FontSize = 9

Select Case status
  Case Is = 1
      Printer.Print Space(27) & "APURAÇÃO PARA CONTRIBUINTE NORMAL DO ICMS"
  Case Is = 3, 4, 5, 6, 7
     Printer.Print Space(27) & "APURAÇÃO PARA CONTRIBUINTES ESPECIAIS DO ICMS - EPP"
  Case Is = 2
     Printer.Print Space(27) & "APURAÇÃO PARA CONTRIBUINTES ESPECIAIS DO ICMS - MICROEMPRESA"
End Select

Printer.FontBold = False
Printer.FontSize = 7

Printer.Print
Select Case status
  Case Is = 1, 7
      VALOR200 = Space(20 - Len(Trim(VALOR200))) & Trim(VALOR200)
      Printer.Print Space(30) & "Faturamento do Mes(Receita Bruta)                  : " & Space(5) & VALOR200
  Case Is = 2, 3, 4, 5, 6
     Printer.Print Space(30) & "Faturamento do Mes(Receita Bruta)     Alíquota(%)                      Imposto     " 'EPP c/aliq.<17
     VALOR214 = Space(20 - Len(Trim(VALOR214))) & Trim(VALOR214)
     VALOR215 = Space(6 - Len(Trim(VALOR215))) & Trim(VALOR215)
     VALOR216 = Space(20 - Len(Trim(VALOR216))) & Trim(VALOR216)
     Printer.Print Space(30) & VALOR214 & Space(17) & VALOR215 & Space(15) & VALOR216
End Select

Printer.Print
linhas_impressas = linhas_impressas + 4

If status = 5 Or status = 6 Then Exit Sub

X = 0
For a = 0 To 17
    If a = 4 Then
        Printer.Print
        linhas_impressas = linhas_impressas + 4
    End If
    If a = 9 Then
        Printer.Print
        linhas_impressas = linhas_impressas + 4
    End If
    
  ' If Len(Trim(apurICMS(1, a))) < 201 Then
    If Len(Trim(apurICMS(1, a))) < 20 Then
      apurICMS(1, a) = Space(20 - Len(apurICMS(1, a))) + apurICMS(1, a)
    End If
    
    If Trim(apurICMS(1, a)) <> "" Then
        X = X + 1
        Printer.Print Space(30) & apurICMS(0, a) & Space(5) & apurICMS(1, a)
        linhas_impressas = linhas_impressas + 1

    End If
Next
If X = 0 Then
     Printer.Print Space(60) & "SEM MOVIMENTOS"
     linhas_impressas = linhas_impressas + 1

End If

End Sub

Function crit_val_contabeis(valor_contabil As Double, base As Double, isentas As Double, outras As Double) As Integer
Dim total_contabil As Double

   crit_val_contabeis = True
   total_contabil = base + isentas + outras
   
   If g_condicao = 1 Or g_condicao = 3 Or g_condicao = 4 Or g_condicao = 7 Then
       If valor_contabil >= CCur(total_contabil) Then
       Else
          crit_val_contabeis = False
       End If
       
   Else
       If valor_contabil <> CCur(total_contabil) Then crit_val_contabeis = False
   End If
   

End Function
Sub Impr_declara(inscricao As String, Razao_social As String, periodo As String, Situacao As String, data_ger As String, CRC As String, ufCRC As String, Substitutiva As Integer)
 Dim snpadetalhe As Recordset
 
On Error GoTo err_impressao

linhas_impressas = 0
LIMPA_VARIAVEIS_IMPRESSAO

query = "Select * from tab_cad_contribuintes " _
& "where tcc_nu_inscricao = '" & inscricao & "'"

Set TabTemp = dbgim.OpenRecordset(query, dbOpenSnapshot)


g_ano_inicio_operacao = Val(Mid(TabTemp.Fields("tcc_da_inicio_op") & "", 5, 4))

'1 - Normais, 2 - EPP(Empr.Peq.Porte),3-Microempresa
'select incluido em 17/02/1998 em funcao da inclusao do tipo de contribuinte 4
'para epp com aliquota maior 17 % (Bellyrio)

Query1 = "Select *  From tab_declaracao " _
 & "Where tde_num_inscricao  = '" & inscricao & "'" _
 & " and tde_periodo_inicial = '" & periodo & "'" _
 & " and tde_substitutiva = " & Substitutiva
Set tab_temp1 = dbgim.OpenRecordset(Query1, dbOpenSnapshot)

If Not tab_temp1.EOF Then
   status = tab_temp1!tde_condicao
   Select Case status
     Case 1
      condicao = "Normal"
     Case 2
      condicao = "Microempresa"
     Case 3, 4, 5, 6, 7
      condicao = "Empresa Pequeno Porte"
   End Select
Else
   condicao = "Não Informado."
End If

Call imprime_cabecalho_declaracao(inscricao, Razao_social, periodo, Situacao, data_ger, CRC, ufCRC, Substitutiva)


'verificar se existe registro detalhe gravado
  
query = "select  Count(*) as qtdreg from tab_detalhe_declaracao where tdd_num_inscricao =   '" & inscricao & "'"
query = query & " and [tdd_periodo_inicial] = '" & periodo & "'"
query = query & " and [tdd_substitutiva] = " & Substitutiva
Set snpadetalhe = dbgim.OpenRecordset(query, dbOpenDynaset)
      
If snpadetalhe!qtdReg = 0 Then
   snpadetalhe.Close
   Printer.FontBold = True
   Printer.FontSize = 9
   Printer.Print Space(40) & "DECLARAÇÃO SEM MOVIMENTOS"
   Printer.EndDoc
   Exit Sub
End If

Call cargar_Array_ApurICMS(inscricao, periodo, Substitutiva)
Call cargar_Array_Entrada(inscricao, periodo, Substitutiva)
Call cargar_Array_saidas(inscricao, periodo, Substitutiva)
Call cargar_Array_recolhimento(inscricao, periodo, Substitutiva)
Call cargar_Array_DetDesp(inscricao, periodo, Substitutiva)

Call cargar_Array_faturamento(inscricao, periodo, Substitutiva)


imprime_dem_entradas


Printer.FontBold = False
Printer.Print String(130, "-")
Printer.Print
linhas_impressas = linhas_impressas + 2

'Imprime operações de Saida
   Select Case status
     Case 1, 3, 4, 7
        imprime_dem_saidas
        Printer.FontBold = False
        Printer.Print String(130, "-")
        Printer.Print
        linhas_impressas = linhas_impressas + 2
   End Select

'Imprime APURACAO DO ICMS
If status <> 2 Then
   imprime_apuracao_icms
Else
   imprime_faturamento_micro
End If
Printer.FontBold = False
Printer.Print String(130, "-")
Printer.Print
linhas_impressas = linhas_impressas + 2

'Imprime operações de Recolhimento
Call imprime_recolhimento_periodo(inscricao, Razao_social, periodo, Situacao, data_ger, CRC, ufCRC, Substitutiva)
 

Printer.FontBold = False
Printer.Print String(130, "-")
Printer.Print
linhas_impressas = linhas_impressas + 2

'imprime entrada de ativo fixo ou consumo
'Call imprime_entrada_ativo_fixo(inscricao, Razao_social, periodo, Situacao, data_ger, CRC, ufCRC, Substitutiva)
Dim AnoSis As Double
AnoSis = Val(Year(Date))
  

 If Mid(periodo, 5, 2) = "03" Then
    If g_ano_inicio_operacao < AnoSis Then
'imprime dados complementares
         If linhas_impressas >= 65 Then
            Printer.NewPage
            Call imprime_cabecalho_declaracao(inscricao, Razao_social, periodo, Situacao, data_ger, CRC, ufCRC, Substitutiva)
         End If
        Call imprime_dados_complementares(inscricao, Razao_social, periodo, Situacao, data_ger, CRC, ufCRC, Substitutiva)
     End If
  End If


Printer.EndDoc
MsgBox "Fim de Impressão", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT

err_impressao:
If Err <> 0 Then
     MsgBox "Não foi possivel Imprimir a Declaração, Verificar Impressora", MB_ICONEXCLAMATION, NOMBRE_SISTEMA_DEFAULT
End If
End Sub
Function critica_apuracao_icms_eppmaior(FATURAMENTO As String, grid_ApurICMS() As Variant, tipo_erro As Integer)
Dim total_debito As Double
Dim total_credito As Double
Dim tot_deb_menos_credito As Double
Dim tot_cred_menos_debito As Double
Dim sal_devedor_menos_deducoes As Double
Dim coluna As Integer
Dim linha As Integer
Dim faturamento_wk As Double

critica_apuracao_icms_eppmaior = True
If FATURAMENTO = "" Then
    faturamento_wk = 0#
Else
   faturamento_wk = CDbl(FATURAMENTO)
End If


End Function

Function converte_linha_texto(linha As Integer, pagina As String) As String
converte_linha_texto = ""

If pagina = "E" Then
  Select Case linha
    Case Is = 1
          converte_linha_texto = "Compras no Estado"
    Case Is = 2
          converte_linha_texto = "Transferencias no Estado"
    Case Is = 3
          converte_linha_texto = "Devoluções no Estado"
    Case Is = 4
          converte_linha_texto = "Outras no Estado"
    Case Is = 5
          converte_linha_texto = "Compras de outros Estados"
    Case Is = 6
          converte_linha_texto = "Transf. de outros Estados"
    Case Is = 7
          converte_linha_texto = "Devoluções de outros Estados"
    Case Is = 8
          converte_linha_texto = "Outras de outros Estados"
    Case Is = 9
          converte_linha_texto = "Compras no Exterior"
    Case Is = 10
          converte_linha_texto = "Devoluções do Exterior"
    Case Is = 11
          converte_linha_texto = "Outras entradas do Exterior"
    Case Is = 12
          converte_linha_texto = "Totais de Entradas"
  End Select
Else
   Select Case linha
      Case Is = 1
          converte_linha_texto = "Vendas para o Estado"
      Case Is = 2
          converte_linha_texto = "Transferencias para o Estado"
      Case Is = 3
          converte_linha_texto = "Devoluções para o Estado"
      Case Is = 4
          converte_linha_texto = "Outras para o Estado"
      Case Is = 5
          converte_linha_texto = "Vendas para outros Estados"
      Case Is = 6
          converte_linha_texto = "Transf. para outros Estados"
      Case Is = 7
          converte_linha_texto = "Devoluções para outros Estados"
      Case Is = 8
          converte_linha_texto = "Outras para outros Estados"
      Case Is = 9
          converte_linha_texto = "Vendas para o Exterior"
      Case Is = 10
          converte_linha_texto = "Devoluções para o Exterior"
      Case Is = 11
          converte_linha_texto = "Outras para o Exterior"
      Case Is = 12
          converte_linha_texto = "Totais de Saídas"
     End Select
End If
End Function
Sub imprime_recibo()

    On Error Resume Next
Dim i As Integer

Printer.FontName = "Courier new"
Printer.FontSize = 12
Printer.FontBold = True
Printer.Print "ESTADO DO MARANHÃO"; Spc(38); "DATA: "; Date
Printer.Print "SECRETARIA DE ESTADO DA FAZENDA"; Spc(40); "Pagina: "; Format$(Printer.Page, "000")
Printer.Print
'Printer.Print Space(20) & "Recibo de Guia de Informação do ICMS - GIM"
Printer.Print Space(20) & "DIEF - Declaração de Informações Econômico-Fiscais"

Printer.Print
Printer.Print Space(30) & "Declarações"
Printer.Print
Printer.FontSize = 9
Printer.Print "Inscrição Estadual " & Space(3) & "Período" & Space(3) & "Substitutiva" & Space(4) & "Condição"
Printer.Print
Printer.FontBold = False


    For i = 1 To UBound(gera)
       Printer.Print Space(4); Format(gera(i).inscricao, "@@@.@@@.@@@") & Space(7);
       Printer.Print Mid(gera(i).periodo_inicial, 5, 2) & "/" & Mid(gera(i).periodo_inicial, 1, 4); Space(3);
        If gera(i).tipo_decla = 0 Then
           Printer.Print "Original" & Space(8);
        Else
           Printer.Print "Substitutiva" & gera(i).tipo_decla & Space(3);
        End If
        Select Case gera(i).condicao
               Case Is = 1
                 Printer.Print "Normal"
               Case Is = 3, 4, 5, 6, 7
                 Printer.Print "Empresa Pequeno Porte"
               Case Is = 2
                 Printer.Print "Microempresa"
        End Select
        
    Next


Printer.EndDoc


End Sub





























Function formata_p_impressao(VARIAVEL As String, tam As Integer) As String

  Select Case Len(Trim(VARIAVEL))
       Case Is = 1
          formata_p_impressao = String$((tam - 1), "0") & Trim(VARIAVEL)
      Case Is = 2
         formata_p_impressao = String$((tam - 2), "0") & Trim(VARIAVEL)
       Case Is = 3
          formata_p_impressao = String$((tam - 3), "0") & Trim(VARIAVEL)
      Case Is = 4
         formata_p_impressao = String$((tam - 4), "0") & Trim(VARIAVEL)
       Case Is = 5
          formata_p_impressao = String$((tam - 5), "0") & Trim(VARIAVEL)
      Case Is = 6
         formata_p_impressao = String$((tam - 6), "0") & Trim(VARIAVEL)
       Case Is = 7
          formata_p_impressao = String$((tam - 7), "0") & Trim(VARIAVEL)
      Case Is = 8
         formata_p_impressao = String$((tam - 8), "0") & Trim(VARIAVEL)
      Case Is = 9
         formata_p_impressao = String$((tam - 9), "0") & Trim(VARIAVEL)
      Case Is = 10
         formata_p_impressao = String$((tam - 10), "0") & Trim(VARIAVEL)
     
    End Select
       
End Function

Sub PulaCampo(Teclas As Integer)

   If Teclas = 13 Then
      SendKeys "{tab}"
   End If

End Sub
Sub carComboLogra(Logra As ComboBox)

Logra.Clear
Logra.AddItem "Alameda   " & Space(20) & "ALD"
Logra.AddItem "Avenida   " & Space(20) & "AVE"
Logra.AddItem "Caminho   " & Space(20) & "CAM"
Logra.AddItem "Condominio" & Space(20) & "CON"
Logra.AddItem "Estrada   " & Space(20) & "ETR"
Logra.AddItem "Ilha      " & Space(20) & "ILH"
Logra.AddItem "Ladeira   " & Space(20) & "LAD"
Logra.AddItem "Loteamento" & Space(20) & "LOT"
Logra.AddItem "Lugarejo  " & Space(20) & "LUG"
Logra.AddItem "Morro     " & Space(20) & "MRO"
Logra.AddItem "Praça     " & Space(20) & "PCA"
Logra.AddItem "Ponte     " & Space(20) & "PTE"
Logra.AddItem "Praia     " & Space(20) & "PRA"
Logra.AddItem "Parque    " & Space(20) & "PRQ"
Logra.AddItem "Viaduto   " & Space(20) & "VDO"
Logra.AddItem "Rua       " & Space(20) & "RUA"
Logra.AddItem "Rodovia   " & Space(20) & "RDV"
Logra.AddItem "Travessa  " & Space(20) & "TRV"
Logra.AddItem "Vila      " & Space(20) & "VLA"
Logra.AddItem "Centro    " & Space(20) & "CTR"
Logra.AddItem "Aeroporto " & Space(20) & "AER"
Logra.AddItem "Chacara   " & Space(20) & "CHC"
Logra.AddItem "Favela    " & Space(20) & "FAV"
Logra.AddItem "Galeria   " & Space(20) & "GAL"
Logra.AddItem "Altos     " & Space(20) & "ALT"
Logra.AddItem "Conjunto  " & Space(20) & "CNJ"
Logra.AddItem "Fazenda   " & Space(20) & "FAZ"
Logra.AddItem "Feira     " & Space(20) & "FEI"
Logra.AddItem "Fortaleza " & Space(20) & "FRT"
Logra.AddItem "Jardim    " & Space(20) & "JAR"
Logra.AddItem "Largo     " & Space(20) & "LGO"
Logra.AddItem "Passagem  " & Space(20) & "PAS"
Logra.AddItem "Patio     " & Space(20) & "PAT"
Logra.AddItem "Povoado   " & Space(20) & "POV"
Logra.AddItem "Cais      " & Space(20) & "CAI"
Logra.AddItem "Mercado   " & Space(20) & "MER"
Logra.AddItem "Beco      " & Space(20) & "BEC"





End Sub
Sub carComboUF(uf As ComboBox)

uf.Clear

   'carrega a combo das UFS
    uf.AddItem " AC "
    uf.AddItem " AL "
    uf.AddItem " AM "
    uf.AddItem " AP "
    uf.AddItem " BA "
    uf.AddItem " CE "
    uf.AddItem " DF "
    uf.AddItem " ES "
    uf.AddItem " GO "
    uf.AddItem " MA "
    uf.AddItem " MG "
    uf.AddItem " MS "
    uf.AddItem " MT "
    uf.AddItem " PA "
    uf.AddItem " PB "
    uf.AddItem " PE "
    uf.AddItem " PI "
    uf.AddItem " PR "
    uf.AddItem " RJ "
    uf.AddItem " RN "
    uf.AddItem " RO "
    uf.AddItem " RR "
    uf.AddItem " RS "
    uf.AddItem " SC "
    uf.AddItem " SE "
    uf.AddItem " SP "
    uf.AddItem " TO "
    uf.ListIndex = 9
    
End Sub


Public Sub CentrarFormu(Forma As Form)
    
    Forma.Move (Screen.Width - Forma.Width) / 2, (Screen.Height - Forma.Height) / 2 + 100

End Sub

Function crit_periodo(cperiodo As String) As Integer
 Dim cMES, cAno
   cMES = Mid$(cperiodo, 1, 2)
   cAno = Mid$(cperiodo, 3, 4)
   If cperiodo <> "" Then
       If cMES <> "  " And cAno <> "  " Then
          ' If cDIA >= "01" And cDIA <= "31" Then
               If cMES >= "01" And cMES <= "12" Then
                   If Len(cAno) = 4 Then
'                       If (InStr("01.03.05.07.08.10.12", cMES) <> 0 And cDIA <= "31") Or (InStr("04.06.09.11", cMES) <> 0 And cDIA <= "30") Or ((cMES = "02" And Val(cano) Mod 4 = 0) And cDIA <= "29") Or (cMES = "02" And cDIA <= "28") Then
                           crit_periodo = 0
                           Exit Function
 '                      End If
                   End If
               End If
          ' End If
       End If
       crit_periodo = 1
   End If

End Function

Function crit_tipo_decl(tipo As String) As Integer
   crit_tipo_decl = True

   If UCase(Trim(tipo)) <> "S" And UCase(Trim(tipo)) <> "N" Then
     crit_tipo_decl = False
   End If

End Function

Sub Imprime_contador(cmbuf_crc As String, txt_num_crc As String)

    Dim L As Integer
    Dim n As Integer
    
   strSQLA = "Select * from tab_cad_contadores where tco_num_crc = '" & txt_num_crc & "'"
   strSQLA = strSQLA & "and tco_uf_crc = '" & Trim(cmbuf_crc) & "'"
   Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
    If PESQ_SNAP.EOF And PESQ_SNAP.BOF Then
       Exit Sub
    End If


Printer.FontName = "Courier new"
Printer.FontSize = 12
Printer.FontBold = True

   Printer.Print
   Printer.Print "SECRETARIA DE ESTADO DA FAZENDA "; Spc(42); "DATA: "; Date
   Printer.Print "MARANHÃO"; Spc(56); "Pagina: "; Format$(Printer.Page, "000")
   Printer.Print

Printer.FontSize = 9
   Printer.Print Tab(25); "Cadastro de Contador"
    Printer.Print
   ' Call ImprimeLinha
    Printer.FontBold = True
     Printer.Print "Num. CRC : ";
    Printer.FontBold = False
    Printer.Print txt_num_crc;

    Printer.FontBold = True
    Printer.Print "    UF CRC : ";
    Printer.FontBold = False
    Printer.Print cmbuf_crc;
    Printer.FontBold = True
    Printer.Print " Num. CPF: ";
    Printer.FontBold = False
    Printer.Print Format(PESQ_SNAP.Fields("tco_num_cpf"), "@@@.@@@.@@@-@@")
   


  '  Call ImprimeLinha
   Printer.FontBold = True
   Printer.Print "Nome do Contador  : ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tco_nome_contador")

   Printer.FontBold = True
   Printer.Print "Tipo Logr.: ";
   Printer.FontBold = False
   Printer.Print converte_tipo_logradouro(PESQ_SNAP.Fields("tco_ti_logradouro"));
   Printer.FontBold = True
   Printer.Print " Nome Logr.: ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tco_logradouro");
   Printer.FontBold = True
   Printer.Print " Num.: ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tco_numero");
   Printer.FontBold = True
   Printer.Print " Compl.: ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tco_complemento")

   Printer.FontBold = True
   Printer.Print "Bairro : ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tco_bairro");
   Printer.FontBold = True
   Printer.Print " Municipio : ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tco_municipio");
   
   Dim SNAPMUNIC As Recordset
    query = "SELECT * FROM tab_cod_munic WHERE Cod_Munic = " & PESQ_SNAP.Fields("tco_municipio")
    Set SNAPMUNIC = dbgim.OpenRecordset(query, dbOpenSnapshot)
    If Not SNAPMUNIC.EOF Then
       Printer.Print " - " & Trim(SNAPMUNIC("Nome_Munic"));
    Else
       Printer.Print " - " & "MUNICIPIOS DE OUTROS ESTADOS";
    End If
   SNAPMUNIC.Close
   
   Printer.FontBold = True
   Printer.Print " UF : ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tco_cod_uf");
   Printer.FontBold = True
   Printer.Print " CEP: ";
   Printer.FontBold = False
   Printer.Print Format(PESQ_SNAP.Fields("tco_cep"), "@@.@@@-@@")

  Printer.FontBold = True
  Printer.Print "DDD : ";
  Printer.FontBold = False
  Printer.Print PESQ_SNAP.Fields("tco_ddd_fone");
  Printer.FontBold = True
  Printer.Print " Telefone: ";
  Printer.FontBold = False
  Printer.Print PESQ_SNAP.Fields("tco_num_fone");
  Printer.FontBold = True
  Printer.Print " FAX : ";
  Printer.FontBold = False
  Printer.Print PESQ_SNAP.Fields("tco_ddd_fax"); " - " & PESQ_SNAP.Fields("tco_num_fax")
  Printer.FontBold = True
  Printer.Print "Corr.Eletrônico: ";
  Printer.FontBold = False
  Printer.Print PESQ_SNAP.Fields("tco_corr_eletronico")

   Printer.EndDoc


    Exit Sub


End Sub

Sub IMprime_contrib(msk_inscricao As String)
   
   On Error Resume Next
   
    Dim L As Integer
    Dim n As Integer
    
   strSQLA = "Select * from tab_cad_contribuintes where tcc_nu_inscricao = '" & msk_inscricao & "'"
   Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
    If PESQ_SNAP.EOF And PESQ_SNAP.BOF Then
       Exit Sub
    End If

Printer.FontName = "Courier new"
Printer.FontSize = 12
Printer.FontBold = True

   Printer.Print
   Printer.Print "SECRETARIA DE ESTADO DA FAZENDA "; Spc(42); "DATA: "; Date
   Printer.Print "MARANHÃO"; Spc(56); "Pagina: "; Format$(Printer.Page, "000")
   Printer.Print
   
Printer.FontSize = 9
Printer.Print Tab(25); "Cadastro de Contribuinte"
    Printer.Print
   ' Call ImprimeLinha
    Printer.FontBold = True
   

   Printer.Print "Num. Inscrição: ";
    Printer.FontBold = False
    Printer.Print msk_inscricao;
    Printer.FontBold = True
    Printer.Print "    CGC : ";
    Printer.FontBold = False
    Printer.Print Format(PESQ_SNAP.Fields("tcc_cad_geral_contrib"), "@@.@@@.@@@/@@@@-@@")
  '  Call ImprimeLinha
   Printer.FontBold = True
   Printer.Print "Nome Empresarial : ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tcc_razao_social")

   Printer.FontBold = True
   Printer.Print "Nome Fantasia : ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tcc_no_fantasia")
   Printer.FontBold = True
   Printer.Print "Tipo Logr.: ";
   Printer.FontBold = False
   Printer.Print converte_tipo_logradouro(PESQ_SNAP.Fields("tcc_ti_logradouro"));

   Printer.FontBold = True
   Printer.Print " Nome Logr.: ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tcc_no_logradouro");
   Printer.FontBold = True
   Printer.Print " Num.: ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tcc_nu_logradouro");
   Printer.FontBold = True
   Printer.Print " Compl.: ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tcc_complemento")

   Printer.FontBold = True
   Printer.Print "Bairro : ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tcc_bairro");
   Printer.FontBold = True
   Printer.Print " Municipio : ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tcc_municipio");
   
Dim SNAPMUNIC As Recordset
    query = "SELECT * FROM tab_cod_munic WHERE Cod_Munic = " & PESQ_SNAP.Fields("tcc_municipio")
    Set SNAPMUNIC = dbgim.OpenRecordset(query, dbOpenSnapshot)
    If Not SNAPMUNIC.EOF Then
       Printer.Print " - " & Trim(SNAPMUNIC("Nome_Munic"));
    Else
       Printer.Print " - " & "MUNICIPIOS DE OUTROS ESTADOS";
    End If
   SNAPMUNIC.Close
   
   Printer.FontBold = True
   Printer.Print " UF : ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tcc_cod_uf");
   Printer.FontBold = True
   Printer.Print " CEP: ";
   Printer.FontBold = False
   Printer.Print Format(PESQ_SNAP.Fields("tcc_cep"), "@@.@@@-@@")

  Printer.FontBold = True
  Printer.Print "DDD : ";
  Printer.FontBold = False
  Printer.Print PESQ_SNAP.Fields("tcc_ddd_fone");
  Printer.FontBold = True
  Printer.Print " Telefone: ";
  Printer.FontBold = False
  Printer.Print PESQ_SNAP.Fields("tcc_num_fone");
  Printer.FontBold = True
  Printer.Print " DDD FAX : ";
  Printer.FontBold = False
  Printer.Print PESQ_SNAP.Fields("tcc_ddd_fax");
  Printer.FontBold = True
  Printer.Print " FAX : ";
  Printer.FontBold = False
  Printer.Print PESQ_SNAP.Fields("tcc_num_fax")
  
  Printer.FontBold = True
  Printer.Print "Corr.Eletrônico: ";
  Printer.FontBold = False
  Printer.Print PESQ_SNAP.Fields("tcc_corr_eletronico")

  Printer.FontBold = True
   Printer.Print "Codigo Ativ. Economica: ";
  Printer.Print PESQ_SNAP.Fields("tcc_ativ_economica");
  
  Dim SNAPATIV As Recordset
   query = "Select *  From AtividadeEconomica3 WHERE IT_CO_ATIVIDADE = " & PESQ_SNAP.Fields("tcc_ativ_economica")

    
    Set SNAPATIV = dbgim.OpenRecordset(query, dbOpenSnapshot)
  If Not SNAPATIV.EOF Then
      Printer.Print " - " & SNAPATIV("IT_NO_ATIVIDADE")
   End If
   
   SNAPATIV.Close

   Printer.FontBold = True
   Printer.Print "Localizacao Livro Fiscal: ";
   Printer.FontBold = False
 '  Printer.Print PESQ_SNAP.Fields("tcc_loc_livro_fiscal")
   Select Case PESQ_SNAP.Fields("tcc_loc_livro_fiscal")
     Case Is = 1
   
      Printer.Print "Na Empresa ";
      Case 2
      Printer.Print "No Contador";
   End Select
  Printer.FontBold = True
  Printer.Print "Condição do Contribuinte: ";
 '  Printer.Print PESQ_SNAP.Fields("tcc_reg_declaracao")
 Printer.FontBold = False
Select Case PESQ_SNAP.Fields("tcc_TIPO_contribuinte")
   Case 1
      Printer.Print "Normal"
   Case 3, 4, 5, 6, 7
      Printer.Print "Pequeno Porte"
   Case 2
      Printer.Print "Microempresa"
 End Select

   Printer.FontBold = True
'   Printer.Print " Data de Inicio de Atividades: ";
'   Printer.FontBold = False
'   Printer.Print Format(PESQ_SNAP.Fields("tcc_da_inicio_op"), "@@/@@/@@@@");
'

   Printer.FontBold = True
   Printer.Print "UF CRC: ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tcc_uf_crc_contador");
   Printer.FontBold = True
   Printer.Print " Num. CRC: ";
   Printer.FontBold = False
   Printer.Print PESQ_SNAP.Fields("tcc_num_crc_contador");
   Printer.FontBold = True
 '  Printer.Print " Data Inicio: ";
 '  Printer.FontBold = False
 '  Printer.Print Format(PESQ_SNAP.Fields("tcc_da_inicio_contador"), "@@/@@/@@@@");

 '  Printer.FontBold = True
 '  Printer.Print " Relacao: ";
 '  Printer.FontBold = False
   'Printer.Print PESQ_SNAP.Fields("tcc_relacao_contador");
 '  Select Case PESQ_SNAP.Fields("tcc_relacao_contador")
 '  Case 1
 '     Printer.Print "Empregado"
 '  Case 2
 '       Printer.Print "Autônomo "
 '  End Select
 Printer.EndDoc


 Exit Sub
  
End Sub

Sub grd_selecao_por_nome(datCGC As data, GrdLista As TDBGrid, inscricao As String, achei As Integer)
 
    achei = 0
    strSQLA = "Select * from tab_cad_contadores where tco_nome_contador Like '" & inscricao & "*'"

    Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
   If Not PESQ_SNAP.EOF And Not PESQ_SNAP.BOF Then
      datCGC.RecordSource = strSQLA
      datCGC.DatabaseName = App.Path & "\Dief.MDB"
      datCGC.Refresh
    Else
       achei = 1
    End If
    

End Sub



Sub monta_tela_contador(data_contrib As data)
   Dim i As Integer
  
   
    dec001.txt_num_crc_conta = data_contrib.Recordset("tco_num_crc") & ""
      For i = 0 To dec001.cmb_uf_crc_conta.ListCount - 1
            If data_contrib.Recordset("tco_uf_crc") = Trim(dec001.cmb_uf_crc_conta.List(i)) Then
               dec001.cmb_uf_crc_conta.ListIndex = i 'data_contrib.Recordset("tco_uf_crc") & ""
               Exit For
            End If
      Next i
    
    dec001.msk_cpf = data_contrib.Recordset("tco_num_cpf") & ""
    dec001.txt_nome = data_contrib.Recordset("tco_nome_contador") & ""
    
   For i = 0 To dec001.cmb_tipo_logra_conta.ListCount - 1
         If data_contrib.Recordset("tco_ti_logradouro") = Trim(Right(dec001.cmb_tipo_logra_conta.List(i), 3)) Then
            dec001.cmb_tipo_logra_conta.ListIndex = i 'data_contrib.Recordset("tco_ti_logradouro") & ""
            Exit For
         End If
   Next i
    dec001.txt_log_conta.Text = data_contrib.Recordset("tco_logradouro") & ""
    dec001.txt_num_conta.Text = data_contrib.Recordset("tco_numero") & ""
    dec001.txt_comp_conta.Text = data_contrib.Recordset("tco_complemento") & ""
    dec001.txt_bairro_conta.Text = data_contrib.Recordset("tco_bairro") & ""
    dec001.txt_nome_esp.Text = data_contrib.Recordset("tco_nome_conj") & ""
    dec001.Msk_cpf_esp.Text = data_contrib.Recordset("tco_cpf_conj") & ""
    
    
    For i = 0 To dec001.cmb_mun_conta.ListCount - 1
       If Format$(data_contrib.Recordset("tco_municipio"), "00000") = Trim(Right(dec001.cmb_mun_conta.List(i), 5)) Then
          dec001.cmb_mun_conta.ListIndex = i 'data_contrib.Recordset("tco_municipio") & ""
          Exit For
       End If
    Next i
         
     For i = 0 To dec001.cmb_uf_conta.ListCount - 1
          If data_contrib.Recordset("tco_cod_uf") = Trim(dec001.cmb_uf_conta.List(i)) Then
               dec001.cmb_uf_conta.ListIndex = i 'data_contrib.Recordset("tco_cod_uf") & ""
               Exit For
          End If
    Next i
 
    dec001.MSKcep_conta = data_contrib.Recordset("tco_cep") & ""
    dec001.mskddd_conta = data_contrib.Recordset("tco_ddd_fone") & ""
    dec001.msktel_conta = data_contrib.Recordset("tco_num_fone") & ""
    dec001.mskdddfax_conta = data_contrib.Recordset("tco_ddd_fax") & ""
    dec001.mskfax_conta = data_contrib.Recordset("tco_num_fax") & ""

    dec001.txt_corr_eletr_conta = data_contrib.Recordset("tco_corr_eletronico") & ""

 teve_alteracao_contador = False

End Sub

Sub monta_tela_contribuinte(data_contrib As data)



Dim snativ As Recordset
Dim i As Integer
Dim codativ As String
'strSQLA = "Select * from tab_cad_contribuintes where tcc_nu_inscricao = '" & g_nro_inscricao & "'"

'Set snativ = dbgim.OpenRecordset(strSQLA, dbOpenDynaset)'
' Screen.MousePointer = 11
'If Not snativ.EOF Then

    dec001.msk_inscricao = Format(data_contrib.Recordset("tcc_nu_inscricao"), "@@.@@@.@@@-@")
    dec001.msk_cgc = data_contrib.Recordset("tcc_cad_geral_contrib") & ""

    dec001.txt_razao = data_contrib.Recordset("tcc_razao_social") & ""
    dec001.txt_fantasia = data_contrib.Recordset("tcc_no_fantasia") & ""
      
   For i = 0 To dec001.cmb_tipo_logradouro.ListCount - 1
      If data_contrib.Recordset("tcc_ti_logradouro") = Trim(Right(dec001.cmb_tipo_logradouro.List(i), 3)) Then
         dec001.cmb_tipo_logradouro.ListIndex = i 'data_contrib.Recordset("tcc_ti_logradouro") & ""
         Exit For
      End If
   Next i

    dec001.txt_log = data_contrib.Recordset("tcc_no_logradouro") & ""
    dec001.txt_num = data_contrib.Recordset("tcc_nu_logradouro") & ""
    dec001.txt_comp = data_contrib.Recordset("tcc_complemento") & ""

    dec001.txt_bairro = data_contrib.Recordset("tcc_bairro") & ""
      For i = 0 To dec001.cmb_MUN.ListCount - 1
            If Format$(data_contrib.Recordset("tcc_municipio"), "00000") = Format$(Right(dec001.cmb_MUN.List(i), 5), "00000") Then
                  dec001.cmb_MUN.ListIndex = i 'data_contrib.Recordset("tcc_municipio") & ""
                  Exit For
            End If
      Next i
    
      For i = 0 To dec001.CMB_UF.ListCount - 1
            If data_contrib.Recordset("tcc_cod_uf") = Trim(dec001.CMB_UF.List(i)) Then
                  dec001.CMB_UF.ListIndex = i 'data_contrib.Recordset("tcc_cod_uf") & ""
                  Exit For
            End If
      Next i

    dec001.MSKcep = data_contrib.Recordset("tcc_cep") & ""
    dec001.mskddd = data_contrib.Recordset("tcc_ddd_fone") & ""
    dec001.msktel = data_contrib.Recordset("tcc_num_fone") & ""
    dec001.mskddd_fax = data_contrib.Recordset("tcc_ddd_fax") & ""
    dec001.mskfax = data_contrib.Recordset("tcc_num_fax") & ""

    dec001.txt_corr_eletronico = data_contrib.Recordset("tcc_corr_eletronico") & ""
          
     codativ = data_contrib.Recordset("tcc_ativ_economica")
     strSQLA = "Select * from AtividadeEconomica3 where Mid$(trim(cstr(IT_CO_ATIVIDADE)),1,5) = '" & codativ & "'"

   '  Set snativ = dbgim.OpenRecordset(strSQLA, dbOpenDynaset) '
     Set snativ = dbgim.CreateSnapshot(strSQLA) '
'     Set restauraconta = dbgim.CreateDynaset(query)
     dec001.Label1.Text = snativ!IT_CO_ATIVIDADE & "  " & snativ!IT_NO_ATIVIDADE
       
  
      For i = 0 To dec001.cmbCrc.ListCount - 1
         If data_contrib.Recordset("tcc_num_crc_contador") = Trim(Right(dec001.cmbCrc.List(i), 8)) Then
            dec001.cmbCrc.ListIndex = i 'data_contrib.Recordset("tcc_uf_crc_contador") & ""
            Exit For
         End If
      Next i

dec001.Text2 = data_contrib.Recordset("tcc_uf_crc_contador") & ""
    dec001.txt_num_crc = data_contrib.Recordset("tcc_num_crc_contador") & ""
    dec001.mskdtinicio = data_contrib.Recordset("tcc_da_inicio_contador") & ""
  dec001.Mask_dt_inicio_op = data_contrib.Recordset("tcc_da_inicio_op") & ""
   
   For i = 0 To dec001.cmb_relacao.ListCount - 1
'       If data_contrib.Recordset("tcc_relacao_contador") = Mid$(Trim(dec001.cmb_relacao.List(i)), 1, 1) Then
      If data_contrib.Recordset("tcc_relacao_contador") = i + 1 Then
         dec001.cmb_relacao.ListIndex = i 'data_contrib.Recordset("tcc_relacao_contador") & ""
         Exit For
      End If
   Next i

      For i = 0 To dec001.cmb_localizacao.ListCount - 1
      '   If data_contrib.Recordset("tcc_loc_livro_fiscal") = Mid$(Trim(dec001.cmb_localizacao.List(i)), 1, 1) Then
         If data_contrib.Recordset("tcc_loc_livro_fiscal") = i + 1 Then
   '         dec001.cmb_localizacao.ListIndex = i 'data_contrib.Recordset("tcc_loc_livro_fiscal") & ""
            dec001.cmb_localizacao.ListIndex = i 'data_contrib.Recordset("tcc_loc_livro_fiscal") & ""
            Exit For
         End If
   Next i
 
   For i = 0 To dec001.cmb_tipo.ListCount - 1
'         If data_contrib.Recordset("tcc_tipo_contribuinte") = Mid$(Trim(dec001.cmb_tipo.List(i)), 1, 1) Then
        If data_contrib.Recordset("tcc_tipo_contribuinte") = i + 1 Then
               dec001.cmb_tipo.ListIndex = i 'data_contrib.Recordset("tcc_tipo_contribuinte") & ""
               Exit For
         End If
   Next i

 teve_alteracao = False
 
End Sub
 Function Rotcpf(Campo As String)

    
    If Mod11(Campo, 10, 10) Then
        If Mod11(Campo, 11, 11) Then
            Rotcpf = True
            Exit Function
        End If
    End If

    Rotcpf = False


End Function
Function Mod11(Campo As String, nPOSDV As Double, nPESOMAX As Double)
 Dim nDV, nPESO, nIMOD, nRESTO

   nDV = 0
   nPESO = 1

   For nIMOD = nPOSDV To 1 Step -1
       nDV = nDV + (nPESO * (Val(Mid$(Campo, nIMOD, 1))))
       nPESO = nPESO + 1
       If nPESO > nPESOMAX Then
          nPESO = 2
       End If
   Next nIMOD

   nRESTO = nDV Mod 11
   If nRESTO = 0 Then
       Mod11 = True
       Exit Function
   Else
      If nRESTO = 1 And Val(Mid$(Campo, nPOSDV, 1)) = 0 Then
          Mod11 = True
          Exit Function
      End If
   End If

   Mod11 = False

End Function

Function Crit_relacao(relacao As String) As Integer
Crit_relacao = True

If Trim(relacao) <> "1" And Trim(relacao) <> "2" Then
  Crit_relacao = False
End If

End Function

Function CritUF(cUF As String) As String
    
Dim i As Integer

    cUF = UCase(cUF)
    CritUF = ""
    
    If Len(Trim$(cUF)) = 0 Then
        CritUF = "Unidade da Federação ausente."
    ElseIf Len(Trim$(cUF)) = 1 Then
        CritUF = "Unidade da Federação inválida."
    ElseIf Len(Trim$(cUF)) = 2 Then
        
        For i = 1 To UBound(TabUf)
            If Trim$(UCase(cUF)) = TabUf(i) Then
               Exit For
            End If
        Next i
        
        If i = 28 Then
            CritUF = "Unidade da Federação inválida."
        End If

    End If

End Function

Function CritATIVID1(ativid As String) As Integer
   
   Dim ativ As String
   Dim snativ As Recordset
   
   CritATIVID1 = 0
   
   If Len(Mid(ativid, 1, 5)) = 0 Or ativid = "  " Then
      CritATIVID1 = 1
      Exit Function
   ElseIf Len(Mid(ativid, 1, 5)) <> 5 Then
      CritATIVID1 = 2
      Exit Function
   End If
   
   ativid = Mid(Trim(ativid), 1, 5)
   ativ = "SELECT * FROM AtividadeEconomica3 WHERE it_co_atividade =" & ativid
   Set snativ = dbgim.CreateSnapshot(ativ) 'criação do snapshot
   
   If snativ.RecordCount = 0 Then
      CritATIVID1 = 3
   End If
End Function

Function critica_tipo_log(tipo_log As String) As Integer

critica_tipo_log = True
Select Case UCase(Trim(tipo_log))
  Case Is = "AL", "AV", "CAM", "CON", "ETR"
    critica_tipo_log = True
   Case Is = "ILH", "LAD", "LOT", "LUG"
       critica_tipo_log = True
   Case Is = "MRO", "PCA", "PRA", "PRQ"
       critica_tipo_log = True
   Case Is = "R", "RDV", "TR", "VL"
       critica_tipo_log = True
   Case Else
     critica_tipo_log = True
End Select

End Function


Function CritString(ctexto As String) As String

Dim nTexto As String
Dim i As Integer

  nTexto = Trim(ctexto)
    
  For i = 1 To Len(nTexto)
      If Mid$(nTexto, i, 1) = " " Then
         Mid$(nTexto, i, 1) = "0"
      End If
  Next i
  
  CritString = ""
  If Len(Trim(ctexto)) < 1 And Len(Trim(ctexto)) <> 0 Then
    CritString = " com tamanho inválido."
  ElseIf Left(ctexto, 1) = " " Then
    CritString = " começando com espaço em branco."
  'ElseIf IsNumeric(nTexto) = True Then
  '  CritString = " contém apenas números."
  ElseIf Len(Trim(ctexto)) = 0 Then
    CritString = " ausente."
  End If

End Function
    

Sub abre_base_gim()
    Dim sMens As String
    On Error Resume Next
    
    Set dbgim = OpenDatabase(App.Path & "\DIEF.MDB", False, False)
    If Err <> 0 Then
      Err = 0
      RepairDatabase (App.Path & "\DIEF.MDB")
      If Err = 0 Then
        CompactaBase_gim
      End If
      Err = 0
      Set dbgim = OpenDatabase(App.Path & "\DIEF.MDB", False, False)
      If Err <> 0 Then
        sMens = "Não foi possível abrir o banco, "
        sMens = sMens & "reinstale o sistema e recupere os"
        sMens = sMens & " dados da sua última cópia de segurança."
        MsgBox sMens, MB_ICONEXCLAMATION, NOMBRE_SISTEMA_DEFAULT
        End
      End If
    End If

End Sub


Sub CompactaBase_gim()
'On Error Resume Next

'Apaga arquivo "old" se existir
Kill App.Path & "\dief.OLD"

'Renomeia o banco para "old"
Name App.Path & "\dief.MDB" As App.Path & "\Dief.OLD"

'Compacta o arquivo "old" gerando a base definitiva
CompactDatabase App.Path & "\Dief.OLD", App.Path & "\Dief.MDB"
End Sub

 Function CRIT_CRC(CRC As String) As Integer

  CRIT_CRC = 0
  'Criticando se a primeira posicao e branco.
 If Left(CRC, 1) = " " Then
    CRIT_CRC = 1
  'Criticar se campo contem mais de um branco entre os caracteres.
 ElseIf InStr(Trim$(CRC), "  ") <> "0" Then
   CRIT_CRC = 2
 ElseIf Len(CRC) < 3 Then
    CRIT_CRC = 3
 End If



End Function

Function CRIT_Data(cdata As String)
 Dim cDIA, cMES, cAno
   cDIA = Mid$(cdata, 1, 2)
   cMES = Mid$(cdata, 3, 2)
   cAno = Mid$(cdata, 5, 4)
   If cdata <> "" Then
       If cDIA <> "  " And cMES <> "  " And cAno <> "  " Then
           If cDIA >= "01" And cDIA <= "31" Then
               If cMES >= "01" And cMES <= "12" Then
                   If Len(cAno) = 4 Then
                       If (InStr("01.03.05.07.08.10.12", cMES) <> 0 And cDIA <= "31") Or (InStr("04.06.09.11", cMES) <> 0 And cDIA <= "30") Or ((cMES = "02" And Val(cAno) Mod 4 = 0) And cDIA <= "29") Or (cMES = "02" And cDIA <= "28") Then
                           CRIT_Data = 0
                           Exit Function
                       End If
                   End If
               End If
           End If
       End If
       CRIT_Data = 1
   End If
End Function

Function EjecutarQuery(query As String) As Integer

    EjecutarQuery = False

On Error GoTo err_label:
    dbgim.Execute query
    EjecutarQuery = True
    Exit Function

err_label:
 
     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOMBRE_SISTEMA_DEFAULT
     MsgBox "Erro na Gravação. Dados não foram gravados.", MB_ICONEXCLAMATION, NOMBRE_SISTEMA_DEFAULT
End Function

Function GenDV(rif As String, xbase As Integer) As Integer
    Dim i As Integer
    Dim tem As Integer
    Dim rif_tmp As String
    
    tem = 0
    rif_tmp = Format(rif, "000000000000")
    For i = 1 To Len(rif_tmp)
        tem = tem + Val(Mid(rif_tmp, i, 1)) * (((Len(rif_tmp) - i) Mod (xbase - 1)) + 2)
    Next
    tem = tem Mod 11
    If tem > 1 Then
        tem = 11 - tem
    Else
        tem = 0
    End If
    GenDV = tem
End Function

Sub Grd_Selecao(datCGC As data, GrdLista As TDBGrid, achei As Integer)

    achei = 0
    
    strSQLA = "Select * from tab_cad_contribuintes "
    Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
    If Not PESQ_SNAP.EOF And Not PESQ_SNAP.BOF Then
        datCGC.RecordSource = strSQLA
        datCGC.DatabaseName = App.Path & "\Dief.MDB"
        datCGC.Refresh
    Else
        achei = 2
    End If
   
End Sub
 Sub Grd_Selecao_contador(datCGC As data, GrdLista As TDBGrid, achei As Integer)
   achei = 0
     
    strSQLA = "Select * from tab_cad_contadores "
   
    Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
    If Not PESQ_SNAP.EOF And Not PESQ_SNAP.BOF Then
        datCGC.RecordSource = strSQLA
        datCGC.DatabaseName = App.Path & "\Dief.MDB"
        datCGC.Refresh
    Else
        achei = 2
    End If
    
    
    


End Sub
Sub grd_selecao_por_fantasia(datCGC As data, GrdLista As TDBGrid, inscricao As String, achei As Integer)
  achei = 0

 

  strSQLA = "Select * from tab_cad_contribuintes where tcc_no_fantasia Like '" & inscricao & "*'"
  
  Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
  If Not PESQ_SNAP.EOF And Not PESQ_SNAP.BOF Then
    datCGC.RecordSource = strSQLA
    datCGC.DatabaseName = App.Path & "\Dief.MDB"
    datCGC.Refresh
  Else
    achei = 1
  End If
  
  


End Sub

Sub grd_selecao_por_inscricao(datCGC As data, GrdLista As TDBGrid, inscricao As String, achei As Integer)

 
    achei = 0
    strSQLA = "Select * from tab_cad_contribuintes where tcc_nu_inscricao = '" & inscricao & "'"
    Screen.MousePointer = 13

Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
If Not PESQ_SNAP.EOF And Not PESQ_SNAP.BOF Then
    datCGC.RecordSource = strSQLA
    datCGC.DatabaseName = App.Path & "\Dief.MDB"
    datCGC.Refresh
Else
   achei = 1
End If
    
PESQ_SNAP.Close

End Sub

Sub grd_selecao_por_razao(datCGC As data, GrdLista As TDBGrid, inscricao As String, achei As Integer)
    achei = 0
    

    strSQLA = "Select * from tab_cad_contribuintes where tcc_razao_social LIKE '" & inscricao & "*'"
    Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
    If Not PESQ_SNAP.EOF And Not PESQ_SNAP.BOF Then
       datCGC.RecordSource = strSQLA
       datCGC.DatabaseName = App.Path & "\Dief.MDB"
       datCGC.Refresh
    Else
       achei = 1
       
    End If

End Sub
Sub grd_selecao_por_razao_e_fantasia(datCGC As data, GrdLista As TDBGrid, inscricao As String, fantasia As String, achei As Integer)
   achei = 0
    

    strSQLA = "Select * from tab_cad_contribuintes where tcc_razao_social LIKE '" & inscricao & "*' and  tcc_no_fantasia Like '" & fantasia & "*'"
    Set PESQ_SNAP = dbgim.OpenRecordset(strSQLA, dbOpenSnapshot)
    If Not PESQ_SNAP.EOF And Not PESQ_SNAP.BOF Then
       datCGC.RecordSource = strSQLA
       datCGC.DatabaseName = App.Path & "\Dief.MDB"
       datCGC.Refresh
    Else
       achei = 1
       
    End If


End Sub

Sub init_system()
    Dim Res As Integer
    Dim WinDate_Str As String

    'Lectura o Ajuste de parámetros de conexion, desde windows registry
    Dim EstaEjecucion As Long
   ' EstaEjecucion = GetSetting(APPNAME = "TAXSOLUTIONS", section = "GENERAL", Key = "CorridaNumero", Default = PRIMERACORRIDA)
   ' If EstaEjecucion = 0 Then
   '     SaveSetting "TAXSOLUTIONS", "GENERAL", "SEP_DEC", DEC_SEP_DEFAULT
   '     SaveSetting "TAXSOLUTIONS", "GENERAL", "BASELOCAL", DBFILE_DEFAULT
   '     SaveSetting "TAXSOLUTIONS", "GENERAL", "HOSTNAME", HOSTNAME_DEFAULT
    ''  End If
  '  NOMBRE_SISTEMA = GetSetting(APPNAME:="TAXSOLUTIONS", section:="GENERAL", Key:="SystemName", Default:=NOMBRE_SISTEMA_DEFAULT)
   ' DEC_SEP = GetSetting(APPNAME = "TAXSOLUTIONS", section = "GENERAL", Key = "SEP_DEC", Default = DEC_SEP_DEFAULT)
   ' DBFILE = GetSetting(APPNAME = "TAXSOLUTIONS", section = "GENERAL", Key = "BASELOCAL", Default = DBFILE_DEFAULT)
   ' HOSTNAME = GetSetting(APPNAME = "TAXSOLUTIONS", section = "GENERAL", Key = "HOSTNAME", Default = HOSTNAME_DEFAULT)
    
    'Guarda el número de ejecución de la aplicación en ésta máquina
   ' SaveSetting "TAXSOLUTIONS", "GENERAL", "CorridaNumero", EstaEjecucion + 1
    

End Sub


Function teclas_especiais(tecla As Integer) As Integer

teclas_especiais = True

If tecla >= 33 And tecla <= 47 Then
   teclas_especiais = False
ElseIf tecla >= 58 And tecla <= 64 Then
    teclas_especiais = False
ElseIf tecla >= 91 And tecla <= 96 Then
   teclas_especiais = False
ElseIf tecla >= 123 And tecla <= 127 Then
   teclas_especiais = False
End If

End Function

Function tirapictu(cString As String, cChar As String) As String
     Dim nPoint As Integer
     
     If Len(cChar) = 1 Then
          nPoint = InStr(cString, cChar)
          While nPoint
               cString = Left(cString, nPoint - 1) & Mid$(cString, nPoint + 1, Len(cString))
               nPoint = InStr(nPoint, cString, cChar)
          Wend
     End If
     tirapictu = cString
End Function

Function Validar_Cgc(mycgc As String) As Integer

  Dim Title
  Dim numero, NumAux, Response
  Title = NOMBRE_SISTEMA_DEFAULT
  If mycgc = "00000000000000" Then
    Validar_Cgc = False
    MsgBox "CGC inválido.", MB_ICONEXCLAMATION, Title
    Exit Function
  End If
  If Len(mycgc) <> 14 Then
    Response = MsgBox("CGC inválido.", 48, Title)
    Validar_Cgc = False
    Exit Function
  Else
    numero = 5 * Val(Left(mycgc, 1))
    numero = numero + 4 * Val(Mid(mycgc, 2, 1))
    numero = numero + 3 * Val(Mid(mycgc, 3, 1))
    numero = numero + 2 * Val(Mid(mycgc, 4, 1))
    numero = numero + 9 * Val(Mid(mycgc, 5, 1))
    numero = numero + 8 * Val(Mid(mycgc, 6, 1))
    numero = numero + 7 * Val(Mid(mycgc, 7, 1))
    numero = numero + 6 * Val(Mid(mycgc, 8, 1))
    numero = numero + 5 * Val(Mid(mycgc, 9, 1))
    numero = numero + 4 * Val(Mid(mycgc, 10, 1))
    numero = numero + 3 * Val(Mid(mycgc, 11, 1))
    numero = numero + 2 * Val(Mid(mycgc, 12, 1))
    numero = numero Mod 11
    numero = IIf(numero = 0 Or numero = 1, 0, 11 - numero)
    If numero <> Val(Mid(mycgc, 13, 1)) Then
      Response = MsgBox("CGC inválido.", 48, Title)
      Validar_Cgc = False
      Exit Function
    Else
      NumAux = numero
      numero = 6 * Val(Left(mycgc, 1))
      numero = numero + 5 * Val(Mid(mycgc, 2, 1))
      numero = numero + 4 * Val(Mid(mycgc, 3, 1))
      numero = numero + 3 * Val(Mid(mycgc, 4, 1))
      numero = numero + 2 * Val(Mid(mycgc, 5, 1))
      numero = numero + 9 * Val(Mid(mycgc, 6, 1))
      numero = numero + 8 * Val(Mid(mycgc, 7, 1))
      numero = numero + 7 * Val(Mid(mycgc, 8, 1))
      numero = numero + 6 * Val(Mid(mycgc, 9, 1))
      numero = numero + 5 * Val(Mid(mycgc, 10, 1))
      numero = numero + 4 * Val(Mid(mycgc, 11, 1))
      numero = numero + 3 * Val(Mid(mycgc, 12, 1))
      numero = numero + 2 * NumAux
      numero = numero Mod 11
      numero = IIf(numero = 0 Or numero = 1, 0, 11 - numero)
      If numero <> Val(Right(mycgc, 1)) Then
         Response = MsgBox("CGC inválido.", 48, Title)
         Validar_Cgc = False
         Exit Function
      End If
    End If
  End If

  Validar_Cgc = True

End Function
Function Validar_Cgc_c(mycgc As String, msg As String) As Integer
  Dim Title
  Dim numero, NumAux, Response
  Title = "Atenção!"
  If mycgc = "00000000000000" Then
    Validar_Cgc_c = False
    msg = "CGC inválido."
    Exit Function
  End If



 If Len(mycgc) <> 14 Then
    msg = "CGC inválido."
    Validar_Cgc_c = False
    Exit Function
  Else
    numero = 5 * Val(Left(mycgc, 1))
    numero = numero + 4 * Val(Mid(mycgc, 2, 1))
    numero = numero + 3 * Val(Mid(mycgc, 3, 1))
    numero = numero + 2 * Val(Mid(mycgc, 4, 1))
    numero = numero + 9 * Val(Mid(mycgc, 5, 1))
    numero = numero + 8 * Val(Mid(mycgc, 6, 1))
    numero = numero + 7 * Val(Mid(mycgc, 7, 1))
    numero = numero + 6 * Val(Mid(mycgc, 8, 1))
    numero = numero + 5 * Val(Mid(mycgc, 9, 1))
    numero = numero + 4 * Val(Mid(mycgc, 10, 1))
    numero = numero + 3 * Val(Mid(mycgc, 11, 1))
    numero = numero + 2 * Val(Mid(mycgc, 12, 1))
    numero = numero Mod 11
    numero = IIf(numero = 0 Or numero = 1, 0, 11 - numero)
    If numero <> Val(Mid(mycgc, 13, 1)) Then
      msg = "CGC inválido."
      Validar_Cgc_c = False
      Exit Function
    Else
      NumAux = numero
      numero = 6 * Val(Left(mycgc, 1))
      numero = numero + 5 * Val(Mid(mycgc, 2, 1))
      numero = numero + 4 * Val(Mid(mycgc, 3, 1))
      numero = numero + 3 * Val(Mid(mycgc, 4, 1))
      numero = numero + 2 * Val(Mid(mycgc, 5, 1))
      numero = numero + 9 * Val(Mid(mycgc, 6, 1))
      numero = numero + 8 * Val(Mid(mycgc, 7, 1))
      numero = numero + 7 * Val(Mid(mycgc, 8, 1))
      numero = numero + 6 * Val(Mid(mycgc, 9, 1))
      numero = numero + 5 * Val(Mid(mycgc, 10, 1))
      numero = numero + 4 * Val(Mid(mycgc, 11, 1))
      numero = numero + 3 * Val(Mid(mycgc, 12, 1))
      numero = numero + 2 * NumAux
      numero = numero Mod 11
      numero = IIf(numero = 0 Or numero = 1, 0, 11 - numero)
      If numero <> Val(Right(mycgc, 1)) Then
         msg = "CGC inválido."
         Validar_Cgc_c = False
         Exit Function
      End If
    End If
  End If

  Validar_Cgc_c = True

End Function
