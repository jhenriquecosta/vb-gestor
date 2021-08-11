VERSION 5.00
Begin VB.Form ImpEnvio 
   BorderStyle     =   0  'None
   Caption         =   "Importar Declaracao"
   ClientHeight    =   1740
   ClientLeft      =   1815
   ClientTop       =   2415
   ClientWidth     =   4365
   ClipControls    =   0   'False
   FillColor       =   &H00C0C0C0&
   Icon            =   "IMPENVIO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1740
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Informe a unidade de leitura :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1065
      Left            =   150
      TabIndex        =   1
      Top             =   570
      Width           =   2895
      Begin VB.PictureBox optDrv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   1650
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   3
         Top             =   390
         Width           =   495
      End
      Begin VB.PictureBox optDrv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   435
         TabIndex        =   2
         Top             =   390
         Width           =   495
      End
   End
   Begin VB.PictureBox panGerEnvio 
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   1890
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   240
      ScaleHeight     =   225
      ScaleWidth      =   3555
      TabIndex        =   6
      Top             =   240
      Width           =   3615
      Begin VB.Image Image2 
         Height          =   240
         Left            =   225
         Picture         =   "IMPENVIO.frx":0442
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.PictureBox cmdSair 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   1005
      TabIndex        =   8
      Top             =   1200
      Width           =   1065
   End
   Begin VB.PictureBox cmdgerar 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   1005
      TabIndex        =   7
      Top             =   720
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Height          =   1695
      Left            =   30
      Top             =   30
      Width           =   4305
   End
   Begin VB.Label Label2 
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1830
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inscrição Estadual: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   660
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "ImpEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim reg As String
Dim DrvSelec  As String
Dim diretorio As String

Dim registro As String
Dim soma_hasharq As Long
Dim snap_detalhe As Recordset
Dim nFile As Integer
Dim num_inscricao As String
Dim periodo_inicial As String
Dim tipo_decla As Integer
Dim data_geracao As String
Dim qtd_reg As Long
Dim msgerro As String

Sub atualiza_data_geracao()
    Dim datager As String
    Dim i As Integer
    
    On Error GoTo err_label:
    
    datager = Format$(Date, "yyyymmdd")
    'atualiza tab_cad_contadores
    query = "update tab_cad_contadores set tco_da_geracao = '" & datager & "' ,"
    query = query & "tco_Estado = 'G' where tco_estado = 'A'"
    
'    dbgim.Execute query ', dbSQLPassThrough
     If Not EjecutarQuery(query) Then
         Screen.MousePointer = 0
     End If
    'atualiza tab_cad_contribuintes
    query = "update tab_cad_contribuintes set tcc_da_geracao = '" & datager & "' , "
    query = query & " tcc_Estado = 'G' where tcc_estado = 'A'"
        
    'dbgim.Execute query ', dbSQLPassThrough
    If Not EjecutarQuery(query) Then
       Screen.MousePointer = 0
     End If
    
    'atualiza tab_declaracao
    For i = 1 To UBound(gera)
        
        num_inscricao = gera(i).inscricao
        periodo_inicial = gera(i).periodo_inicial
        tipo_decla = gera(i).tipo_decla
        data_geracao = gera(i).data_geracao
        
        
        query = "update tab_declaracao set tde_da_geracao = '" & datager & "' , "
        query = query & " tde_Estado = 'G' , tde_versao_gerador = '" & Versao & "'"
        query = query & " Where tde_num_inscricao  = '" & gera(i).inscricao & "'"
        query = query & " and tde_periodo_inicial = '" & gera(i).periodo_inicial & "'"
        query = query & " and tde_substitutiva = " & gera(i).tipo_decla
        
      '  dbgim.Execute query
        If Not EjecutarQuery(query) Then
           Screen.MousePointer = 0
         End If
    Next
    
    
    Exit Sub

err_label:
 
 '    bd.Rollback

     MsgBox "Error Message :" & Str(Err), MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT

    
End Sub


Function critica_tipo_cinco(detalhe As String) As Integer
Dim snap_declaracao As Recordset
critica_tipo_cinco = True

    If Not IsNumeric(Mid(detalhe, 2, 9)) Then
     MsgBox "Erro no Registro Tipo tres. Arquivo Não esta Integro"
      critica_tipo_cinco = False
    End If
    If Not IsNumeric(Mid(detalhe, 11, 14)) Then
     MsgBox "Erro no Registro Tipo tres. Arquivo Não esta Integro"
      critica_tipo_cinco = False
    End If

    query = "Select *  From tab_cad_contribuintes "
    query = query & "Where tcc_nu_inscricao  = '" & Mid(detalhe, 2, 9) & "'"
    Set snap_declaracao = dbgim.OpenRecordset(query, dbOpenSnapshot)
    If snap_declaracao.EOF Then
        query = "insert into tab_cad_contribuintes"
        query = query & "(tcc_nu_inscricao, tcc_cad_geral_contrib, "
        query = query & " tcc_razao_social, tcc_no_fantasia ,"
        query = query & " tcc_ti_logradouro, tcc_no_logradouro, "
        query = query & " tcc_nu_logradouro , tcc_complemento, "
        query = query & " tcc_bairro ,tcc_municipio , "
        query = query & " tcc_cod_uf, tcc_cep , "
        query = query & " tcc_ddd_fone , tcc_num_fone,  "
        query = query & " tcc_ddd_fax , tcc_num_fax ,"
        query = query & " tcc_corr_eletronico, tcc_ativ_economica , "
        query = query & " tcc_uf_crc_contador , tcc_num_crc_contador, "
        query = query & " tcc_da_inicio_contador ,tcc_relacao_contador , "
        query = query & " tcc_loc_livro_fiscal, tcc_tipo_contribuinte , "
        query = query & "tcc_da_inicio_op ) "
        query = query & "values('" & Mid(detalhe, 2, 9) & "','" & Mid(detalhe, 11, 14) & "','"
        query = query & Mid(detalhe, 25, 59) & "','" & Mid(detalhe, 84, 40) & "','"
        query = query & Trim(Mid(detalhe, 124, 3)) & "','" & Trim(Mid(detalhe, 127, 30)) & "','"
        query = query & Mid(detalhe, 157, 5) & "','" & Mid(detalhe, 162, 20) & "','"
        query = query & Mid(detalhe, 182, 20) & "','" & Mid(detalhe, 202, 5) & "','"
        query = query & Mid(detalhe, 207, 2) & "','" & Mid(detalhe, 209, 8) & "','"
        query = query & Mid(detalhe, 217, 5) & "','" & Mid(detalhe, 222, 10) & "','"
        query = query & Mid(detalhe, 232, 5) & "','" & Mid(detalhe, 237, 10) & "','"
        query = query & Mid(detalhe, 247, 30) & "','" & Mid(detalhe, 277, 5) & "','"
        query = query & Mid(detalhe, 282, 2) & "','" & Mid(detalhe, 284, 7) & "','"
        query = query & Mid(detalhe, 291, 8) & "','" & Mid(detalhe, 299, 1) & "','"
        query = query & Mid(detalhe, 300, 1) & "','" & Mid(detalhe, 301, 1) & "','"
        query = query & Mid(detalhe, 302, 8) & "')"
        If Not EjecutarQuery(query) Then
        End If
    End If
    snap_declaracao.Close
    

End Function

Function critica_tipo_quatro(detalhe As String) As Integer
critica_tipo_quatro = True
'REGISTRO TIPO 4
'Tipo de registro       N1  "4"
'ITEM                   N3 2-4  TAB_DETALHE_DECLARACAO!TDD_ITEM
'VALOR                  N15 5-19 TAB_DETALHE_DECLARACAO!TDD_VALOR
'HASH                   N4  CALCULADO
Dim valor As String * 15
If Not IsNumeric(Mid(detalhe, 2, 3)) Then
 MsgBox "Erro no Registro Tipo quatro."
  critica_tipo_quatro = False
End If
If Not IsNumeric(Mid(detalhe, 5, 15)) Then
 MsgBox "Erro no Registro Tipo quatro."
  critica_tipo_quatro = False
End If
'TAB_DETALHE_DECALRACAO
Dim reg4 As String
query = "Select *  From TAB_DETALHE_DECLARACAO "
query = query & "Where tdd_num_inscricao  = '" & num_inscricao & "'"
query = query & " and tdd_periodo_inicial = '" & periodo_inicial & "'"
query = query & " and tdd_substitutiva = " & tipo_decla
query = query & " and tdd_item = " & Val(Mid(detalhe, 2, 3))
Set snap_detalhe = dbgim.OpenRecordset(query, dbOpenSnapshot)
If snap_detalhe.EOF Then
   query = "insert into TAB_DETALHE_DECLARACAO"
   query = query & "(tdd_num_inscricao, tdd_periodo_inicial, tdd_substitutiva, "
   query = query & "tdd_item , tdd_valor ) "
   query = query & "values('" & Trim(num_inscricao) & "','" & periodo_inicial & "','"
   query = query & tipo_decla & "',"
   If (Val(Mid(detalhe, 2, 3)) < 217) Or (Val(Mid(detalhe, 2, 3)) > 222) Then
     valor = CDbl(Mid(detalhe, 5, 15) / 100)
     valor = TROCA_PONTO_DECIMAL(valor)
     query = query & Val(Mid(detalhe, 2, 3)) & "," & valor & ")"
   Else
       query = query & Val(Mid(detalhe, 2, 3)) & "," & Mid(detalhe, 5, 15) & ")"
   End If
Else
   query = "update TAB_DETALHE_DECLARACAO set "
   If (Val(Mid(detalhe, 2, 3)) < 217) Or (Val(Mid(detalhe, 2, 3)) > 222) Then
       valor = CDbl(Mid(detalhe, 5, 15) / 100)
       valor = TROCA_PONTO_DECIMAL(valor)
       query = query & "tdd_valor = " & valor
   Else
       query = query & "tdd_valor = " & Mid(detalhe, 5, 15)
   End If
   query = query & " where tdd_num_inscricao = '" & Trim(num_inscricao) & "'"
   query = query & " and tdd_periodo_inicial = '" & periodo_inicial & "'"
   query = query & " and tdd_substitutiva = " & Val(tipo_decla)
   query = query & " and tdd_item = " & Val(Mid(detalhe, 2, 3))

End If

snap_detalhe.Close
    
If Not EjecutarQuery(query) Then
End If

End Function
Function critica_tipo_seis(detalhe As String) As Integer

critica_tipo_seis = True
Dim snap_declaracao As Recordset
query = "Select *  From tab_cad_contadores "
query = query & "Where tco_uf_crc  = '" & Mid(detalhe, 2, 2) & "'"
query = query & "and  tco_num_crc  = '" & Mid(detalhe, 4, 10) & "'"
Set snap_declaracao = dbgim.OpenRecordset(query, dbOpenSnapshot)
    
If snap_declaracao.EOF Then
    query = "insert into tab_cad_contadores"
    query = query & "(tco_uf_crc, tco_num_crc, "
    query = query & " tco_num_cpf, tco_nome_contador ,"
    query = query & " tco_ti_logradouro, tco_logradouro, "
    query = query & "tco_numero , tco_complemento, "
    query = query & " tco_bairro ,tco_municipio , "
    query = query & " tco_cod_uf, tco_cep , "
    query = query & " tco_ddd_fone , tco_num_fone,  "
    query = query & " tco_ddd_fax , tco_num_fax ,"
    query = query & " tco_corr_eletronico )  "
    query = query & "values('" & Mid(detalhe, 2, 2) & "','" & Mid(detalhe, 4, 9) & "','"
    query = query & Mid(detalhe, 11, 11) & "','" & Mid(detalhe, 22, 32) & "','"
    query = query & Trim(Mid(detalhe, 54, 3)) & "','" & Mid(detalhe, 57, 30) & "','"
    query = query & Mid(detalhe, 87, 5) & "','" & Mid(detalhe, 92, 20) & "','"
    query = query & Mid(detalhe, 112, 20) & "','" & Trim(Mid(detalhe, 132, 5)) & "','"
    query = query & Mid(detalhe, 137, 2) & "','" & Mid(detalhe, 139, 8) & "','"
    query = query & Mid(detalhe, 147, 5) & "','" & Mid(detalhe, 152, 10) & "','"
    query = query & Mid(detalhe, 162, 5) & "','" & Mid(detalhe, 167, 10) & "','"
    query = query & Mid(detalhe, 177, 30) & "') "
    If Not EjecutarQuery(query) Then
       MsgBox "Os dados do Arquivo não estão Integros", 16, "Aviso do Sistema"
    End If
End If
snap_declaracao.Close
 


End Function

Function critica_tipo_tres(detalhe As String) As Integer
Dim snap_declaracao As Recordset
critica_tipo_tres = True

    'REGISTRO TIPO 3          pos
    'Tipo de registro       N1 1  "3"
    'Versao do gerador      N3 2-4 "V00"   '  para teste quando estiver ok colocar V01
    'CODIGO FORMULARIO      A3 5-7 "GIM"
    'NUMERO DE INSCRICAO    N9 8-16 TAB_DECLARACAO!TDE_NUM_INSCRICAO
    'PERIODO                A6 17-22 TAB_DECLARACAO!TDE_PERIODO_INICIAL
    'SUBSTITUTIVA           N2 23-24 TAB_DECLARACAO!TDE_SUBSTITUTIVA ->00/01/02...
    'REGIME DE DECLARACAO   N1 25-25 TAB_DECLARACAO!TDE_CONDICAO       SAIU -> TAB_CAD_CONTRIBUINTES!TCC_TIPO_CONTRIBUINTE
    'CONTROLE               N15 25-10 TAB_DECLARACAO!TDE_CONTROLE
    'HASH                   N4  CALCULADO

If Mid(detalhe, 2, 3) <> "V00" Then
 MsgBox "Erro no Registro Tipo tres."
  critica_tipo_tres = False
End If
If Mid(detalhe, 5, 3) <> "GIM" Then
 MsgBox "Erro no Registro Tipo tres."
  critica_tipo_tres = False
End If
If Not IsNumeric(Mid(detalhe, 8, 9)) Then
 MsgBox "Erro no Registro Tipo tres."
  critica_tipo_tres = False
End If
If Not IsNumeric(Mid(detalhe, 17, 6)) Then
 MsgBox "Erro no Registro Tipo tres."
  critica_tipo_tres = False
End If
If Not IsNumeric(Mid(detalhe, 23, 2)) Then
 MsgBox "Erro no Registro Tipo tres."
  critica_tipo_tres = False
End If
If Not IsNumeric(Mid(detalhe, 25, 1)) Then
 MsgBox "Erro no Registro Tipo tres."
  critica_tipo_tres = False
End If
num_inscricao = Mid(detalhe, 8, 9)
periodo_inicial = Mid(detalhe, 17, 6)
tipo_decla = Mid(detalhe, 23, 2)
condicao = Mid(detalhe, 25, 1)
    
    query = "Select *  From tab_declaracao "
    query = query & "Where tde_num_inscricao  = '" & num_inscricao & "'"
    query = query & " and tde_periodo_inicial = '" & periodo_inicial & "'"
    query = query & " and tde_substitutiva = " & tipo_decla
    query = query & " and tde_condicao = " & condicao

    Set snap_declaracao = dbgim.OpenRecordset(query, dbOpenSnapshot)
    
    If snap_declaracao.EOF Then
        query = "insert into tab_declaracao"
        query = query & "(tde_num_inscricao, tde_periodo_inicial, tde_substitutiva, "
        query = query & "tde_Estado , tde_aliquota_m17, tde_condicao ) "
        
        query = query & "values('" & Trim(num_inscricao) & "','" & periodo_inicial & "','"
        query = query & tipo_decla & "','"
        query = query & "T" & "'," & "0" & "," & condicao & ")"
        
        If Not EjecutarQuery(query) Then
        End If
    Else
        MsgBox "Esta declaração já existe "
    End If
    
    snap_declaracao.Close
    
    '****************************************

End Function

Function critica_tipo_zero(detalhe As String) As Integer
             
critica_tipo_zero = True

If Mid(detalhe, 2, 2) <> "GE" Then
 MsgBox "Erro no Registro Tipo zero"
  critica_tipo_zero = False
End If

End Function

Sub Ger_Reg03()
    
    Dim snap_declaracao     As Recordset
    Dim snap_contribuintes  As Recordset
    Dim controle As Double
    
    'somatorio dos valores do registro tipo 4
    controle = calcula_controle(num_inscricao, periodo_inicial, tipo_decla) * 100
    
    'REGISTRO TIPO 3
    'Tipo de registro       N1  "3"
    'Versao do gerador      N3  "V00"   '  para teste quando estiver ok colocar V01
    'CODIGO FORMULARIO      A3  "GIM"
    'NUMERO DE INSCRICAO    N9  TAB_DECLARACAO!TDE_NUM_INSCRICAO
    'PERIODO                A6  TAB_DECLARACAO!TDE_PERIODO_INICIAL
    'SUBSTITUTIVA           N2  TAB_DECLARACAO!TDE_SUBSTITUTIVA ->00/01/02...
    'REGIME DE DECLARACAO   N1  TAB_DECLARACAO!TDE_CONDICAO       SAIU -> TAB_CAD_CONTRIBUINTES!TCC_TIPO_CONTRIBUINTE
    'CONTROLE               N15 TAB_DECLARACAO!TDE_CONTROLE
    'HASH                   N4  CALCULADO

    'TAB_DECLARACAO
    Dim reg3 As String
    
    query = "Select *  From tab_declaracao "
    query = query & "Where tde_num_inscricao  = '" & num_inscricao & "'"
    query = query & " and tde_periodo_inicial = '" & periodo_inicial & "'"
    query = query & " and tde_substitutiva = " & tipo_decla

    Set snap_declaracao = dbgim.OpenRecordset(query, dbOpenSnapshot)
    
       
    'TAB_CAD_CONTRIBUINTES
    query = "Select tcc_tipo_contribuinte  From tab_cad_contribuintes "
    query = query & "Where tcc_nu_inscricao  = '" & num_inscricao & "'"
    
    Set snap_contribuintes = dbgim.OpenRecordset(query, dbOpenSnapshot)
    
    reg3 = "3" & Versao & "GIM" & snap_declaracao!tde_num_inscricao & snap_declaracao!tde_periodo_inicial
    reg3 = reg3 & Format$(snap_declaracao!tde_substitutiva, "00") & snap_declaracao!tde_condicao
    reg3 = reg3 & Format$(CStr(controle), "000000000000000")
    reg3 = reg3 & Calculo_crc(reg3)
    
    qtd_reg = qtd_reg + 1
    
    Grav_Reg reg3
    
    snap_contribuintes.Close
    snap_declaracao.Close
    
End Sub

Sub gera_txt()
    Dim i As Integer
    Dim qtd_decl As Long
    Dim ret As Long
    Dim condicao As Byte
    
    On Error GoTo errogeracao
    
    
    nFile = FreeFile(0)
    
    Open DrvSelec & "SEFAZ.GIM" For Output As nFile
    
    soma_hasharq = 0
    
    qtd_reg = 0
    
    Ger_Reg00
    
    qtd_reg = qtd_reg + 1
    
    For i = 1 To UBound(gera)
        
        num_inscricao = gera(i).inscricao
        periodo_inicial = gera(i).periodo_inicial
        tipo_decla = gera(i).tipo_decla
        data_geracao = gera(i).data_geracao
        condicao = gera(i).condicao
        Ger_Reg03
        'Call verifica_declaracao(num_inscricao, periodo_inicial, tipo_decla, condicao)
        
        Ger_Reg04
        
    Next
    
    Ger_Reg05 ("")
    Ger_Reg06 ("")
    Ger_Reg09
    
    Close #nFile
    'Atualizar o campo data de geracao nas tabelas
    
    atualiza_data_geracao
    MsgBox "Arquivo Gerado com Sucesso.", 48, NOME_SISTEMA_DEFAULT
    resposta = MsgBox("Deseja Imprimir Relação da(s) Declaração(ões) Gerada(s)?", vbYesNo + vbDefaultButton1, NOME_SISTEMA_DEFAULT)
    If resposta = 6 Then
      imprime_recibo
    End If
Exit Sub

errogeracao:

  Close #nFile
  MsgBox Error & "Não foi possível gerar o arquivo.", 48, NOME_SISTEMA_DEFAULT
    
End Sub
Function Calculo_crc(reg As String) As String
    
    Dim fim As Integer
    Dim soma As Double
    Dim i As Integer
    
    fim = Len(reg)
    
    soma = 0
    
    For i = 1 To fim
        
        soma = soma + (Asc(Mid(reg, i, 1)) * 2 ^ (i Mod 8))
        
    Next
    soma_hasharq = soma_hasharq + soma Mod 10000
    
    Calculo_crc = Format$(soma Mod 10000, "0000")
    
End Function
Sub Ger_Reg00()
      Dim reg0 As String
      
      reg0 = "0" & "GE" & Versao & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Now, "YYYY")
      reg0 = reg0 & Calculo_crc(reg0)
      
      Grav_Reg reg0
      
End Sub


Sub Grav_Reg(registro As String)
   
    On Error GoTo errodisco

    Print #nFile, registro
    
Exit Sub

errodisco:
    
    trata_retorno
    
End Sub
Sub Ger_Reg04()
    
    'REGISTRO TIPO 4
    'Tipo de registro       N1  "4"
    'ITEM                   N3  TAB_DETALHE_DECLARACAO!TDD_ITEM
    'VALOR                  N15 TAB_DETALHE_DECLARACAO!TDD_VALOR
    'HASH                   N4  CALCULADO
    
    'TAB_DETALHE_DECALRACAO
    Dim reg4 As String
    
    query = "Select *  From TAB_DETALHE_DECLARACAO "
    query = query & "Where tdd_num_inscricao  = '" & num_inscricao & "'"
    query = query & " and tdd_periodo_inicial = '" & periodo_inicial & "'"
    query = query & " and tdd_substitutiva = " & tipo_decla & " order by tdd_item "
    
    Set snap_detalhe = dbgim.OpenRecordset(query, dbOpenSnapshot)
      
    While Not snap_detalhe.EOF
        reg4 = ""
        
        reg4 = "4" & Format$(snap_detalhe!tdd_item, "000")
        If (Val(snap_detalhe!tdd_item) < 217) Or (Val(snap_detalhe!tdd_item) > 222) Then
            reg4 = reg4 & Format$(CStr(snap_detalhe!tdd_valor * 100), "000000000000000")
        Else
            'para estes item o valor corresponde o mes do recolhimento
            reg4 = reg4 & Format$(CStr(snap_detalhe!tdd_valor), "000000000000000")
        End If
        
        reg4 = reg4 & Calculo_crc(reg4)
        
        qtd_reg = qtd_reg + 1
        
        Grav_Reg reg4
        
        snap_detalhe.MoveNext
        
    Wend
    
    snap_detalhe.Close
    

End Sub
'CONTRIBUINTE
'
Sub Ger_Reg05(data_geracao As String)
    
   
    Dim snap_contrib As Recordset
    
    Dim reg5 As String
    
    Dim c1 As String * 9
    Dim c2 As String * 14
    Dim c3 As String * 59
    Dim c4 As String * 40
    Dim c5 As String * 3
    Dim c6 As String * 30
    Dim c7 As String * 5
    Dim c8 As String * 20
    Dim c9 As String * 20
    Dim c10 As String * 5
    Dim c11 As String * 2
    Dim c12 As String * 8
    Dim c13 As String * 5
    Dim c14 As String * 10
    Dim c15 As String * 5
    Dim c16 As String * 10
    Dim c17 As String * 30
    Dim c18 As String * 5
    Dim c19 As String * 2
    Dim c20 As String * 2
    Dim c21 As String * 7
    Dim c22 As String * 8
    Dim c23 As String * 1
    Dim c24 As String * 1
    Dim c25 As String * 1
    
    On Error GoTo erroreg5
    
    query = "Select * From tab_cad_contribuintes "
    query = query & "Where tcc_estado = 'A'"
    'quando a selecao for por data de geracao
    If data_geracao <> "" Then
        query = query & " and tcc_da_geracao = '" & data_geracao & "'"
     End If

    Set snap_contrib = dbgim.OpenRecordset(query, dbOpenSnapshot)
    
    While Not snap_contrib.EOF()
        
         c1 = Space(9)
         c2 = Space(14)
         c3 = Space(59)
         c4 = Space(40)
         c5 = Space(3)
         c6 = Space(30)
         c7 = Space(5)
         c8 = Space(20)
         c9 = Space(20)
         c10 = Space(5)
         c11 = Space(2)
         c12 = Space(8)
         c13 = Space(5)
         c14 = Space(10)
         c15 = Space(5)
         c16 = Space(10)
         c17 = Space(30)
         c18 = Space(5)
         c19 = Space(2)
         c20 = Space(2)
         c21 = Space(7)
         c22 = Space(8)
         c23 = Space(1)
         c24 = Space(1)
         c25 = Space(1)
    
    
        c1 = "" & snap_contrib("tcc_nu_inscricao")
        c2 = "" & snap_contrib("tcc_cad_geral_contrib")
        c3 = "" & snap_contrib("tcc_razao_social")
        c4 = "" & snap_contrib("tcc_no_fantasia")
        c5 = "" & snap_contrib("tcc_ti_logradouro")
        c6 = "" & snap_contrib("tcc_no_logradouro")
        c7 = "" & snap_contrib("tcc_nu_logradouro")
        c8 = "" & snap_contrib("tcc_complemento")
        c9 = "" & snap_contrib("tcc_bairro")
        c10 = "" & snap_contrib("tcc_municipio")
        c11 = "" & snap_contrib("tcc_cod_uf")
        c12 = "" & snap_contrib("tcc_cep")
        c13 = "" & snap_contrib("tcc_ddd_fone")
        c14 = "" & snap_contrib("tcc_num_fone")
        c15 = "" & snap_contrib("tcc_ddd_fax")
        c16 = "" & snap_contrib("tcc_num_fax")
        c17 = "" & snap_contrib("tcc_corr_eletronico")
        c18 = "" & snap_contrib("tcc_ativ_economica")
      '  c19 = "" & snap_contrib("tcc_reg_declaracao")
        c20 = "" & snap_contrib("tcc_uf_crc_contador")
        c21 = "" & snap_contrib("tcc_num_crc_contador")
        c22 = "" & snap_contrib("tcc_da_inicio_contador")
        c23 = "" & snap_contrib("tcc_relacao_contador")
        c24 = "" & snap_contrib("tcc_loc_livro_fiscal")
        c25 = "" & snap_contrib("tcc_tipo_contribuinte")
        
        reg5 = ""
        
        reg5 = "5" & c1 & c2 & c3 & c4 & c5 & c6 & c7 & c8 & c9 & c10 & c11 & c12 & c13 & c14
        reg5 = reg5 & c15 & c16 & c17 & c18 & c20 & c21 & c22 & c23 & c24 & c25
        reg5 = reg5 & Calculo_crc(reg5)
        
        qtd_reg = qtd_reg + 1
        
        Grav_Reg reg5
        
        snap_contrib.MoveNext
        
   Wend
    
   snap_contrib.Close
Exit Sub

erroreg5:
   
MsgBox Error(Err), 48, NOME_SISTEMA_DEFAULT

End Sub

Sub Ger_Reg06(data_geracao As String)
    
    Dim snap_contador As Recordset
    Dim reg6 As String
    
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim c5 As String
    Dim c6 As String
    Dim c7 As String
    Dim c8 As String
    Dim c9 As String
    Dim c10 As String
    Dim c11 As String
    Dim c12 As String
    Dim c13 As String
    Dim c14 As String
    Dim c15 As String
    Dim c16 As String
    Dim c17 As String
    
    query = "Select * From tab_cad_contadores "
    query = query & "Where tco_estado = 'A'"
    
    'quando a selecao for por data de geracao
    If data_geracao <> "" Then
        query = query & " and tco_da_geracao = '" & data_geracao & "'"
     End If

    Set snap_contador = dbgim.OpenRecordset(query, dbOpenSnapshot)
    
    While Not snap_contador.EOF()
        
         c1 = Space(9)
         c2 = Space(14)
         c3 = Space(59)
         c4 = Space(40)
         c5 = Space(3)
         c6 = Space(30)
         c7 = Space(5)
         c8 = Space(20)
         c9 = Space(20)
         c10 = Space(5)
         c11 = Space(2)
         c12 = Space(8)
         c13 = Space(5)
         c14 = Space(10)
         c15 = Space(5)
         c16 = Space(10)
         c17 = Space(30)
    
        c1 = "" & snap_contador("tco_uf_crc")
        c2 = "" & snap_contador("tco_num_crc")
        c3 = "" & snap_contador("tco_num_cpf")
        c4 = "" & snap_contador("tco_nome_contador")
        c5 = "" & snap_contador("tco_ti_logradouro")
        c6 = "" & snap_contador("tco_logradouro")
        c7 = "" & snap_contador("tco_numero")
        c8 = "" & snap_contador("tco_complemento")
        c9 = "" & snap_contador("tco_bairro")
        c10 = "" & snap_contador("tco_municipio")
        c11 = "" & snap_contador("tco_cod_uf")
        c12 = "" & snap_contador("tco_cep")
        c13 = "" & snap_contador("tco_ddd_fone")
        c14 = "" & snap_contador("tco_num_fone")
        c15 = "" & snap_contador("tco_ddd_fax")
        c16 = "" & snap_contador("tco_num_fax")
        c17 = "" & snap_contador("tco_corr_eletronico")
        
        reg6 = "6" & c1 & c2 & c3 & c4 & c5 & c6 & c7 & c8 & c9 & c10 & c11 & c12 & c13 & c14
        reg6 = reg6 & c15 & c16 & c17
        reg6 = reg6 & Calculo_crc(reg6)
        
        qtd_reg = qtd_reg + 1
        Grav_Reg reg6
        
        snap_contador.MoveNext
        
   Wend
   
    snap_contador.Close
    
End Sub
Sub Ger_Reg09()
    Dim reg9 As String
    
    msgerro = "Trailler"
    
    qtd_reg = qtd_reg + 1
    
    reg9 = "9" & Format$(qtd_reg, "000000")
    reg9 = reg9 & Calculo_crc(reg9) & Format$(Int(soma_hasharq / 10000), "0000")
    
    Grav_Reg reg9
    
End Sub

Sub importa_txt()

Dim caminho As String
Dim resp As Integer
Dim tipo_reg As String
Dim erro_abertura As Integer
Dim detalhe As String
le_disco:
On Error GoTo erro_abertura
nFile = FreeFile(0)
    
Open DrvSelec & "DIEF.DEC" For Input As nFile

If erro_abertura Then
       Close #nFile
       Unload Me
       Exit Sub
End If
    
Do While Not EOF(nFile)

      Line Input #nFile, detalhe
      tipo_reg = Left$(detalhe, 1)
      Select Case tipo_reg
          Case "0"
           If Not critica_tipo_zero(detalhe) Then
              Close #nFile
              Unload Me
              Exit Sub
           End If
          Case "3"
           If Not critica_tipo_tres(detalhe) Then
              Close #nFile
              Unload Me
              Exit Sub
           End If
          Case "4"
           If Not critica_tipo_quatro(detalhe) Then
              Close #nFile
              Unload Me
              Exit Sub
           End If
          Case "5"
           If Not critica_tipo_cinco(detalhe) Then
              Close #nFile
              Unload Me
              Exit Sub
           End If
          Case "6"
           If Not critica_tipo_seis(detalhe) Then
              Close #nFile
              Unload Me
              Exit Sub
           End If
           
           
       End Select
Loop
Close #nFile
  
MsgBox "Arquivo Importado com sucesso."
Exit Sub
  
sai:
 Exit Sub
erro_abertura:
    Select Case Err
      Case 71
        resp = MsgBox(" Não há disco na Unidade." & "Insira um disco e tente novamente.", 49, "Atenção")
        If resp = 1 Then
           Resume le_disco
        Else
           erro_abertura = True
           Resume sai
        End If
      Case Else
        MsgBox "Coloque o disquete com o Arquivo DIEF.DEC na Unidade selecionada"
        erro_abertura = True
        Exit Sub
    End Select
End Sub

Function TROCA_PONTO_DECIMAL(VARIAVEL As String) As String
Dim i As Integer
Dim temp As String
    temp = ""
    
    For i = 1 To Len(VARIAVEL)
       If IsNumeric(Mid(VARIAVEL, i, 1)) Then
            temp = temp & Mid(VARIAVEL, i, 1)
       ElseIf Mid(VARIAVEL, i, 1) = "," Then
            temp = temp & "."
        End If
    Next
    TROCA_PONTO_DECIMAL = temp


End Function

Private Sub cmdgerar_Click()
  
 On Error Resume Next

Dim indreg As Integer
Dim diret_cor As String
Dim tem_arquivos() As String
Dim ret As Integer
Dim a As Long
Screen.MousePointer = 11

Repetir:
    
    Kill DrvSelec & "~CIAT.tmp"
    If Err <> 0 And Err <> 53 Then
      ret = trata_retorno()
        If ret = 4 Then
            GoTo Repetir
        ElseIf ret = 2 Then
            If Mid$(UCase(DrvSelec), 1, 2) = "A:" Then
               optDrv(1).SetFocus
            Else
                optDrv(2).SetFocus
             End If
              Screen.MousePointer = 0
            Exit Sub
        End If
     End If
    Err = 0
    

cmdgerar.Enabled = False
cmdSair.Enabled = False
   
importa_txt
    
cmdgerar.Enabled = True
cmdSair.Enabled = True
Screen.MousePointer = 0
 
End Sub

Sub abre_arquivo()
    Dim caminho As String
    Dim resp As Integer
    Dim arquivo As String
    Dim Arq_Imp, Nome_Arq_Imp As String

le_disco:
   
    arquivo = FreeFile

    If Right$(Arq_Imp, 1) = "\" Then
        caminho = Arq_Imp & Nome_Arq_Imp
    Else
        caminho = Arq_Imp & "\" & Nome_Arq_Imp
    End If
    
    'caminho = Arq_Imp

    Open caminho For Input As #arquivo

sai:

Exit Sub

End Sub
Private Sub cmdSair_Click()
 Unload Me
End Sub

Private Sub Form_Load()
    CentrarFormu Me
    ChDrive Left$(App.Path, 2)
    diretorio = CurDir
    ChDir App.Path
    diretorio = CurDir

End Sub


Private Sub optDrv_Click(Index As Integer, Value As Integer)
    If Index = 1 Then
       DrvSelec = "a:\"
   Else
        DrvSelec = "b:\"
   End If

End Sub


Private Sub optDrv_GotFocus(Index As Integer)

      optDrv_Click Index, True
   
End Sub


Private Sub optDrv_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Or KeyCode = 10 Then
       optDrv_Click Index, True
    End If

End Sub


