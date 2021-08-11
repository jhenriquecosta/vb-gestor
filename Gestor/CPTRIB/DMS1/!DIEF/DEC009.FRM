VERSION 5.00
Begin VB.Form GerEnvio 
   BorderStyle     =   0  'None
   Caption         =   "SIAT - Declaração Eletrônica:  Gerar Envio"
   ClientHeight    =   1785
   ClientLeft      =   2010
   ClientTop       =   4755
   ClientWidth     =   4245
   ClipControls    =   0   'False
   FillColor       =   &H00C0C0C0&
   HelpContextID   =   8
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1785
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Selecione a unidade de gravação:"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   90
      TabIndex        =   1
      Top             =   600
      Width           =   2835
      Begin VB.PictureBox optDrv 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   1800
         ScaleHeight     =   315
         ScaleWidth      =   525
         TabIndex        =   3
         Top             =   390
         Width           =   585
      End
      Begin VB.PictureBox optDrv 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   420
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
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   1890
      Visible         =   0   'False
      Width           =   5295
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
      Left            =   90
      ScaleHeight     =   225
      ScaleWidth      =   3915
      TabIndex        =   8
      Top             =   90
      Width           =   3975
      Begin VB.Image Image1 
         Height          =   240
         Left            =   225
         Picture         =   "DEC009.frx":0000
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.PictureBox cmdSair 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2970
      ScaleHeight     =   315
      ScaleWidth      =   1005
      TabIndex        =   5
      Top             =   1080
      Width           =   1065
   End
   Begin VB.PictureBox cmdgerar 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2970
      ScaleHeight     =   315
      ScaleWidth      =   1005
      TabIndex        =   4
      Top             =   660
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   1695
      Left            =   30
      Top             =   60
      Width           =   4185
   End
   Begin VB.Label Label2 
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1860
      TabIndex        =   7
      Top             =   450
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inscrição Estadual: "
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
   End
End
Attribute VB_Name = "GerEnvio"
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


Sub Ger_Reg03()
    
    Dim snap_declaracao     As Recordset
    Dim snap_contribuintes  As Recordset
    Dim controle As Double
    Dim RAZAO As String * 59

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
    'RAZAO SOCIAL           N1  tab_cad_contribuintes!tcc_razao_social

    'CONTROLE               N15 TAB_DECLARACAO!TDE_CONTROLE
    'HASH                   N4  CALCULADO

    'TAB_DECLARACAO
    Dim reg3 As String
    
    query = "Select *  From tab_declaracao "
    query = query & "Where tde_num_inscricao  = '" & num_inscricao & "'"
    query = query & " and tde_periodo_inicial = '" & periodo_inicial & "'"
    query = query & " and tde_substitutiva = " & tipo_decla

    Set snap_declaracao = dbgim.OpenRecordset(query, dbOpenSnapshot)
    
    'If snap_declaracao.EOF Then
    '    MsgBox "nao achei decalaracao no reg03"
    'End If
    
 '   If Trim(Str(controle)) <> snap_declaracao!tde_controle Then
 '      MsgBox "controle nao bateu "
 '      Exit Sub
 '   End If
       
    'TAB_CAD_CONTRIBUINTES
    query = "Select tcc_razao_social,tcc_tipo_contribuinte  From tab_cad_contribuintes "
    query = query & "Where tcc_nu_inscricao  = '" & num_inscricao & "'"
    
    Set snap_contribuintes = dbgim.OpenRecordset(query, dbOpenSnapshot)
    RAZAO = Space(59)
    RAZAO = snap_contribuintes("tcc_razao_social")

    reg3 = "3" & Versao & "GIM" & snap_declaracao!tde_num_inscricao & snap_declaracao!tde_periodo_inicial
    reg3 = reg3 & Format$(snap_declaracao!tde_substitutiva, "00") & snap_declaracao!tde_condicao
    reg3 = reg3 & RAZAO
    If Len(CStr(controle)) > 15 Then
        controle = 0
    End If
    
    reg3 = reg3 & Format$(CStr(controle), "000000000000000")
    reg3 = reg3 & Calculo_crc(reg3)
    
    qtd_reg = qtd_reg + 1
    
    Grav_Reg reg3
    
    snap_contribuintes.Close
    snap_declaracao.Close
    
    '****************************************
End Sub

Sub gera_txt()
    Dim i As Integer
    Dim qtd_decl As Long
    Dim ret As Long
    Dim condicao As Byte
    
    On Error GoTo errogeracao
    
    
    nFile = FreeFile(0)
    
    Open DrvSelec & "DIEF.DEC" For Output As nFile
    
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
  
    
    Calculo_crc = Format$(soma Mod 10000, "0000")   'foi alterado para mod 100.000
  
    
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
        If (Val(snap_detalhe!tdd_item) < 217) Or (Val(snap_detalhe!tdd_item) > 219) Then
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
    Dim c26 As String * 8
    
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
         c26 = Space(8)
    
    
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
        c20 = "" & snap_contrib("tcc_uf_crc_contador")
        c21 = "" & snap_contrib("tcc_num_crc_contador")
        c22 = "" & snap_contrib("tcc_da_inicio_contador")
        c23 = "" & snap_contrib("tcc_relacao_contador")
        c24 = "" & snap_contrib("tcc_loc_livro_fiscal")
        c25 = "" & snap_contrib("tcc_tipo_contribuinte")
        c26 = "" & snap_contrib("tcc_da_inicio_op")
        
        reg5 = ""
        
        reg5 = "5" & c1 & c2 & c3 & c4 & c5 & c6 & c7 & c8 & c9 & c10 & c11 & c12 & c13 & c14
        reg5 = reg5 & c15 & c16 & c17 & c18 & c20 & c21 & c22 & c23 & c24 & c25 & c26
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
    
    Dim c1 As String * 2
    Dim c2 As String * 7
    Dim c3 As String * 11
    Dim c4 As String * 32
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
    'Dim c18 As String * 11
    
    query = "Select * From tab_cad_contadores "
    query = query & "Where tco_estado = 'A'"
    'quando a selecao for por data de geracao
    If data_geracao <> "" Then
        query = query & " and tco_da_geracao = '" & data_geracao & "'"
     End If
    Set snap_contador = dbgim.OpenRecordset(query, dbOpenSnapshot)
    While Not snap_contador.EOF()
         c1 = Space(2)
         c2 = Space(7)
         c3 = Space(11)
         c4 = Space(32)
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
         'c18 = Space(11)
    
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
        'c18 = "" & snap_contador("tco_cpf_conj")
        
        reg6 = "6" & c1 & c2 & c3 & c4 & c5 & c6 & c7 & c8 & c9 & c10 & c11 & c12 & c13 & c14
        reg6 = reg6 & c15 & c16 & c17
        '& c18
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
    
    ReDim tem_arquivos(1)
    tem_arquivos(0) = Dir(DrvSelec & "*.*")
    If tem_arquivos(0) <> "" Then
        MsgBox "Coloque um disquete vazio.", MB_ICONEXCLAMATION, NOME_SISTEMA_DEFAULT
        If Mid$(UCase(DrvSelec), 1, 2) = "A:" Then
           optDrv(1).SetFocus
        Else
           optDrv(2).SetFocus
        End If
         Screen.MousePointer = 0
        Exit Sub
    End If
 
   cmdgerar.Enabled = False
   cmdSair.Enabled = False
   
   gera_txt
    
   cmdgerar.Enabled = True
   cmdSair.Enabled = True
  
 Screen.MousePointer = 0
 
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


