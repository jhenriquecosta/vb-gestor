VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TPRT101_OS 
   Caption         =   "TPRT101"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ActiveTabs.SSActiveTabs tabEtapa 
      Height          =   6255
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11033
      _Version        =   131082
      TabCount        =   3
      Tabs            =   "TPRT101_OS.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5865
         Left            =   -99969
         TabIndex        =   14
         Top             =   360
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   10345
         _Version        =   131082
         TabGuid         =   "TPRT101_OS.frx":00B7
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   5865
         Left            =   -99969
         TabIndex        =   11
         Top             =   360
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   10345
         _Version        =   131082
         TabGuid         =   "TPRT101_OS.frx":00DF
         Begin VTOcx.fraVISUAL fra 
            Height          =   1425
            Index           =   1
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   2514
            Altura          =   1905
            Caption         =   " Procedimento"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtDescricaoProcedimento 
               Height          =   1005
               Left            =   120
               TabIndex        =   3
               Top             =   360
               Width           =   5865
               _ExtentX        =   10345
               _ExtentY        =   1773
               Caption         =   "Descrição"
               Text            =   ""
               TipoLetras      =   0
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtValorProcedimento 
               Height          =   525
               Left            =   6000
               TabIndex        =   4
               Top             =   360
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   926
               Caption         =   "Valor"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               CorTexto        =   4194304
            End
            Begin VTOcx.cmdVISUAL cmdSalvarProcedimento 
               Height          =   375
               Left            =   6000
               TabIndex        =   5
               Top             =   990
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   661
               Caption         =   "Salvar"
               Acao            =   3
               CorBorda        =   8421504
               CorFrente       =   16384
            End
         End
         Begin VTOcx.grdVISUAL grdProcedimentos 
            Height          =   4215
            Left            =   0
            TabIndex        =   13
            Top             =   1560
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   7435
            CorBorda        =   16711680
            Caption         =   "Procedimentos Cadastrados"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5865
         Left            =   30
         TabIndex        =   7
         Top             =   360
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   10345
         _Version        =   131082
         TabGuid         =   "TPRT101_OS.frx":0107
         Begin VTOcx.fraVISUAL fra 
            Height          =   1065
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   120
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   1879
            Altura          =   1905
            Caption         =   " Filtro"
            CorTexto        =   16777215
            CorFaixa        =   16711680
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cboVISUAL cboPeriodo 
               Height          =   510
               Left            =   120
               TabIndex        =   0
               Tag             =   "Tipo"
               Top             =   360
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   900
               Caption         =   "Período"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
            Begin VTOcx.cmdVISUAL cmdGerar 
               Height          =   375
               Left            =   6000
               TabIndex        =   2
               Top             =   495
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   661
               Caption         =   "Gerar"
               Acao            =   3
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cboVISUAL cboFiscal 
               Height          =   510
               Left            =   1440
               TabIndex        =   1
               Tag             =   "Tipo"
               Top             =   360
               Width           =   4545
               _ExtentX        =   8017
               _ExtentY        =   900
               Caption         =   "Fiscal"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Alinhamento     =   1
            End
         End
         Begin VTOcx.fraVISUAL frameAltera 
            Height          =   4545
            Left            =   0
            TabIndex        =   15
            Top             =   1320
            Visible         =   0   'False
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   8017
            Altura          =   1905
            Caption         =   " Pontuação"
            CorTexto        =   0
            CorFaixa        =   12632256
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Begin VTOcx.cmdVISUAL cmdVoltar 
               Height          =   375
               Left            =   6240
               TabIndex        =   19
               Top             =   4080
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   661
               Caption         =   "Voltar"
               CorBorda        =   8421504
               CorFrente       =   16384
            End
            Begin VTOcx.cmdVISUAL cmdMais 
               Height          =   375
               Left            =   720
               TabIndex        =   18
               Top             =   4080
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   661
               Caption         =   "+1"
               Acao            =   1
               CorBorda        =   8421504
               CorFrente       =   16384
               CorFundo        =   8454016
            End
            Begin VTOcx.cmdVISUAL cmdMenos 
               Height          =   375
               Left            =   120
               TabIndex        =   17
               Top             =   4080
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   661
               Caption         =   "-1"
               Acao            =   9
               CorBorda        =   8421504
               CorFrente       =   16384
               CorFundo        =   8421631
            End
            Begin VTOcx.grdVISUAL grdPontos 
               Height          =   3975
               Left            =   50
               TabIndex        =   16
               Top             =   360
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   7011
               CorBorda        =   8421504
               Caption         =   "Pontos"
               CorTitulo       =   12632256
               CorCaption      =   0
               CorDica         =   16711680
            End
         End
         Begin VTOcx.grdVISUAL grdPesquisa 
            Height          =   4575
            Left            =   0
            TabIndex        =   10
            Top             =   1320
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   8070
            CorBorda        =   16711680
            Caption         =   "Listagem"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1138
      Formulario      =   "Pontuação"
      Descricao       =   "Controle de Pontuação"
      Icone           =   "TPRT101_OS.frx":012F
   End
End
Attribute VB_Name = "TPRT101_OS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codigoProcedimento As Integer
Dim item As Integer

Private Sub cboPeriodo_Click()
    Dim sql As String
    
    If grdPesquisa.Preencher(Bdados, "SELECT PERIODO,FISCAL FROM TAB_PROD_FISCAL WHERE PERIODO = '" & cboPeriodo.Text & "' ORDER BY PERIODO,FISCAL") Then
    End If
End Sub

Private Sub cmdGerar_Click()
    Dim sql
    sql = "INSERT INTO TAB_PROD_FISCAL (FISCAL, PERIODO,COMPENSADO) VALUES ('" & cboFiscal.Text & "','" & cboPeriodo.Text & "',0  );"
    Bdados.Executa (sql)
    Dim Rs1 As VSRecordset
    If Bdados.AbreTabela("SELECT CODIGO, VALOR FROM TAB_PROD_PROCEDIMENTO", Rs1) Then
        Do While Not Rs1.EOF
            sql = "INSERT INTO TAB_PROD_FISCAL_ITEM (FISCAL, PERIODO,QUANT,ITEM,VALOR) VALUES ('" & cboFiscal.Text & "','" & cboPeriodo.Text & "',0,"
            sql = sql & Rs1(0) & " , '" & Replace(Rs1(1), ",", ".") & "');"
            Bdados.Executa (sql)
            Rs1.MoveNext
        Loop
    End If
    Mensagem ("Registro salvo com SUCESSO!!")
    preenche
End Sub

Private Sub preenche()
    If grdPesquisa.Preencher(Bdados, "SELECT PERIODO,FISCAL FROM TAB_PROD_FISCAL ORDER BY PERIODO, FISCAL") Then
    End If
End Sub

Private Sub cmdMais_Click()
    atualizar (1)
End Sub

Private Sub cmdMenos_Click()
    atualizar (-1)
End Sub
Private Sub atualizar(numero As Integer)
    Dim sql As String
    sql = "UPDATE TAB_PROD_FISCAL_ITEM SET QUANT=QUANT + " & numero & " WHERE FISCAL ='" & cboFiscal & "' and PERIODO ='" & cboPeriodo & "' AND ITEM=" & item
    Bdados.Executa (sql)
    pontuacao
End Sub


Private Sub cmdSalvarProcedimento_Click()
    Dim sql, valor As String
    valor = Replace(txtValorProcedimento, ",", ".")
    If codigoProcedimento = 0 Then
        sql = "INSERT INTO TAB_PROD_PROCEDIMENTO (NOME, VALOR) VALUES ('" & txtDescricaoProcedimento & "','" & valor & "');"
    Else
        sql = "UPDATE TAB_PROD_PROCEDIMENTO SET NOME = '" & txtDescricaoProcedimento & "', "
        sql = sql & " VALOR ='" & valor & "' WHERE CODIGO = " & codigoProcedimento
    End If
    Bdados.Executa (sql)
    
    txtValorProcedimento = ""
    txtDescricaoProcedimento = ""
    codigoProcedimento = 0
    Mensagem ("Registro salvo com SUCESSO!!")
    preencheProcedimentos
    txtDescricaoProcedimento.SetFocus
End Sub

Private Sub cmdVoltar_Click()
    frameAltera.Visible = False
End Sub

Private Sub Form_Load()
    codigoProcedimento = 0
    preencheProcedimentos
    preenche
    
    Dim ano As String
    Dim x As Integer
    ano = Format(Now, "YYYY")
    For x = 1 To 12
        cboPeriodo.AddItem (Format(x, "00") & "-" & ano)
    Next x
    Dim Rs As VSRecordset
    If Bdados.AbreTabela("SELECT TUS_COD_USUARIO FROM TAB_USUARIO ORDER BY TUS_COD_USUARIO", Rs) Then
        Do While Not Rs.EOF
            cboFiscal.AddItem Rs(0)
            Rs.MoveNext
        Loop
    End If
End Sub
Private Sub preencheProcedimentos()
    If grdProcedimentos.Preencher(Bdados, "SELECT CODIGO,VALOR,NOME FROM TAB_PROD_PROCEDIMENTO") Then
    End If
End Sub

Private Sub grdPesquisa_DblClick()
    If grdPesquisa.ListItems.Count = 0 Then Exit Sub
    cboPeriodo.Text = grdPesquisa.SelectedItem
    cboFiscal.Text = grdPesquisa.SelectedItem.SubItems(1)
    frameAltera.Visible = True
    pontuacao
End Sub
Private Sub pontuacao()
    Dim sql As String
    sql = "SELECT ITEM,VALOR,QUANT,TOTAL,NOME FROM  VIS_PROD_FISCAL_ITEM WHERE FISCAL ='" & cboFiscal & "' and PERIODO ='" & cboPeriodo & "' ORDER BY ITEM"
    If grdPontos.Preencher(Bdados, sql) Then
    End If
End Sub

Private Sub grdPontos_Click()
    If grdPontos.ListItems.Count = 0 Then Exit Sub
    item = grdPontos.SelectedItem
End Sub

Private Sub grdProcedimentos_Click()
    If grdProcedimentos.ListItems.Count = 0 Then Exit Sub
    codigoProcedimento = grdProcedimentos.SelectedItem
    txtValorProcedimento = grdProcedimentos.SelectedItem.SubItems(1)
    txtDescricaoProcedimento = grdProcedimentos.SelectedItem.SubItems(2)
End Sub

