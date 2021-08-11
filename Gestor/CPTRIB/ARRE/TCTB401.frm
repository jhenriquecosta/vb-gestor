VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCTB401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   5415
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   1085
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   465
         Left            =   7950
         TabIndex        =   9
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   820
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprime 
         Height          =   465
         Left            =   4740
         TabIndex        =   8
         Top             =   90
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   820
         Caption         =   "&Imprimir Resumo"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   465
         Left            =   6720
         TabIndex        =   7
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   820
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   465
         Left            =   9150
         TabIndex        =   10
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   820
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   60
      ScaleHeight     =   585
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCTB401.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Threed.SSFrame fra 
      Height          =   1695
      Index           =   2
      Left            =   30
      TabIndex        =   12
      Top             =   690
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   2990
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      ShadowStyle     =   1
      Begin VTOcx.cboVISUAL cboAgente 
         Height          =   510
         Left            =   90
         TabIndex        =   1
         Top             =   540
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   900
         Caption         =   "Banco"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.txtVISUAL txtNumLote 
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   30
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   873
         Caption         =   "No. Lote"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.cboVISUAL cboCodSucursal 
         Height          =   510
         Left            =   90
         TabIndex        =   2
         Top             =   1050
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   900
         Caption         =   "Agencia"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboNumConta 
         Height          =   510
         Left            =   1890
         TabIndex        =   3
         Top             =   1050
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   900
         Caption         =   "Conta Corrente"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboSit 
         Height          =   510
         Left            =   3720
         TabIndex        =   4
         Top             =   1050
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   900
         Caption         =   "Situacão Lote"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.txtVISUAL txtDataInicial 
         Height          =   495
         Left            =   6510
         TabIndex        =   5
         Top             =   1080
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   873
         Caption         =   "Data Inicial Lote"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtDataFinal 
         Height          =   495
         Left            =   8370
         TabIndex        =   6
         Top             =   1080
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   873
         Caption         =   "Data Final Lote"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
      End
   End
   Begin VTOcx.grdVISUAL lstLote 
      Height          =   2805
      Left            =   60
      TabIndex        =   13
      Top             =   2430
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4948
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   690
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   1138
      Icone           =   "TCTB401.frx":2123
   End
End
Attribute VB_Name = "TCTB401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim CodImposto As String
Dim NumAgente  As Double
Dim NumLote As Double

Private Sub cboAgente_Click()
    If cboAgente.ListIndex >= 0 Then
        cboCodSucursal.Preencher Bdados, "Select tcb_cod_sucursal from tab_conta_bancaria where tcb_tar_cod_agente =" & cboAgente.Coluna(0).Valor
    End If
End Sub

Private Sub cboCodSucursal_Click()
    cboNumConta.Preencher Bdados, "Select tcb_num_conta from tab_conta_bancaria where tcb_tar_cod_agente =" & cboAgente.Coluna(0).Valor & " and tcb_cod_sucursal ='" & cboCodSucursal & "'"
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdImprime_Click()
    Dim SELECAO As String
    Screen.MousePointer = 11
    With Rpt
        SELECAO = " {TAB_DARM_RECEBIDO.TDR_SIT_PAGO} <> 2 "
        
        If Trim(txtNumLote) <> "" Then
            SELECAO = SELECAO & "  and {TAB_LOTE_PAGAMENTO.TLP_COD_LOTE} =" & txtNumLote
        End If
        If Trim(txtDataInicial) <> "" And Trim(txtDataFinal) = "" Then
            SELECAO = SELECAO & "  and ( {TAB_LOTE_PAGAMENTO.TLP_DATA_ARRECADACAO} =  Date (" & Year(txtDataInicial) & "," & Month(txtDataInicial) & "," & Day(txtDataInicial) & ")   and {TAB_DARM_RECEBIDO.TDR_SIT_PAGO} <> 2"
        ElseIf Trim(txtDataInicial) <> "" And Trim(txtDataFinal) <> "" Then
            SELECAO = SELECAO & "  and ( {TAB_LOTE_PAGAMENTO.TLP_DATA_ARRECADACAO} in  Date (" & Year(txtDataInicial) & "," & Month(txtDataInicial) & "," & Day(txtDataInicial) & _
            ") to Date (" & Year(txtDataFinal) & "," & Month(txtDataFinal) & "," & Day(txtDataFinal) & ")   and {TAB_DARM_RECEBIDO.TDR_SIT_PAGO} <> 2"
        End If
        If cboAgente <> "" Then
            SELECAO = SELECAO & "  and {TAB_LOTE_PAGAMENTO.TLP_TAR_COD_AGENTE} = " & cboAgente.Coluna(0).Valor
        End If
        If Trim(cboSit) <> "" Then
            SELECAO = SELECAO & "  and {TAB_LOTE_PAGAMENTO.TLP_SITUACAO_LOTE} = " & cboSit.Coluna(1).Valor
        End If
        If Trim(cboCodSucursal) <> "" Then
            SELECAO = SELECAO & "  AND {TAB_LOTE_PAGAMENTO.TLP_NUM_SUCURSAL} ='" & cboCodSucursal & "'"
        End If
        If Trim(cboNumConta) <> "" Then
            SELECAO = SELECAO & "  AND {TAB_LOTE_PAGAMENTO.TLP_NUM_CONTA} ='" & cboNumConta & "'"
        End If
        If Trim(txtNumLote) = "" Then
            If Not .DefinirArquivo(Bdados, App.Path & "\TLoteResumo.rpt") Then Exit Sub
        Else
            If Not .DefinirArquivo(Bdados, App.Path & "\TLoteResumoEspec.rpt") Then Exit Sub
        End If
        .SELECAO = SELECAO
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        .Arvore = False
        .Visualizar
    End With
    Set Rpt = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtNumLote.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Condicao As String
    Dim Controle As Control
    Dim Conteudo As String
    'On Error Resume Next
    Screen.MousePointer = 11
    Sql = "SELECT tdr_tgt_cod_pagamento as Cod_Pagamento, tdr_inscricao as Inscricao," & _
        " tdr_tip_cod_imposto as Cod_Receita,tdr_periodo as Periodo, " & _
        " tdr_data_vencimento as Dt_Vence,tdr_data_pagamento as Dt_Pago," & _
        Bdados.Converte("tdr_valor_original", TCDuplo) & " as Vl_Original," & _
        Bdados.Converte("tdr_juros", TCDuplo) & " as Juros, " & _
        Bdados.Converte("tdr_valor_real_juros", TCDuplo) & " as Juros_Pago, " & _
        Bdados.Converte("tdr_multa", TCDuplo) & "as Multa, " & _
        Bdados.Converte("tdr_valor_real_Multa", TCDuplo) & " as Multa_Pago, " & _
        Bdados.Converte("tdr_valor_total", TCDuplo) & " as Vl_Total," & _
        Bdados.Converte("tdr_valor_real_pago", TCDuplo) & " as Vl_Pago,TLP_COD_LOTE as Lote, " & _
        " TLP_TAR_COD_AGENTE as Agente from TAB_DARM_RECEBIDO,TAB_LOTE_PAGAMENTO " & _
        " where  TLP_COD_LOTE=TDR_TLP_COD_LOTE  and tdr_sit_pago <> 2 "

    Condicao = ""
    If Trim(txtNumLote) <> "" Then
        Condicao = " AND TLP_COD_LOTE =" & txtNumLote
    End If
    If Trim(txtDataInicial) <> "" And Trim(txtDataFinal) = "" Then
        Condicao = Condicao & " AND TLP_DATA_ARRECADACAO =" & Bdados.Converte(txtDataInicial, TCDataHora)
    ElseIf Trim(txtDataInicial) <> "" And Trim(txtDataFinal) <> "" Then
        Condicao = Condicao & " AND (TLP_DATA_ARRECADACAO >=" & Bdados.Converte(txtDataInicial, TCDataHora) & _
                " AND TLP_DATA_ARRECADACAO  <=" & Bdados.Converte(txtDataFinal, TCDataHora) & ")"
    End If
    If Trim(cboAgente) <> "" Then
        Condicao = Condicao & " AND TLP_TAR_COD_AGENTE =" & cboAgente.Coluna(0).Valor
    End If
    If Trim(cboCodSucursal) <> "" Then
        Condicao = Condicao & " AND TLP_NUM_SUCURSAL ='" & cboCodSucursal & "'"
    End If
    If Trim(cboNumConta) <> "" Then
        Condicao = Condicao & " AND TLP_NUM_CONTA ='" & cboNumConta & "'"
    End If
    If Trim(cboSit) <> "" Then
        Condicao = Condicao & " AND TLP_SITUACAO_LOTE =" & cboSit.Coluna(1).Valor
    End If
    Sql = Sql & Condicao
    lstLote.Preencher Bdados, Sql
    
    If lstLote.ListItems.Count > 0 Then lstLote.Mensagem = "Total: R$" & Format(lstLote.Colunas(13).Soma, Const_Monetario)
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim rs As VSRecordset
    cabVisual.Exibir Bdados, Me.Name, App.Path
    cboAgente.Clear
    cboCodSucursal.Clear
    cboNumConta.Clear
    cboSit.Clear
    cboAgente.Preencher Bdados, "Select tar_cod_agente,tar_nome_agente from tab_agente_arrecadador where tar_ativo =0", 1
    cboCodSucursal.Preencher Bdados, "Select tcb_cod_sucursal from tab_conta_bancaria "
    cboNumConta.Preencher Bdados, "Select tcb_num_conta from tab_conta_bancaria "
    
    cboSit.PreencherGeral Bdados, "STATUS GRADE LOTE"
    cboSit.AddItem " "
    cboAgente.AddItem " "
    cboCodSucursal.AddItem " "
    cboNumConta.AddItem " "
    AtualizaCabecalho lstLote
    DoEvents
End Sub

Private Sub lstLote_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstLote, ColumnHeader
End Sub

Private Sub txtDtArrecada_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtNumLote_LostFocus()
    On Error Resume Next
    Dim rs As VSRecordset
    Dim Sql As String
    
'    If Trim(txtNumLote) = "" Then Exit Sub
    Exit Sub
    Sql = "Select tar_nome_agente,TLP_NUM_SUCURSAL,TLP_NUM_CONTA," & _
        " TLP_DATA_ARRECADACAO,TLP_DATA_RECEPCAO,TLP_SITUACAO_LOTE from TAB_LOTE_PAGAMENTO, Tab_Agente_Arrecadador" & _
        " where TLP_TAR_COD_AGENTE=tar_cod_agente and TLP_COD_LOTE=" & txtNumLote
    If Bdados.AbreTabela(Sql, rs) Then
        cboAgente = rs!tar_nome_agente
        cboCodSucursal = Format(rs!TLP_NUM_SUCURSAL, "00000")
        cboNumConta = rs!TLP_NUM_CONTA
        txtDataInicial = rs!TLP_DATA_RECEPCAO
        cboSit.SetarLinha rs!TLP_SITUACAO_LOTE, 1
    Else
        cboAgente.ListIndex = -1
        cboCodSucursal.ListIndex = -1
        cboNumConta.ListIndex = -1
        cboSit.ListIndex = -1
    End If
    Bdados.FechaTabela rs
End Sub

