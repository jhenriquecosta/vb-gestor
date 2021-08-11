VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TNOT101 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   13
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TNOT101.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin ActiveTabs.SSActiveTabs tabNotificacao 
      Height          =   3240
      Left            =   75
      TabIndex        =   5
      Top             =   3015
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   5715
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TagVariant      =   ""
      Tabs            =   "TNOT101.frx":2123
      Images          =   "TNOT101.frx":21D4
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   2820
         Index           =   0
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   4974
         _Version        =   131082
         TabGuid         =   "TNOT101.frx":2E75
         Begin VTOcx.grdVISUAL lstNot 
            Height          =   2745
            Left            =   30
            TabIndex        =   10
            Top             =   45
            Width           =   10230
            _ExtentX        =   18045
            _ExtentY        =   4842
            CorFundo        =   -2147483633
            Caption         =   "Débitos em Aberto"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   2820
         Index           =   1
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   4974
         _Version        =   131082
         TabGuid         =   "TNOT101.frx":2E9D
         Begin VB.TextBox txtTexto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   2910
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   30
            Width           =   10170
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   2820
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   4974
         _Version        =   131082
         TabGuid         =   "TNOT101.frx":2EC5
         Begin VTOcx.grdVISUAL grdNotifica 
            Height          =   2760
            Left            =   45
            TabIndex        =   4
            Top             =   30
            Width           =   10260
            _ExtentX        =   18098
            _ExtentY        =   4868
            CorFundo        =   -2147483633
            Caption         =   "Notificacoes emitidas"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   7
      Top             =   6375
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   900
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   4680
         TabIndex        =   14
         Top             =   90
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         Caption         =   "Limpar Historico"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprime 
         Height          =   375
         Left            =   7455
         TabIndex        =   1
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdParcela 
         Height          =   375
         Left            =   6480
         TabIndex        =   0
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Gerar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9720
         TabIndex        =   3
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   8670
         TabIndex        =   2
         Top             =   90
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   1138
      Icone           =   "TNOT101.frx":2EED
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   2265
      Left            =   45
      TabIndex        =   15
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   675
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   3995
      Altura          =   1905
      Caption         =   " Dados do Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   300
         Left            =   900
         TabIndex        =   23
         Top             =   1815
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         Caption         =   "Período"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1590
         TabIndex        =   28
         Top             =   1440
         Width           =   7725
      End
      Begin VB.TextBox txtIm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1605
         TabIndex        =   27
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox txtContrib 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1590
         TabIndex        =   26
         Top             =   1080
         Width           =   7725
      End
      Begin VB.ComboBox cboDest 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "TNOT101.frx":3207
         Left            =   7875
         List            =   "TNOT101.frx":3211
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1800
         Width           =   1485
      End
      Begin VTOcx.txtVISUAL txtData 
         Height          =   285
         Left            =   4650
         TabIndex        =   24
         Top             =   1815
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   503
         Caption         =   "Venc. Notificação"
         Text            =   ""
         Formato         =   0
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   0
         Left            =   7110
         TabIndex        =   22
         Top             =   1830
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   423
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   -2147483626
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Destino"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   5
         RoundedCorners  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   945
         TabIndex        =   21
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaIM 
         Height          =   315
         Left            =   2820
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   720
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   5490
         TabIndex        =   19
         Top             =   720
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "Cadastro do Imóvel"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   315
         Left            =   8970
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   720
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   1
         Left            =   795
         TabIndex        =   17
         Top             =   1095
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   423
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   -2147483626
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Nome"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   5
         RoundedCorners  =   0   'False
      End
      Begin Threed.SSPanel lbl 
         Height          =   240
         Index           =   2
         Left            =   720
         TabIndex        =   16
         Top             =   1470
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   423
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   -2147483626
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Endereço"
         BorderWidth     =   1
         BevelOuter      =   0
         AutoSize        =   3
         Alignment       =   5
         RoundedCorners  =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IM"
         Height          =   135
         Left            =   1350
         TabIndex        =   29
         Top             =   795
         Width           =   180
      End
   End
   Begin VB.Menu mnuNotifica 
      Caption         =   "."
      Visible         =   0   'False
      Begin VB.Menu mnuEmitir 
         Caption         =   "&Emitir notificação ..."
      End
   End
End
Attribute VB_Name = "TNOT101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim CodPagamento  As Double
Dim Documento As String
Private Rpt As New VSRelatorio

Private Sub cboDest_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub cboImposto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub


Private Sub cmdCancela_Click()
    Dim Rs As VSRecordset
    Dim Sql As String
    cboImposto.Enabled = True
    cmdParcela.Enabled = True
    Edita.LimpaCampos Me
    lstNot.ListItems.Clear
    lstNot.Mensagem = ""
    grdNotifica.Preencher Bdados, ""
    grdNotifica.Mensagem = ""
    lstNot.Preencher Bdados, ""
    CodPagamento = 0
    Sql = "Select TPT_TEXTO FROM TAB_PARAMETRO_TEXTO WHERE TPT_PARAMETRO = 'CARTA OMISSOS'"
    If Bdados.AbreTabela(Sql, Rs) Then
        txtTexto = "" & Rs!TPT_TEXTO
    End If
    cboImposto.SetFocus
End Sub

Private Sub cmdEnter_Click()
'    SendKeys "{TAB}"
End Sub

Private Sub cmdExcluir_Click()
    If grdNotifica.ListItems.Count >= 1 Then
        If Confirma("Deseja excluir?") Then
            Bdados.DeletaDados "Tab_Notificacao", "Tnt_cod_notificacao = '" & grdNotifica.SelectedItem & "'"
            Bdados.DeletaDados "tab_pagamento_extrato", " tpe_cod_pagamento_extrato  = '" & grdNotifica.SelectedItem & "'"
            Call PEga_Notificacao
        End If
    End If
End Sub

Private Sub cmdImprime_Click()
    On Error GoTo Trata
    Dim i As Integer
    Dim ImAnterior As String
    Dim SelecaoRpt As String
    Dim Conta As New ContaCorrente
    Dim Valores As String
    Dim campos As String
    Dim Cobranca As New VSCobranca
    Dim ValorFinal As Double
    Dim InsCad As String
    Dim InsMun As String
    Dim Insc As String
    Dim NovaData As String
    If grdNotifica.ListItems.Count <= 0 Then Exit Sub
    If txtIm <> "" Then
        Insc = txtIm
    Else
        Insc = txtImovel
    End If
    
    If cboDest.ListIndex < 0 Then
        Avisa "Informe destino da(s) notificação(s)."
        cboDest.SetFocus
        Exit Sub
    End If
    NovaData = Imposto.DataVencimentoNova(grdNotifica.SelectedItem.SubItems(3))
    If Trim(NovaData) = "" Then
        Avisa "Informe a data de vencimento."
        txtData.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = 11
    '1.
    Bdados.GravaDados "TAB_PARAMETRO_TEXTO", Bdados.PreparaValor(txtTexto), "TPT_TEXTO", "TPT_PARAMETRO = 'CARTA OMISSOS'"
    '2
    ImprimirNotificacao Insc, Trim(Mid(cboImposto, Edita.PosPic(cboImposto, "-") + 1)), CStr(ValorFinal), NovaData, grdNotifica.SelectedItem, cboDest.ListIndex, ""
    ImAnterior = ""
    CodPagamento = 0
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Avisa Err.Description
        Err.Clear
        Resume
    End If
End Sub
Private Sub cmdParcela_Click()
    Dim Conta As New ContaCorrente
    Dim PosTraco As Byte
    Dim Selecao As String
    Dim Sql As String
    Dim Obrig As New Obrigacao
    Dim Insc As String
    Dim Contador As Integer
    Dim modo As TipoInscricaoObrigacao
    Dim TemDivida  As Boolean
    
        If Not Edita.CriticaCampos(Me) Then Exit Sub
    If txtIm <> "" And txtImovel <> "" Or txtIm = "" And txtImovel = "" Then
        Util.Avisa "Informe a inscrição Municipal ou a Inscrição Cadastral."
        txtIm.SetFocus
        Exit Sub
    End If
    lstNot.Preencher Bdados, ""
    Screen.MousePointer = 11
    PosTraco = InStr(1, cboImposto, "-")
    '
    If txtIm <> "" Then
        modo = etiContribuinte
        Insc = txtIm
    Else
        modo = etiImovel
        Insc = txtImovel
    End If
    If txtIm <> "" Then
        Conta.ExecutaAtualizacao txtIm, etiContribuinte, False, , , txtData
        If Not Obrig.CarregaListaObrigacaoAtualizada(lstNot, txtIm, , txtPeriodoInicial, etlNaoPagosVencidos, , etiContribuinte, IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False)) Then
            Util.Avisa "Consulta sem resultados."
            TemDivida = False
        Else
            TemDivida = True
        End If
    Else
        If Aplicacoes.municipio = "PETROLINA" Then
            Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(50), "TOC_DESCONTO", _
                "TOC_INSCRICAO ='" & txtImovel & "' AND TOC_PERIODO = 2004 AND TOC_TIP_COD_IMPOSTO ='11120200'"
        End If
        Conta.ExecutaAtualizacao txtImovel, etiImovel, , , , txtData
        If Not Obrig.CarregaListaObrigacaoAtualizada(lstNot, txtImovel, , txtPeriodoInicial, etlNaoPagosVencidos, , etiImovel, IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False)) Then
            Util.Avisa "Consulta sem resultados."
            TemDivida = False
        Else
            TemDivida = True
        End If
        If Aplicacoes.municipio = "PETROLINA" Then
            Bdados.AtualizaDados "TAB_OBRIGACAO_CONTRIBUINTE", Bdados.PreparaValor(0), "TOC_DESCONTO", _
                "TOC_INSCRICAO ='" & txtImovel & "' AND TOC_PERIODO = 2004 AND TOC_TIP_COD_IMPOSTO ='11120200'"
        End If
    End If
    
    If TemDivida = True Then
            Sql = Sql & " SELECT TNT_COD_NOTIFICACAO AS Notificação,"
        Sql = Sql & " TNT_INSCRICAO as Inscrição,  TNT_DT_EMISSAO as Emissão, "
        Sql = Sql & " TNT_VENCIMENTO as Vencimento,  TNT_VALOR_NOTIFICACAO as Valor,"
        Sql = Sql & " TNT_TUS_COD_USUARIO As Usuário"
        Sql = Sql & " , TNT_ENTREGA as Entrega"
        Sql = Sql & " From TAB_NOTIFICACAO where 1 = 1  and TNT_TIPO=1"
        If txtIm <> "" Then
            Inscri = txtIm
        Else
            Inscri = txtImovel
        End If
        Sql = Sql & " and TNT_INSCRICAO =  '" & Inscri & "'"
        grdNotifica.Preencher Bdados, Sql
        If Confirma("Notificação gerada com sucesso, deseja imprimir?", "Aviso") = True Then
            Call Imprimir_Notif
        End If
    Else
        Util.Avisa "Débitos não encontrados..."
    End If
    tabNotificacao.Tabs(2).Selected = True
      Screen.MousePointer = 0
      
    Exit Sub
    Selecao = "Select TCC_CODIGO_CONTA as Documento, TCC_INSCRICAO as INSCRICAO, VIN_RAZAO as Contribuinte, tCC_periodo as Periodo, " & _
        Bdados.Converte("tcc_imposto_original + tcc_juros_atual + tcc_multa_atual", TCDuplo) & " as [Debito(R$)] ,tip_sigla_imposto AS Tributo," & _
        " TIP_COD_IMPOSTO, tcc_data_vencimento as Vencimento FROM TAB_IMPOSTO, TAB_CONTA_CONTRIBUINTE, VIS_INSCRICAO "
    Selecao = Selecao & " where tcc_data_vencimento < " & Bdados.Converte(Date, VSClass.TCDataHora) & " and TCC_INSCRICAO= VIN_INSCRICAO and TCC_tip_cod_imposto = TIP_COD_IMPOSTO AND TCC_SALDO_ATUAL > 0 and tcc_status_conta <> 3 "
    If Trim(cboImposto) <> "" Then
        Sql = Sql & " AND TCC_tip_cod_imposto='" & cboImposto.Coluna(0).Valor & "'"
    End If
    If Trim(Insc) <> "" Then Sql = Sql & " and TCC_INSCRICAO='" & Insc & "'"
    If Trim(txtPeriodoInicial) <> "" Then Sql = Sql & " and TCC_PERIODO=" & txtPeriodoInicial
    Selecao = Selecao & Sql
    
    Screen.MousePointer = 0
    CodPagamento = 0
    
End Sub

Private Sub cmdPesquisaIM_Click()
    'blnConsultaIM = True
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIm, txtContrib
    'blnConsultaIM = False
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
    Dim Rs As VSRecordset
    Dim Sql As String
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
        
    Sql = "Select TPT_TEXTO FROM TAB_PARAMETRO_TEXTO WHERE TPT_PARAMETRO = 'CARTA OMISSOS'"
    If Bdados.AbreTabela(Sql, Rs) Then
        txtTexto = "" & Rs!TPT_TEXTO
    End If
    cboImposto.Preencher Bdados, "Select  tip_cod_imposto,TIP_sigla_IMPOSTO  " & Bdados.Concatena & " ' # ' " & Bdados.Concatena & " tip_nome_imposto,tip_nome_imposto From TAB_IMPOSTO order by TIP_sigla_IMPOSTO asc", 1
End Sub


Private Sub lstNot_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid lstNot, ColumnHeader
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
    If Chr(Asc(KeyAscii)) = "/" Then Exit Sub
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

'Private Sub grdNotifica_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Not grdNotifica.SelectedItem Is Nothing Then
'        If Button = 2 Then
'            mnuEmitir.Caption = "Emitir notificação nº " & grdNotifica.SelectedItem
'            mnuEmitir.Tag = grdNotifica.SelectedItem.SubItems(1) & "|" & grdNotifica.SelectedItem.SubItems(2) & "|" & grdNotifica.SelectedItem.SubItems(3)
'            Me.PopupMenu mnuNotifica
'        End If
'    End If
'
'End Sub

Private Sub lstNot_ItemClick(ByVal Item As MSComctlLib.IListItem)
    cboImposto.SetarLinha Imposto.BuscaCodImposto(lstNot.SelectedItem.SubItems(2))
End Sub

Private Sub mnuEmitir_Click()
    Dim Contador As Integer
    Dim Insc        As String
    If txtIm <> "" Then
        Insc = txtIm
    Else
        Insc = txtImovel
    End If
    For Contador = 1 To grdNotifica.ListItems.Count
       
    Next
    ImprimirNotificacao Insc, , Util.ParseString(mnuEmitir.Tag, "|", 3), Util.ParseString(mnuEmitir.Tag, "|", 2), Util.ParseString(mnuEmitir.Caption, "nº", 2), 1
End Sub

Private Sub txtic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
    End If
End Sub

Private Sub TXTINSCRICAO_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    Dim Notific As New Notificacao
    'If Trim(txtInscricao) = "" Then Exit Sub
    grdNotifica.Mensagem = ""
    grdNotifica.Preencher Bdados, ""
    'txtInscricao = BuscaContribuinte(txtInscricao, txtRazao, txtEndereco, Documento)
    'Notific.ExibirNotificacoes grdNotifica, txtInscricao
    Bdados.FechaTabela Rs
End Sub

Private Sub ImprimirNotificacao(Optional Im As String, Optional Imposto As String, Optional Valor As String, Optional Prazo As String, Optional CodPagamento As Double, Optional Destino As Integer, Optional Ic As String, Optional PerInicial As String, Optional PerFinal As String)
    On Error GoTo Trata
    Dim Cobranca As New VSCobranca
    Dim LinhaDigitavel As String
    Dim SelecaoRpt As String
    Dim CodBarra As New CodigoDeBarra
    Dim PicBarra As PictureBox
    Screen.MousePointer = 11
    If Not Rpt.DefinirArquivo(Bdados, App.Path + "\TOmisosPetrolina.rpt") Then Exit Sub
    SelecaoRpt = ""
    If txtIm <> "" Then
        Inscri = txtIm
    Else
        Inscri = txtImovel
    End If
    With Rpt
'        .Formulas "VT_TOTAL", Format(Valor, Const_Monetario)
'        .Formulas "VT_PRAZO", Prazo
        .Selecao = " {TAB_NOTIFICACAO.TNT_COD_NOTIFICACAO} = " & CodPagamento
        .Formulas "VT_EXTRATO", CStr(CodPagamento)
        .Formulas "VT_EMISSAO", Format(Date, "DD/MM/YYYY")
        .Formulas "VT_CLIENTE", txtContrib
        .Formulas "VT_IM", Inscri
        .Formulas "VT_PREFEITURA", Temp.PegaParametro(Bdados, "CLIENTE")
        .Formulas "VT_SECRETARIA", Temp.PegaParametro(Bdados, "SEMFAZ")
        .Formulas "VT_CIDADE", AplicacoesVTFuncoes.municipio & " - " & Temp.PegaParametro(Bdados, "Estado Cliente")
        .Formulas "VT_EnderecoContrib", txtEndereco
        .Formulas "VT_USUARIO", AplicacoesVTFuncoes.Usuario
        If Temp.PegaParametro(Bdados, "PADRAO ARRECADACAO") = "CBR643" Then
            LinhaDigitavel = CodBarra.CriaLinhaDigitavelCBR(Im, Const_Notificacao, CDbl(Valor), Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), PicBarra, txtData, 0, CStr(CodPagamento))
        Else
           LinhaDigitavel = CodBarra.CriaLinhaDigitavel(Im, Const_Notificacao, CDbl(Valor), Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), txtData)
        End If
        .Formulas "LinhaDigitavel", LinhaDigitavel
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            .Formulas "VT_LinhaBarra", CodBarra.LinhaBarraGerada
        End If
        .Formulas "VT_LinhaBarra", CodBarra.LinhaBarraGerada
        .Titulo = "Notificação de Débitos Tributários"
        If Destino = 1 Then
            .Arvore = False
            .Visualizar
        Else
            .Imprimir
        End If

    End With
    Set Rpt = Nothing
    Screen.MousePointer = 0
Exit Sub
    Dim Documento As VSClass.VSDocumento
    
    Set Documento = New VSClass.VSDocumento
        If Documento.Novo(App.Path & "\Modelos\CND.dot") Then
            Documento.textoObjeto "@NumNotificacao", CStr(CodPagamento)
            Documento.textoObjeto "@DataNotificacao", Format(Date, "dd/mm/yyyy")
            Documento.textoObjeto "@IM", Im
            Documento.textoObjeto "@VenctoNotificacao", Prazo
            
            Documento.Substituir "@NumNotificacao", CStr(CodPagamento)
            'Documento.Substituir "@DataNotificacao", Format(Date, "dd/mm/yyyy")
            'Documento.Substituir "@IM", Im
            'Documento.Substituir "@VenctoNotificacao", Prazo
            Documento.Substituir "@Prefeitura", "Prefeitura Municipal de Balsas"
            Documento.Substituir "@Secretaria", "Secretaria Municipal de Fazenda"
            Documento.Substituir "@Departamento", "Departamento de Arrecadação de Tributos"
            Documento.Ativar
        End If
    Set Documento = Nothing
    Screen.MousePointer = 0
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Avisa Err.Description
        Exit Sub
        Resume
        Err.Clear
    End If

End Sub


Private Sub txtim_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtIm_LostFocus()
    Dim Ic As String
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(txtIm) = 10 Or Len(txtIm) = 11 Then
            Ic = Imposto.FormataInscricao(txtIm, InscContrib)
        Else
            Ic = txtIm
        End If
    Else
        Ic = txtIm
    End If
    If Trim(txtIm) <> "" Then
        txtIm = BuscaContribuinte(Ic, txtContrib, txtEndereco)
        If Trim(txtIm) = "" Then
            Avisa "Inscricão não encontrada"
            txtIm.SetFocus
        End If
    End If
    Call PEga_Notificacao
End Sub

Private Sub txtImovel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtImovel_LostFocus()
 Dim Ic As String
  
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtContrib, txtEndereco, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtIm.SetFocus
        End If
        Call PEga_Notificacao
    End If
End Sub
Private Sub PEga_Notificacao()
    Dim Sql As String
    Dim Inscri As String
    
    Sql = Sql & " SELECT TNT_COD_NOTIFICACAO AS Notificação,"
    Sql = Sql & " TNT_INSCRICAO as Inscrição,  TNT_DT_EMISSAO as Emissão, "
    Sql = Sql & " TNT_VENCIMENTO as Vencimento,  TNT_VALOR_NOTIFICACAO as Valor,"
    Sql = Sql & " TNT_TUS_COD_USUARIO As Usuário"
    Sql = Sql & " , TNT_ENTREGA as Entrega"
    Sql = Sql & " From TAB_NOTIFICACAO where 1 = 1 and TNT_TIPO=1"
    If txtIm <> "" Then
        Inscri = txtIm
    Else
        Inscri = txtImovel
    End If
    Sql = Sql & " and TNT_INSCRICAO =  '" & Inscri & "'"
    grdNotifica.Preencher Bdados, Sql
End Sub






Private Sub Imprimir_Notif()
    On Error GoTo Trata
    Dim i As Integer
    Dim ImAnterior As String
    Dim SelecaoRpt As String
    Dim Conta As New ContaCorrente
    Dim Valores As String
    Dim campos As String
    Dim Cobranca As New VSCobranca
    Dim ValorFinal As Double
    Dim InsCad As String
    Dim InsMun As String
    Dim Insc As String
    
    If txtIm <> "" Then
        Insc = txtIm
    Else
        Insc = txtImovel
    End If
    
    If cboDest.ListIndex < 0 Then
        Avisa "Informe destino da(s) notificação(s)."
        cboDest.SetFocus
        Exit Sub
    End If
    If Trim(txtData) = "" Then
        Avisa "Informe a data de vencimento."
        txtData.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = 11
    '1.
    Valores = Bdados.PreparaValor("CARTA OMISSOS", txtTexto)
    campos = "tpt_parametro,TPT_TEXTO"
    Bdados.GravaDados "TAB_PARAMETRO_TEXTO", Valores, campos, "TPT_PARAMETRO = 'CARTA OMISSOS'"
    
    '2.
    If CodPagamento = 0 Then
        CodPagamento = Conta.GeraCodPagamento(EtsNotificacao)
        campos = "TPE_COD_PAGAMENTO_EXTRATO,TPE_TGT_COD_PAGAMENTO,TPE_SUB_VALOR,TPE_TIP_COD_IMPOSTO,TPE_INSCRICAO,TPE_TIPO_DOCUMENTO"
        For i = 1 To lstNot.ListItems.Count
            Valores = Bdados.PreparaValor(CodPagamento, lstNot.ListItems(i).Text, _
                Bdados.Converte(lstNot.ListItems(i).SubItems(10), TCDuplo), lstNot.ListItems(i).SubItems(11), _
                Bdados.Converte(Trim(lstNot.ListItems(i).SubItems(1)), tctexto), 2)
            Bdados.InsereDados "TAB_PAGAMENTO_EXTRATO", Valores, campos
        Next
        If Not (lstNot.SelectedItem Is Nothing) Then
            ValorFinal = Format(lstNot.Colunas(9).Soma, Const_Monetario)
        Else
            If grdNotifica.ListItems.Count > 0 Then ValorFinal = grdNotifica.Colunas(4).Soma
        End If
        
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            If txtIm <> "" Then
                InsMun = Insc
                InsCad = ""
            Else
                InsCad = Insc
                InsMun = Documento
            End If
        Else
            If Len(Insc) > 11 Then
                InsCad = Insc
                InsMun = Documento
            Else
                InsMun = Insc
                InsCad = ""
            End If
        End If
        
        Conta.GeraPagamento InsMun, InsCad, Const_Notificacao, Right(Format(Date, "DD/MM/YYYY"), 4) & Mid(Format(Date, "DD/MM/YYYY"), 4, 2), txtData, CDbl(ValorFinal), 0, 0, CStr(CodPagamento), 0, 0, 0, , EtcCreditoTributario
        '3.
        Valores = Bdados.PreparaValor(CodPagamento, Bdados.Converte(Insc, tctexto), Bdados.Converte(Format(Date, "DD/MM/YYYY"), TCDataHora), Bdados.Converte(Format(txtData, "DD/MM/YYYY"), TCDataHora), ValorFinal, AplicacoesVTFuncoes.Usuario, 1, 1)
        campos = "TNT_COD_NOTIFICACAO,TNT_INSCRICAO,TNT_DT_EMISSAO,TNT_VENCIMENTO,TNT_VALOR_NOTIFICACAO,TNT_TUS_COD_USUARIO,TNT_TIPO_NOTIFICACAO,TNT_TIPO "
        Bdados.InsereDados "TAB_NOTIFICACAO", Valores, campos
    End If
    '4.
    ImprimirNotificacao Insc, Trim(Mid(cboImposto, Edita.PosPic(cboImposto, "-") + 1)), CStr(ValorFinal), txtData, CodPagamento, cboDest.ListIndex, ""
    ImAnterior = ""
    CodPagamento = 0
    Exit Sub
Trata:
    If Err.Number <> 0 Then
        Avisa Err.Description
        Err.Clear
    End If
End Sub
