VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TCER101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCER101"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEndereco 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1140
      TabIndex        =   23
      Top             =   3090
      Width           =   8130
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   1138
      Icone           =   "TCER101.frx":0000
   End
   Begin VTOcx.cboVISUAL cboTributo 
      Height          =   315
      Left            =   255
      TabIndex        =   0
      Top             =   690
      Visible         =   0   'False
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   556
      Caption         =   "Tributo"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin ActiveTabs.SSActiveTabs tabCND 
      Height          =   2370
      Left            =   120
      TabIndex        =   13
      Top             =   4500
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   4180
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
      Tabs            =   "TCER101.frx":031A
      Images          =   "TCER101.frx":03CB
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   1950
         Index           =   0
         Left            =   -99969
         TabIndex        =   14
         Top             =   30
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   3440
         _Version        =   131082
         TabGuid         =   "TCER101.frx":1064
         Begin VB.TextBox txtTexto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1620
            Left            =   45
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   150
            Width           =   9270
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   1950
         Index           =   1
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   3440
         _Version        =   131082
         TabGuid         =   "TCER101.frx":108C
         Begin VTOcx.grdVISUAL grdDebitosVencido 
            Height          =   2010
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   3545
            CorBorda        =   32768
            Caption         =   "Créditos vencidos"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   1950
         Left            =   -99969
         TabIndex        =   17
         Top             =   30
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   3440
         _Version        =   131082
         TabGuid         =   "TCER101.frx":10B4
         Begin VTOcx.grdVISUAL grdCPND 
            Height          =   2010
            Left            =   90
            TabIndex        =   18
            Top             =   60
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   3545
            CorBorda        =   32768
            Caption         =   "Certidões emitidas"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
      End
   End
   Begin VTOcx.txtVISUAL txtValidade 
      Height          =   300
      Left            =   7320
      TabIndex        =   3
      Tag             =   "Validade"
      Top             =   3450
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   529
      Caption         =   "Validade"
      Text            =   ""
      Formato         =   0
      Restricao       =   2
      MaxLen          =   10
   End
   Begin VTOcx.txtVISUAL txtFinalidade 
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Tag             =   "Finalidade"
      Top             =   3450
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   529
      Caption         =   "Finalidade"
      Text            =   ""
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   6885
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL CmdImprimir 
         Height          =   375
         Left            =   5280
         TabIndex        =   21
         Top             =   75
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   661
         Caption         =   "Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7680
         TabIndex        =   11
         Top             =   75
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8625
         TabIndex        =   12
         Top             =   75
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdEmitir 
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Top             =   75
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   661
         Caption         =   "&Emitir"
         Acao            =   4
         Enabled         =   0   'False
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.txtVISUAL txtRefInicio 
      Height          =   300
      Left            =   690
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   529
      Caption         =   "Periodo"
      Text            =   ""
      Enabled         =   0   'False
      Restricao       =   2
      MaxLen          =   7
      MinLen          =   4
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtRefFim 
      Height          =   300
      Left            =   2190
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      Caption         =   "até"
      Text            =   ""
      Enabled         =   0   'False
      Restricao       =   2
      MaxLen          =   7
      MinLen          =   4
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtIm 
      Height          =   300
      Left            =   330
      TabIndex        =   1
      Top             =   2415
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   529
      Caption         =   "Inscricão"
      Text            =   ""
      Restricao       =   2
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtRazao 
      Height          =   300
      Left            =   0
      TabIndex        =   24
      Top             =   2750
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   529
      Caption         =   "Nome/Razão"
      Text            =   ""
      Enabled         =   0   'False
      Requerido       =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
      Height          =   315
      Left            =   2625
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2400
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   ""
      Acao            =   5
   End
   Begin VTOcx.txtVISUAL txtImovel 
      Height          =   300
      Left            =   3015
      TabIndex        =   5
      Top             =   2400
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   529
      Caption         =   "Cadastro do Imóvel"
      Text            =   ""
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.cmdVISUAL cmdVISUAL1 
      Height          =   315
      Left            =   6795
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2400
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   ""
      Acao            =   5
   End
   Begin VTOcx.txtVISUAL txtCertidao 
      Height          =   300
      Left            =   7230
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      Caption         =   "No. Certidão"
      Text            =   ""
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtObs 
      Height          =   570
      Left            =   45
      TabIndex        =   7
      Top             =   3830
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   1005
      Caption         =   "Observacão"
      Text            =   ""
   End
   Begin VTOcx.cmdVISUAL cmdBuscarContrib 
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   3840
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   661
      Caption         =   "&Buscar"
      Acao            =   5
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.grdVISUAL grdTributos 
      Height          =   1890
      Left            =   120
      TabIndex        =   27
      Top             =   720
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   3334
      CorBorda        =   32768
      CabecalhoTitulo =   ""
      Caption         =   "TRIBUTOS"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      CheckBox        =   -1  'True
   End
End
Attribute VB_Name = "TCER101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Certidao As iCertidao
Dim Conta As ContaCorrente
Dim Obrig As New Obrigacao
Dim CodCertidao As String
Dim dataValidade As String
Public TipoCertidao As String
Dim InscricaoCad As String, InscricaoMun As String

'   TCER101 CND
'   TCER102 CPD
'   TCER103 CPND

Sub ImprimeCertidao(CodCertidao As String)
    'Dim Documento As VSDoc.VSDocumento
    Dim RELAT As VSRelatorio
    Dim Im As String, Ic As String
    Dim Relatorio As String
    Dim Filtro As String
    Dim Sql As String
    Set RELAT = New VSRelatorio
    Dim marcou As Boolean
    marcou = False
    Dim i As Integer, selecionado As Integer
    selecionado = 1
    For i = 1 To grdTributos.ListItems.Count
        If grdTributos.ListItems(i).Checked Then
            marcou = True
            Exit For
        End If
    Next
    If marcou = False Then
        Avisa "Selecione um TRIBUTO para a emissão da certidão"
        Exit Sub
    End If
'    If Temp.PegaParametro(Bdados, "MODELO CERTIDAO") = "2" Then
'        Sql = "SELECT * FROM VIS_CERTIDAO_NEGATIVA"
'        Sql = Sql & " WHERE TCN_COD_NEGATIVA = " & CodCertidao
'        VisualizarActiveReport AR_CND, Bdados, Sql
'        Exit Sub
'    End If
    
    
    Dim strTipo As String
   
    If txtImovel = "" Then
        strTipo = " MOBILIÁRIOS"
    Else
        strTipo = " IMOBILIÁRIOS"
    End If
    
    With RELAT
        Select Case TipoCertidao
            Case "TCER101": Relatorio = "\TCN.rpt"
            Case "TCER102": Relatorio = "\TCN.rpt"
            Case "TCER103": Relatorio = "\TCN.rpt"
            Case "TCER104": Relatorio = "\TCN.rpt"
        End Select

        If Not .DefinirArquivo(Bdados, App.Path + "\TCN.rpt") Then Exit Sub
            If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "GAF")
            Else
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            End If
            .Selecao = "{TAB_CERTIDAO_NEGATIVA.TCN_COD_NEGATIVA} = " & CodCertidao
            .Titulo = "Certidão Negativa de Débitos"
            .Formulas "VT_CIDADE", AplicacoesVTFuncoes.municipio
            Select Case TipoCertidao
                Case "TCER101"
                    .Formulas "VT_TITULO", "CND - CERTIDÃO NEGATIVA DE DÉBITOS" & strTipo
                    
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, "")
                    Dim Doc As String
                    
                    If txtImovel = "" Then
                        .Formulas "CONTRIBUINTE", txtRazao
                        .Formulas "IM", txtIm
                        .Formulas "DOC", PegaDoc(txtIm)
                    Else
                        .Formulas "CONTRIBUINTE", txtIm & " - " & txtRazao
                        .Formulas "IM", txtImovel
                        .Formulas "DOC", PegaDoc(txtIm)
                    End If
                    .Formulas "ATIVIDADE", Imposto.BuscaNomeCAE(CodAtividade(txtIm))
                    .Formulas "VT_Endereco_Contrib", txtEndereco
                    
                    'If cboTributo.ListIndex >= 0 Then
                      '  .Formulas "TRIBUTO", cboTributo.Text
                    'End If
                    For i = 1 To grdTributos.ListItems.Count ' TOTAL DE TRIBUTOS
                        If selecionado <= 3 Then
                            If grdTributos.ListItems(i).Checked = True Then ' SE FOI MARCADO PARA IMPRESSA
                                .Formulas "TRIBUTO" & selecionado, grdTributos.ListItems(i)
                                selecionado = selecionado + 1
                            End If
                        End If
                    Next i
                    .SubRelatorio = "TextoCert"
                    .Selecao = "{TAB_PARAMETRO_TEXTO.TPT_PARAMETRO} = 'Certidao Negativa'"
                Case "TCER102":
                    .Formulas "VT_TITULO", "CPD - CERTIDÃO POSITIVA DE DÉBITO" & strTipo
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, "")
                    .SubRelatorio = "TextoCert"
                    .Selecao = "{TAB_PARAMETRO_TEXTO.TPT_PARAMETRO} = 'CPD'"
                Case "TCER103":
                    .Formulas "VT_TITULO", "CPND - CERTIDÃO POSITIVA COM EFEITO DE NEGATIVA DE DÉBITO" & strTipo
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, "")
                    .SubRelatorio = "TextoCert"
                    .Selecao = "{TAB_PARAMETRO_TEXTO.TPT_PARAMETRO} = 'CPND'"
                Case "TCER104"
                    .Formulas "VT_TITULO", "CND - CERTIDÃO NEGATIVA DE DÉBITOS" & strTipo
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, "")
                    .SubRelatorio = "TextoCert"
                    .Selecao = "{TAB_PARAMETRO_TEXTO.TPT_PARAMETRO} = 'CNE'"
            End Select
           '.Formulas "OB", txtObs
            .SubRelatorio = ""
            .Selecao = "{TAB_CERTIDAO_NEGATIVA.TCN_COD_NEGATIVA} = " & CodCertidao
            .Arvore = False
            .CopiasDetalhes = 2
            .Visualizar
    End With
    Set RELAT = Nothing
End Sub




Private Sub cmdBuscarContrib_Click()
    Dim Tipo As TipoCertidao
    Dim Obrig As New Obrigacao
    
    If txtIm = "" Then  'Or txtIm <> "" And txtImovel <> "" Then
        Util.Avisa "Certidao de IPTU, informe o IMOVEL e PROPRIETARIO, para os demais tributos, informe o CONTRIBUINTE."
        txtIm.SetFocus
        Exit Sub
    End If
            
    'busca certidoes ja emitidas
'    If Len(Trim(txtIm)) = 10 Then
'        InscricaoMun = txtIm: InscricaoCad = ""
'    Else
'        InscricaoCad = txtIm: InscricaoMun =
'    End If
    If Temp.PegaParametro(Bdados, "VERIFICA NOTIFICACAO") = "SIM" Then
        If Not Obrig.LiberaContribuinteNotificado(IIf(Trim(InscricaoCad) = "", InscricaoMun, InscricaoCad)) Then
            CmdImprimir.Enabled = False
            Exit Sub
        Else
            CmdImprimir.Enabled = True
        End If
    End If
    Select Case TipoCertidao
        Case "TCER101": Tipo = tcCND
            If (txtCertidao <> "") Then
                Certidao.BuscarCertidoes grdCPND, Tipo, , , txtCertidao
                tabCND.Tabs(2).Selected = True
                Exit Sub
            End If
            Certidao.BuscarCertidoes grdCPND, Tipo, InscricaoMun, InscricaoCad
            'busca debitos abertos nao vencidos
            Dim trib As String
            
             If Obrig.MostraObrigacaoGerada(grdDebitosVencido, CStr(grdTributos.SelectedItem.ListSubItems(1)), txtIm, _
                 , , , , , , , txtImovel, etlNaoPagosVencidos, IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False)) Then
                Avisa "Existem débitos pendentes para este contribuinte. Certidão não pode ser liberada."
                cmdEmitir.Enabled = False
                Exit Sub
            Else
                cmdEmitir.Enabled = True
                If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
                tabCND.Tabs(2).Selected = True
            End If
        Case "TCER102": Tipo = tcCPD
            Certidao.BuscarCertidoes grdCPND, Tipo, txtIm
            If Obrig.CarregaListaObrigacao(grdDebitosVencido, txtIm, CStr(cboTributo.Coluna(0).Valor), , etlNaoPagosNaoVencidos) = False Then
                Util.Informa "Não é possível emitir a CPD. Não existem créditos vencidos."
                cmdEmitir.Enabled = False
            Else
                cmdEmitir.Enabled = True
                If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
                tabCND.Tabs(2).Selected = True
            End If
        Case "TCER103": Tipo = tcCPND
            Certidao.BuscarCertidoes grdCPND, Tipo, txtIm
            If Obrig.CarregaListaObrigacao(grdDebitosVencido, txtIm, CStr(cboTributo.Coluna(0).Valor), , etlNaoPagos) = False Then
                Util.Informa "Não é possível emitir a CPND. Não existem créditos não vencidos."
                cmdEmitir.Enabled = False
            Else
                cmdEmitir.Enabled = True
                If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
            End If
        Case "TCER104": Tipo = tcCNE
            Certidao.BuscarCertidoes grdCPND, Tipo, txtIm, ""
            If Obrig.CarregaListaObrigacao(grdDebitosVencido, txtIm, CStr(cboTributo.Coluna(0).Valor), , etlNaoPagos) = False Then
                Util.Informa "Não é possível emitir a CNE. Não existem debitos."
                cmdEmitir.Enabled = False
            Else
                cmdEmitir.Enabled = True
                If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
            End If
        End Select
End Sub

Private Sub cmdEmitir_Click()
    'TCER103 = Negativa
    'TCER102 = Positiva
    'TCER101 = Positiva/Negativa
    If Not Util.Confirma("Confirma a emissão da certidão") Then Exit Sub
    If AplicacoesVTFuncoes.municipio <> "PETROLINA" Then
        Select Case TipoCertidao
            Case "TCER101": If grdDebitosVencido.ListItems.Count > 0 Then Exit Sub
            Case "TCER102": If grdDebitosVencido.ListItems.Count = 0 Then Exit Sub
            Case "TCER103": If grdDebitosVencido.ListItems.Count = 0 Then Exit Sub
            Case "TCER104": If grdDebitosVencido.ListItems.Count = 0 Then Exit Sub
        End Select
    End If
    If Not CriticaCampos(Me) Then Exit Sub
    With Certidao
        Select Case TipoCertidao
            Case "TCER101"
                .GravarTexto "CERTIDAO NEGATIVA", txtTexto
                .Tipo = tcCND
            Case "TCER102"
                .GravarTexto "CPD", txtTexto
                .Tipo = tcCPD
            Case "TCER103"
                .GravarTexto "CPND", txtTexto
                .Tipo = tcCPND
        End Select
        
        CodCertidao = Conta.GeraCodPagamento("37")
        .CodNegativa = CodCertidao
        .Im = txtIm 'ANTES=InscricaoMun GLEYSON
        If Trim(InscricaoCad) <> "" Then .Ic = Bdados.Converte(txtImovel, tctexto)
        .DataNegativa = Format(Date, "DD/MM/YYYY")
        .Finalidade = txtFinalidade
        .Validade = txtValidade
        .PeriodoInicial = Edita.TiraPic(IIf(txtRefInicio <> "", txtRefInicio, Format(Date, "mm/yyyy")), "/")
        .PeriodoFinal = Edita.TiraPic(txtRefFim, "/")
        .Observacao = txtObs
        If cboTributo.ListIndex >= 0 Then
            .Imposto = cboTributo.Coluna(0).Valor
        Else
            .Imposto = ""
        End If
        
        .CodUsuario = AplicacoesVTFuncoes.Usuario
        If .Gravar Then
            ImprimeCertidao CodCertidao
        Select Case TipoCertidao
            Case "TCER101": Util.Informa "CND Emitida para " & txtIm
            Case "TCER102": Util.Informa "CPD Emitida para " & txtIm
            Case "TCER103": Util.Informa "CPND Emitida para " & txtIm
            Case "TCER104": Util.Informa "CND Emitida para " & txtIm
        End Select
'            Call cmdLimpar_Click
        End If
    End With
End Sub

Private Sub cmdImprimir_Click()
    If grdCPND.ListItems.Count >= 1 Then
        ImprimeCertidao grdCPND.SelectedItem
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtIm.Enabled = True
    tabCND.Tabs(1).Selected = True
    grdCPND.ListItems.Clear
    grdDebitosVencido.ListItems.Clear
    InscricaoCad = ""
    InscricaoMun = ""
    txtValidade = DateAdd("d", Nvl(Temp.PegaParametro(Bdados, "VALIDADE CERTIDAO"), 0), Date)
    Select Case TipoCertidao
        Case "TCER101": txtTexto = Certidao.BuscaTexto("CERTIDAO NEGATIVA")
        Case "TCER102": txtTexto = Certidao.BuscaTexto("CPD")
        Case "TCER103": txtTexto = Certidao.BuscaTexto("CPND")
    End Select

    Datevalidade
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Activate()
    cabVISUAL1.Exibir Bdados, TipoCertidao, App.Path
    rodVISUAL1.Exibir Bdados, TipoCertidao, App.Major, App.Minor, App.Revision
    txtValidade = DateAdd("d", Nvl(Temp.PegaParametro(Bdados, "VALIDADE CERTIDAO"), 0), Date)
    Select Case TipoCertidao
        Case "TCER101": txtTexto = Certidao.BuscaTexto("Certidao Negativa"): grdDebitosVencido.Caption = "Débitos"
        Case "TCER102": txtTexto = Certidao.BuscaTexto("CPD"): grdDebitosVencido.Caption = "Débitos vencidos"
        Case "TCER103": txtTexto = Certidao.BuscaTexto("CPND"): grdDebitosVencido.Caption = "Débitos não vencidos"
        Case "TCER104": txtTexto = Certidao.BuscaTexto("CNE"): grdDebitosVencido.Caption = "Débitos vencidos"
    End Select
    DoEvents
End Sub


Private Sub Form_Load()
'    TipoCertidao = "TCER103"
    
    
    Set Certidao = New iCertidao
    Set Conta = New ContaCorrente
    TipoCertidao = "TCER101"
    
    Certidao.PreencherCboImposto cboTributo
    'BCP
    'Dim i As Integer
    'grdTributos.ColumnHeaders.Add , , "", 8500
    'grdTributos.ColumnHeaders.Add , , "", 100
    
    'For i = 1 To cboTributo.ListCount
     '   grdTributos.ListItems.Add , , cboTributo.List(i)
    'Next i
    Certidao.PreencherGridImposto grdTributos
    cmdEmitir.Enabled = False
    Datevalidade
    cboTributo.ListIndex = -1
    
    '
End Sub

Private Sub Datevalidade()
    Dim DataValid As String
    Dim Mes As String
    Dim Ano As String
    Dim Dia As String
    
    DataValid = Format(Date, "dd/mm/yyyy")
    Dia = Mid(DataValid, 1, 2)
    Mes = Mid(DataValid, 4, 2)
    Ano = Mid(DataValid, 7, 4)
    Mes = Mes + 2
     If CInt(Mes) > 12 Then
        Mes = Mes - 12
        Ano = Ano + 1
        Mes = "0" & Mes
    End If
    txtValidade = Dia & "/" & Mes & "/" & Ano
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Certidao = Nothing
    Set Conta = Nothing
End Sub





Private Sub grdCPND_DblClick()
    txtObs = grdCPND.ListItems(grdCPND.SelectedItem.Index).SubItems(6)
End Sub

Private Sub grdTributos_Click()
    If cboTributo.Text = "" Then
        cboTributo.Text = grdTributos.SelectedItem
    End If
End Sub

Private Sub txtIm_LostFocus()

    Dim RetNome As String
    Dim Doc As String
    
    If Trim(txtIm) = "" Then Exit Sub
    txtIm.AgruparValores = False
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, Doc, etiContribuinte)
    InscricaoMun = txtIm: InscricaoCad = ""
    Call BuscaCertidao
    
End Sub
Private Sub BuscaCertidao()
    Certidao.BuscarCertidoes grdCPND, tcCND, txtIm, txtImovel
End Sub


Private Sub txtImovel_LostFocus()
    Dim RetNome As String
    Dim Doc As String
    Dim ri As VSRecordset
    
    If Trim(txtImovel) = "" Then Exit Sub
    txtImovel.AgruparValores = False
    txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, Doc, etiImovel)
    'BCP
        If Len(txtImovel) > 0 Then
            If Bdados.AbreTabela("SELECT TIM_TCI_IM FROM TAB_IMOVEL WHERE TIM_IC='" & txtImovel & "'", ri) Then
                txtIm = IIf(IsNull(ri(0)), "", ri(0))
            Else
                txtIm = ""
            End If
            
        End If
    'FIM
    InscricaoMun = "": InscricaoCad = txtImovel
    Call BuscaCertidao
End Sub

Private Sub txtRefFim_GotFocus()
    txtRefFim = Edita.TiraPic(txtRefFim, "/")
End Sub

Private Sub txtRefFim_LostFocus()
    If Len(txtRefFim) <> 4 Then
        txtRefFim = Format(txtRefFim, "00/0000")
        If Not IsDate(txtRefFim) Then txtRefFim = ""
    End If
End Sub

Private Sub txtRefInicio_GotFocus()
    txtRefInicio = Edita.TiraPic(txtRefInicio, "/")
End Sub

Private Sub txtRefInicio_LostFocus()
    If Len(txtRefInicio) <> 4 Then
        txtRefInicio = Format(txtRefInicio, "00/0000")
        If Not IsDate(txtRefInicio) Then txtRefInicio = ""
    End If
End Sub
Private Function PegaDoc(Contribuinte As String) As String
    Dim Sql As String
    Sql = "Select tci_cgc_cpf from tab_contribuinte where tci_im = '" & Contribuinte & "'"
    If Bdados.AbreTabela(Sql) Then
        PegaDoc = "" & Bdados.Tabela("tci_cgc_cpf")
    End If
End Function

Private Function CodAtividade(Contribuinte As String) As String
    Dim Sql As String
    Sql = "Select tci_tae_cae from tab_contribuinte where tci_im = '" & Contribuinte & "'"
    If Bdados.AbreTabela(Sql) Then
        CodAtividade = "" & Bdados.Tabela("tci_tae_cae")
    End If
End Function

