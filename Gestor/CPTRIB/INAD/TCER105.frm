VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TCER105 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCER105"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
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
      TabIndex        =   5
      Top             =   1845
      Width           =   8130
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   1138
      Icone           =   "TCER105.frx":0000
   End
   Begin VTOcx.cboVISUAL cboTributo 
      Height          =   315
      Left            =   495
      TabIndex        =   0
      Top             =   690
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   556
      Caption         =   "Tributo"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin ActiveTabs.SSActiveTabs tabCND 
      Height          =   3690
      Left            =   120
      TabIndex        =   16
      Top             =   3180
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   6509
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
      Tabs            =   "TCER105.frx":031A
      Images          =   "TCER105.frx":03CB
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3270
         Index           =   0
         Left            =   -99969
         TabIndex        =   17
         Top             =   30
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "TCER105.frx":1064
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
            Height          =   3180
            Left            =   45
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   30
            Width           =   9270
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3270
         Index           =   1
         Left            =   30
         TabIndex        =   19
         Top             =   30
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "TCER105.frx":108C
         Begin VTOcx.grdVISUAL grdDebitosVencido 
            Height          =   3210
            Left            =   90
            TabIndex        =   23
            Top             =   60
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   5662
            CorBorda        =   32768
            Caption         =   "Cr�ditos vencidos"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3270
         Left            =   -99969
         TabIndex        =   20
         Top             =   30
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "TCER105.frx":10B4
         Begin VTOcx.grdVISUAL grdCPND 
            Height          =   3210
            Left            =   90
            TabIndex        =   21
            Top             =   60
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   5662
            CorBorda        =   32768
            Caption         =   "Certid�es emitidas"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
      End
   End
   Begin VTOcx.txtVISUAL txtValidade 
      Height          =   300
      Left            =   7320
      TabIndex        =   7
      Tag             =   "Validade"
      Top             =   2250
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   529
      Caption         =   "Validade"
      Text            =   ""
      Enabled         =   0   'False
      Formato         =   0
      Restricao       =   2
      MaxLen          =   10
   End
   Begin VTOcx.txtVISUAL txtFinalidade 
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Tag             =   "Finalidade"
      Top             =   2250
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   529
      Caption         =   "Finalidade"
      Text            =   ""
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   6900
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL CmdImprimir 
         Height          =   375
         Left            =   5280
         TabIndex        =   10
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   13
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
      TabIndex        =   11
      Top             =   3240
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
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      Caption         =   "at�"
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
      Top             =   1095
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   529
      Caption         =   "Inscric�o"
      Text            =   ""
      Restricao       =   2
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtRazao 
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   1485
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   529
      Caption         =   "Nome/Raz�o"
      Text            =   ""
      Enabled         =   0   'False
      Requerido       =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
      Height          =   315
      Left            =   2865
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1080
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   ""
      Acao            =   5
   End
   Begin VTOcx.txtVISUAL txtImovel 
      Height          =   300
      Left            =   3255
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   529
      Caption         =   "Cadastro do Im�vel"
      Text            =   ""
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.cmdVISUAL cmdVISUAL1 
      Height          =   315
      Left            =   6555
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1080
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   ""
      Acao            =   5
   End
   Begin VTOcx.txtVISUAL txtCertidao 
      Height          =   300
      Left            =   6990
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      Caption         =   "No. Certid�o"
      Text            =   ""
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtObs 
      Height          =   510
      Left            =   90
      TabIndex        =   8
      Top             =   2580
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   900
      Caption         =   "Observac�o"
      Text            =   ""
   End
   Begin VTOcx.cmdVISUAL cmdBuscarContrib 
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   2640
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   661
      Caption         =   "&Buscar"
      Acao            =   5
      CorBorda        =   16711680
      CorFrente       =   0
      CorFundo        =   16777152
   End
End
Attribute VB_Name = "TCER105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Certidao As iCertidao
Dim Conta As ContaCorrente
Dim Obrig As New Obrigacao
Dim CodCertidao As String
Public TipoCertidao As String
Dim InscricaoCad As String, InscricaoMun As String

'   TCER101 CND
'   TCER102 CPD
'   TCER103 CPND
Public Sub BuscaServidor()
    
End Sub
Sub ImprimeCertidao(CodCertidao As String)
    'Dim Documento As VSDoc.VSDocumento
    Dim RELAT As VSRelatorio
    Dim Im As String, Ic As String
    Dim Relatorio As String
    Dim Filtro As String
    Dim Sql As String
    Set RELAT = New VSRelatorio
    
'    If Temp.PegaParametro(Bdados, "MODELO CERTIDAO") = "2" Then
'        Sql = "SELECT * FROM VIS_CERTIDAO_NEGATIVA"
'        Sql = Sql & " WHERE TCN_COD_NEGATIVA = " & CodCertidao
'        VisualizarActiveReport AR_CND, Bdados, Sql
'        Exit Sub
'    End If
    
    With RELAT
        Select Case TipoCertidao
            Case "TCER101": Relatorio = "\TCN.rpt"
            Case "TCER102": Relatorio = "\TCN.rpt"
            Case "TCER103": Relatorio = "\TCN.rpt"
            Case "TCER104": Relatorio = "\TCN.rpt"
        End Select
        
       Dim strTipo As String
     If txtImovel = "" Then
        strTipo = " MOBILI�RIOS" 'DEPOIS VOLTAR PARA MOBILIARIO
    Else
        strTipo = " IMOBILI�RIOS"
    End If
    

        If Not .DefinirArquivo(Bdados, App.Path + "\TCN.rpt") Then Exit Sub
            If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "GAF")
            Else
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            End If
            .Selecao = "{TAB_CERTIDAO_NEGATIVA.TCN_COD_NEGATIVA} = " & CodCertidao
            .Titulo = "Certid�o Negativa de D�bitos"
            .Formulas "VT_CIDADE", AplicacoesVTFuncoes.municipio
            Select Case TipoCertidao
                Case "TCER101"
                    .Formulas "VT_TITULO", "CND - CERTID�O NEGATIVA DE D�BITOS" & strTipo
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, "")
                    .Formulas "CONTRIBUINTE", txtRazao
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
                    
                    If cboTributo.ListIndex >= 0 Then
                        .Formulas "TRIBUTO1", cboTributo.Text
                    End If
                    .SubRelatorio = "TextoCert"
                    .Selecao = "{TAB_PARAMETRO_TEXTO.TPT_PARAMETRO} = 'Certidao Negativa'"
                Case "TCER102":
                    .Formulas "VT_TITULO", "CPD - CERTID�O POSITIVA DE D�BITO" & strTipo
                    
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, "")
                    .SubRelatorio = "TextoCert"
                    .Selecao = "{TAB_PARAMETRO_TEXTO.TPT_PARAMETRO} = 'CPD'"
                Case "TCER103":
                    .Formulas "VT_TITULO", "CPND - CERTID�O POSITIVA COM EFEITO DE NEGATIVA DE D�BITO" & strTipo
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, "")
                    .SubRelatorio = "TextoCert"
                    .Selecao = "{TAB_PARAMETRO_TEXTO.TPT_PARAMETRO} = 'CPND'"
                Case "TCER104"
                    .Formulas "VT_TITULO", "CND - CERTID�O NEGATIVA DE D�BITOS" & strTipo
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, "")
                    .SubRelatorio = "TextoCert"
                    .Selecao = "{TAB_PARAMETRO_TEXTO.TPT_PARAMETRO} = 'CNE'"
            End Select
            
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
    
 '   If Trim(txtIm) = "" Then: Util.Informa "Informe a inscri��o.": Exit Sub
            
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
             If Obrig.MostraObrigacaoGerada(grdDebitosVencido, CStr(cboTributo.Coluna(0).Valor), txtIm, _
                 , , , , , , , txtImovel, etlNaoPagosVencidos) Then
                If Not Confirma("Existem d�bitos. O servidor " & AplicacoesVTFuncoes.Usuario & " responsabiliza-se pela liberac�o?") Then
                    cmdEmitir.Enabled = False
                    Exit Sub
                Else
                    If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
                    tabCND.Tabs(2).Selected = True
                End If
            Else
                cmdEmitir.Enabled = True
                If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
                tabCND.Tabs(2).Selected = True
            End If
        Case "TCER102": Tipo = tcCPD
            Certidao.BuscarCertidoes grdCPND, Tipo, txtIm
            If Obrig.CarregaListaObrigacao(grdDebitosVencido, txtIm, CStr(cboTributo.Coluna(0).Valor), , etlNaoPagosNaoVencidos) = False Then
                Util.Informa "N�o � poss�vel emitir a CPD. N�o existem cr�ditos vencidos."
                cmdEmitir.Enabled = False
            Else
                cmdEmitir.Enabled = True
                If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
                tabCND.Tabs(2).Selected = True
            End If
        Case "TCER103": Tipo = tcCPND
            Certidao.BuscarCertidoes grdCPND, Tipo, txtIm
            If Obrig.CarregaListaObrigacao(grdDebitosVencido, txtIm, CStr(cboTributo.Coluna(0).Valor), , etlNaoPagos) = False Then
                Util.Informa "N�o � poss�vel emitir a CPND. N�o existem cr�ditos n�o vencidos."
                cmdEmitir.Enabled = False
            Else
                cmdEmitir.Enabled = True
                If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
            End If
        Case "TCER104": Tipo = tcCNE
            Certidao.BuscarCertidoes grdCPND, Tipo, txtIm, ""
            If Obrig.CarregaListaObrigacao(grdDebitosVencido, txtIm, CStr(cboTributo.Coluna(0).Valor), , etlNaoPagos) = False Then
                Util.Informa "N�o � poss�vel emitir a CNE. N�o existem debitos."
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
    If Not Util.Confirma("Confirma a emiss�o da certid�o") Then Exit Sub
    If AplicacoesVTFuncoes.municipio <> "PETROLINA" Then
        Select Case TipoCertidao
            Case "TCER101": If grdDebitosVencido.ListItems.Count = 0 Then Exit Sub
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
        .Im = InscricaoMun
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
    Select Case TipoCertidao
        Case "TCER101": txtTexto = Certidao.BuscaTexto("CERTIDAO NEGATIVA")
        Case "TCER102": txtTexto = Certidao.BuscaTexto("CPD")
        Case "TCER103": txtTexto = Certidao.BuscaTexto("CPND")
        
    End Select

    'cboTributo.SetFocus
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
        Case "TCER101": txtTexto = Certidao.BuscaTexto("Certidao Negativa"): grdDebitosVencido.Caption = "D�bitos"
        Case "TCER102": txtTexto = Certidao.BuscaTexto("CPD"): grdDebitosVencido.Caption = "D�bitos vencidos"
        Case "TCER103": txtTexto = Certidao.BuscaTexto("CPND"): grdDebitosVencido.Caption = "D�bitos n�o vencidos"
        Case "TCER104": txtTexto = Certidao.BuscaTexto("CNE"): grdDebitosVencido.Caption = "D�bitos vencidos"
    End Select
    DoEvents
End Sub
Private Sub Form_Load()
'    TipoCertidao = "TCER103"
    
    
    Set Certidao = New iCertidao
    Set Conta = New ContaCorrente
    TipoCertidao = "TCER101"
    
    Certidao.PreencherCboImposto cboTributo
    
    
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
    
    If Trim(txtImovel) = "" Then Exit Sub
    txtImovel.AgruparValores = False
    txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, Doc, etiImovel)
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

