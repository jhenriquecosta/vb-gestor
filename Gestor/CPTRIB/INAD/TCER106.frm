VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "CABECALHO.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "SSTABS2.OCX"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#3.1#0"; "VTCONTROLES.OCX"
Begin VB.Form TCER106 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.txtVISUAL txtContrib 
      Height          =   285
      Left            =   3315
      TabIndex        =   18
      Top             =   1065
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   503
      Caption         =   ""
      Text            =   ""
      Enabled         =   0   'False
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtIm 
      Height          =   285
      Left            =   555
      TabIndex        =   1
      Tag             =   "Insc. Municipal"
      Top             =   1065
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   503
      Caption         =   "Inscricao"
      Text            =   ""
      Restricao       =   2
   End
   Begin VTOcx.cboVISUAL cboTributo 
      Height          =   315
      Left            =   735
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
      TabIndex        =   10
      Top             =   2100
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
      TagVariant      =   ""
      Tabs            =   "TCER106.frx":0000
      Images          =   "TCER106.frx":00B8
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3270
         Index           =   0
         Left            =   -99969
         TabIndex        =   11
         Top             =   30
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "TCER106.frx":0D54
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
            Height          =   2835
            Left            =   45
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   30
            Width           =   9270
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   3270
         Index           =   1
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "TCER106.frx":0D7C
         Begin VTOcx.grdVISUAL grdDebitosVencido 
            Height          =   3180
            Left            =   90
            TabIndex        =   20
            Top             =   60
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   5609
            CorBorda        =   32768
            Caption         =   "Créditos vencidos"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   3270
         Left            =   -99969
         TabIndex        =   14
         Top             =   30
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "TCER106.frx":0DA4
         Begin VTOcx.grdVISUAL grdCPND 
            Height          =   2865
            Left            =   30
            TabIndex        =   15
            Top             =   45
            Width           =   9270
            _ExtentX        =   16351
            _ExtentY        =   5054
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
      Left            =   7560
      TabIndex        =   3
      Tag             =   "Validade"
      Top             =   1740
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
      Left            =   480
      TabIndex        =   2
      Tag             =   "Finalidade"
      Top             =   1740
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
      TabIndex        =   16
      Top             =   5835
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL cmdBuscarContrib 
         Height          =   375
         Left            =   5745
         TabIndex        =   6
         Top             =   75
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7665
         TabIndex        =   8
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
         TabIndex        =   9
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
         Left            =   6705
         TabIndex        =   7
         Top             =   75
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Emitir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   1138
      Icone           =   "TCER106.frx":0DCC
   End
   Begin VTOcx.txtVISUAL txtRefInicio 
      Height          =   300
      Left            =   690
      TabIndex        =   4
      Top             =   2700
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
      TabIndex        =   5
      Top             =   2700
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
   Begin VTOcx.txtVISUAL txtEndereco 
      Height          =   285
      Left            =   1380
      TabIndex        =   19
      Top             =   1395
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   503
      Caption         =   ""
      Text            =   ""
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "TCER106"
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

Sub ImprimeCertidao(CodCertidao As String)
    'Dim Documento As VSDoc.VSDocumento
    Dim RELAT As VSRelatorio
    Dim Im As String, Ic As String
    Dim Relatorio As String
    Dim Filtro As String
    Set RELAT = New VSRelatorio
    With RELAT
        Select Case TipoCertidao
            Case "TCER101": Relatorio = "\TCN.rpt"
            Case "TCER102": Relatorio = "\TCN.rpt"
            Case "TCER103": Relatorio = "\TCN.rpt"
            Case "TCER104": Relatorio = "\TCN.rpt"
        End Select

        If Not .DefinirArquivo(Bdados, App.Path + "\TCN.rpt") Then Exit Sub
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            .Selecao = "{TAB_CERTIDAO_NEGATIVA.TCN_COD_NEGATIVA} = " & CodCertidao
            .Titulo = "Certidão Negativa de Débitos"
            Select Case TipoCertidao
                Case "TCER101"
                    .Formulas "VT_TITULO", "CND - CERTIDÃO NEGATIVA DE DÉBITOS"
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, txtEndereco)
                    .SubRelatorio = "TextoCert"
                    .Selecao = "{TAB_PARAMETRO_TEXTO.TPT_PARAMETRO} = 'Certidao Negativa'"
                Case "TCER102":
                    .Formulas "VT_TITULO", "CPD - CERTIDÃO POSITIVA DE DÉBITO"
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, txtEndereco)
                    .SubRelatorio = "TextoCert"
                    .Selecao = "{TAB_PARAMETRO_TEXTO.TPT_PARAMETRO} = 'CPD'"
                Case "TCER103":
                    .Formulas "VT_TITULO", "CPND - CERTIDÃO POSITIVA COM EFEITO DE NEGATIVA DE DÉBITO"
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, txtEndereco)
                    .SubRelatorio = "TextoCert"
                    .Selecao = "{TAB_PARAMETRO_TEXTO.TPT_PARAMETRO} = 'CPND'"
                Case "TCER104"
                    .Formulas "VT_TITULO", "CND - CERTIDÃO NEGATIVA DE DÉBITOS"
                    .Formulas "VT_ENDERCO_IMOV", IIf(Len(Trim(txtIm)) = 14 Or Len(Trim(txtIm)) = 15, txtIm & " - " & txtEndereco, txtEndereco)
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
    If Trim(txtIm) = "" Then: Util.Informa "Informe a inscrição.": Exit Sub
            
    'busca certidoes ja emitidas
'    If Len(Trim(txtIm)) = 10 Then
'        InscricaoMun = txtIm: InscricaoCad = ""
'    Else
'        InscricaoCad = txtIm: InscricaoMun =
'    End If
    Select Case TipoCertidao
        Case "TCER101": Tipo = tcCND
            Certidao.BuscarCertidoes grdCPND, Tipo, InscricaoMun, InscricaoCad
            'busca debitos abertos nao vencidos
            If Obrig.CarregaListaObrigacao(grdDebitosVencido, txtIm, CStr(cboTributo.Coluna(1).Valor), , etlNaoPagos) Then
                Util.Informa "Não é possível emitir a CND. Existem debitos."
                cmdEmitir.Enabled = False
            Else
                cmdEmitir.Enabled = True
                If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
            End If
        Case "TCER102": Tipo = tcCPD
            Certidao.BuscarCertidoes grdCPND, Tipo, txtIm
            If Obrig.CarregaListaObrigacaoVencida(grdDebitosVencido, txtIm, CStr(cboTributo.Coluna(1).Valor), , etlNaoPagos) = False Then
                Util.Informa "Não é possível emitir a CPD. Não existem créditos vencidos."
                cmdEmitir.Enabled = False
            Else
                cmdEmitir.Enabled = True
                If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
            End If
        Case "TCER103": Tipo = tcCPND
            Certidao.BuscarCertidoes grdCPND, Tipo, txtIm
            If Obrig.CarregaListaObrigacaoNaoVencida(grdDebitosVencido, txtIm, CStr(cboTributo.Coluna(1).Valor), , etlNaoPagos) = False Then
                Util.Informa "Não é possível emitir a CPND. Não existem créditos não vencidos."
                cmdEmitir.Enabled = False
            Else
                cmdEmitir.Enabled = True
                If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
            End If
        Case "TCER104": Tipo = tcCNE
            Certidao.BuscarCertidoes grdCPND, Tipo, txtIm, ""
            If Obrig.CarregaListaObrigacao(grdDebitosVencido, txtIm, CStr(cboTributo.Coluna(1).Valor), , etlNaoPagos) = False Then
                Util.Informa "Não é possível emitir a CNE. Não existem debitos."
                cmdEmitir.Enabled = False
            Else
                cmdEmitir.Enabled = True
                If Util.Confirma("Certidao liberada. Deseja imprimi-la?") Then cmdEmitir_Click
            End If
        End Select
End Sub

Private Sub cmdEmitir_Click()
  
    Select Case TipoCertidao
        Case "TCER101": If grdDebitosVencido.ListItems.Count > 0 Then Exit Sub
        Case "TCER102": If grdDebitosVencido.ListItems.Count = 0 Then Exit Sub
        Case "TCER103": If grdDebitosVencido.ListItems.Count = 0 Then Exit Sub
        Case "TCER104": If grdDebitosVencido.ListItems.Count = 0 Then Exit Sub
    End Select
    'If grdDebitosVencido.ListItems.Count > 0 Then Exit Sub
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
            Case "TCER104"
                .GravarTexto "CNE", txtTexto
                .Tipo = tcCNE
        End Select
        
        
        CodCertidao = Conta.GeraCodPagamento("37")
        .CodNegativa = CodCertidao
        .Im = InscricaoMun
        If Trim(InscricaoCad) <> "" Then .Ic = Bdados.Converte(InscricaoCad, tctexto)
        .DataNegativa = Format(Date, "DD/MM/YYYY")
        .Finalidade = txtFinalidade
        .Validade = txtValidade
        .PeriodoInicial = Edita.TiraPic(IIf(txtRefInicio <> "", txtRefInicio, Format(Date, "mm/yyyy")), "/")
        .PeriodoFinal = Edita.TiraPic(txtRefFim, "/")
        .CodUsuario = Aplicacoes.Usuario
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

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    cmdEmitir.Enabled = True
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
        Case "TCER104": txtTexto = Certidao.BuscaTexto("CNE")
    End Select

    cboTributo.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, TipoCertidao, App.Path
    rodVISUAL1.Exibir Bdados, TipoCertidao, App.Major, App.Minor, App.Revision
    Dim DataValid As String
    Dim Mes As String
    Dim Ano As String
    Dim Dia As String
    Set Certidao = New iCertidao
    Set Conta = New ContaCorrente
    TipoCertidao = "TCER106"
    Select Case TipoCertidao
        Case "TCER101": txtTexto = Certidao.BuscaTexto("CERTIDAO NEGATIVA"): grdDebitosVencido.Caption = "Débitos"
        Case "TCER102": txtTexto = Certidao.BuscaTexto("CPD"): grdDebitosVencido.Caption = "Débitos vencidos"
        Case "TCER103": txtTexto = Certidao.BuscaTexto("CPND"): grdDebitosVencido.Caption = "Débitos não vencidos"
        Case "TCER104": txtTexto = Certidao.BuscaTexto("CNE"): grdDebitosVencido.Caption = "Débitos vencidos"
    End Select
    Certidao.PreencherCboImposto cboTributo
    DataValid = Format(Date, "dd/mm/yyyy")
    Dia = Mid(DataValid, 1, 2)
    Mes = Mid(DataValid, 4, 2)
    Ano = Mid(DataValid, 7, 4)
    Mes = Mes + 4
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
    If Len(Trim(txtIm)) = 10 Then
        txtIm.AgruparValores = False
        txtIm.Formato = formDoisDigitos
        txtIm = BuscaContribuinte(txtIm, txtContrib, txtEndereco, Doc)
        InscricaoMun = txtIm: InscricaoCad = ""
    Else
        txtIm = Obrig.BuscaSujeitoPassivoObrigacao(txtIm, txtContrib, txtEndereco, InscricaoMun)
        txtContrib = InscricaoMun & " - " & txtContrib
        InscricaoCad = txtIm
    End If
    txtIm.Formato = formNenhum
    txtIm.AgruparValores = True
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
