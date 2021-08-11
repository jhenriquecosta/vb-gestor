VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TDNT401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TDNT401"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL GrdDados 
      Height          =   3090
      Left            =   -15
      TabIndex        =   18
      Top             =   4095
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   5450
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TDNT401.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   1138
      Icone           =   "TDNT401.frx":2123
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   15
      Top             =   7200
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL CmdBuscar 
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   75
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdsair 
         Height          =   375
         Left            =   8730
         TabIndex        =   13
         Top             =   75
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Sair"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprime 
         Height          =   375
         Left            =   5580
         TabIndex        =   11
         Top             =   75
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7725
         TabIndex        =   12
         Top             =   75
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   3405
      Left            =   0
      TabIndex        =   16
      Top             =   660
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   6006
      Altura          =   1905
      Caption         =   " Imprimir Por:"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   870
         TabIndex        =   1
         Top             =   840
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   300
         Left            =   3360
         TabIndex        =   8
         Top             =   2940
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   529
         Caption         =   "Periodo Inicial"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoFinal 
         Height          =   300
         Left            =   7020
         TabIndex        =   9
         Top             =   2940
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   529
         Caption         =   "Periodo Final"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   315
         Left            =   9210
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1530
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   5790
         TabIndex        =   4
         Top             =   1530
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "Cadastro do Imóvel"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   960
         TabIndex        =   21
         Top             =   1860
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   556
         Caption         =   "Razão"
         Text            =   ""
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3660
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1500
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   690
         TabIndex        =   19
         Top             =   2220
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   705
         TabIndex        =   3
         Top             =   1530
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   529
         Caption         =   "Inscrição"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtFolhaFinal 
         Height          =   315
         Left            =   7200
         TabIndex        =   7
         Top             =   2565
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         Caption         =   "Folha Final"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtFolhaInicial 
         Height          =   315
         Left            =   3540
         TabIndex        =   6
         Top             =   2565
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         Caption         =   "Folha Inicial"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtLivro 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   2565
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         Caption         =   "Livro"
         Text            =   ""
         Restricao       =   2
         AgruparValores  =   0   'False
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtDoc 
         Height          =   285
         Left            =   510
         TabIndex        =   2
         Tag             =   "Periodo"
         Top             =   1200
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   503
         Caption         =   "Nº Registro"
         Text            =   ""
         TipoLetras      =   0
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Tag             =   "Tipo Documento"
         Top             =   465
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   556
         Caption         =   "Tipo Documento"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
   End
End
Attribute VB_Name = "TDNT401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto

Private Sub cmdBuscar_Click()
    Dim DAT As New cDividaAtiva
    Dim Insc As String
    Dim Tipo As TipoInscricaoObrigacao
    If txtIm <> "" Then
        Insc = txtIm
        Tipo = etiContribuinte
    ElseIf txtImovel <> "" Then
        Tipo = etiImovel
        Insc = txtImovel
    End If
    DAT.CarregaDividaGerada GrdDados, Insc, , txtPeriodoInicial, txtPeriodoFinal, CStr(cboImposto.Coluna(0).Valor), txtDoc, txtFolhaInicial, txtFolhaFinal, txtLivro, Tipo, edNaoTributaria
End Sub
Private Sub cmdImprime_Click()
    Dim Sql As String
    'If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    Dim Autoridade As String
    Dim MetodologiaCalculo As String
    Dim lEI As String
    Dim Cargo As String
    Dim SelecaoRpt As String
    Dim Lancamento As String
    Dim rs As VSRecordset
    Dim Ano As String
    Dim Divida As New cDividaAtiva
    
    txtDoc.Tag = txtDoc.Caption
    
    If cboTipo.Coluna(0).Valor = 4 Then
        If txtFolhaInicial <> "" And txtFolhaFinal <> "" Then
            If Val(txtFolhaInicial) > Val(txtFolhaFinal) Then
                Util.Avisa "Folha inicial deve ser maior que folha final."
                txtFolhaInicial.SetFocus
                Exit Sub
            End If
        End If
    ElseIf cboTipo.Coluna(0).Valor = 1 Or cboTipo.Coluna(0).Valor = 2 Then
        If txtPeriodoInicial = "" Then
            Avisa "Informe o período inicial."
            txtPeriodoInicial.SetFocus
            Exit Sub
        End If
        
        If cboImposto.ListIndex < 0 Then
            Avisa "Selecione tributo."
            cboImposto.SetFocus
            Exit Sub
        End If
    End If
    
    MetodologiaCalculo = Divida.BuscaParametro("METODOLOGIA DE CALCULO", edNaoTributaria)
    
    SelecaoRpt = " {TAB_DIVIDA_ATIVA.TDA_TIPO_DIVIDA} = '2'"
    If txtDoc <> "" Then
        SelecaoRpt = SelecaoRpt & " and  {TAB_DIVIDA_ATIVA.TDA_REGISTRO} = " & txtDoc
    End If
    If txtIm <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_DIVIDA_ATIVA.tda_inscricao} = '" & txtIm & "'"
    ElseIf txtImovel <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_DIVIDA_ATIVA.tda_inscricao} = '" & txtImovel & "'"
    End If
    
    If txtLivro <> "" Then
        SelecaoRpt = "{TAB_DIVIDA_ATIVA.TDA_LIVRO} = '" & txtLivro & "'"
    End If
    If txtFolhaFinal <> "" And txtFolhaInicial <> "" Then
        SelecaoRpt = SelecaoRpt & " AND {TAB_DIVIDA_ATIVA.TDA_FOLHA} >= " & txtFolhaInicial & " AND {TAB_DIVIDA_ATIVA.TDA_FOLHA} <= " & txtFolhaFinal
    Else
        If txtFolhaInicial <> "" And txtFolhaFinal = "" Then
            SelecaoRpt = SelecaoRpt & " AND {TAB_DIVIDA_ATIVA.TDA_FOLHA} >= " & txtFolhaInicial & " AND {TAB_DIVIDA_ATIVA.TDA_FOLHA} <= " & txtFolhaInicial
        End If
    End If
    
    If txtPeriodoInicial <> "" And txtPeriodoFinal <> "" Then
        SelecaoRpt = SelecaoRpt & " AND {TAB_DIVIDA_ATIVA.TDA_ANO_DIVIDA} >= " & txtPeriodoInicial & " AND {TAB_DIVIDA_ATIVA.TDA_ANO_DIVIDA} <= " & txtPeriodoFinal
    Else
        If txtPeriodoInicial <> "" And txtPeriodoFinal = "" Then
            SelecaoRpt = SelecaoRpt & " AND {TAB_DIVIDA_ATIVA.TDA_ANO_DIVIDA} >= " & txtPeriodoInicial & " AND {TAB_DIVIDA_ATIVA.TDA_ANO_DIVIDA} <= " & txtPeriodoInicial
        End If
    End If
    
    If cboImposto.ListIndex >= 0 Then
        SelecaoRpt = SelecaoRpt & " and {TAB_DIVIDA_ATIVA.TDA_TIP_COD_IMPOSTO} = '" & cboImposto.Coluna(0).Valor & "'"
    End If
    Select Case cboTipo.Coluna(0).Valor
        Case 1, 2 'MACAL ou MALIC
            With Rpt
                 If Not .DefinirArquivo(Bdados, App.Path + "\Macal_Malic.rpt") Then: Screen.MousePointer = 0: Exit Sub
                 .Formulas "VT_DOCUMENTO", cboTipo.Text
                 .Selecao = SelecaoRpt
                 .Formulas "VT_DOCUMENTO", cboTipo.Text
                 .Formulas "vt_prefeitura", Temp.PegaParametro(Bdados, "CLIENTE")
                 .Formulas "vt_secretaria", Temp.PegaParametro(Bdados, "SEMFAZ")
                 .SubRelatorio = "TMalic_Macal.rpt"
                 .Formulas "VT_AUTORIDADE", Autoridade
                 .Formulas "VT_CARGO", Cargo
                 .Selecao = " {VIS_MALIC_MACAL.TPD_CODIGO} = '" & cboTipo.Coluna(0).Valor & "' and {VIS_MALIC_MACAL.TPD_ANO} = '" & txtPeriodoInicial & "' and {VIS_MALIC_MACAL.TPD_IMPOSTO} = '" & cboImposto.Coluna(0).Valor & "'"
                 .SubRelatorio = ""
                 .Titulo = cboTipo.Text
                 .Arvore = False
                 .Visualizar
             End With

        Case 3, 5 'TERMODAT
            
            With Rpt
                If cboTipo.Coluna(0).Valor = 3 Then
                    If Not .DefinirArquivo(Bdados, App.Path + "\TTermoDAT.rpt") Then Exit Sub
                Else
                    If Not .DefinirArquivo(Bdados, App.Path + "\TCertidao.rpt") Then Exit Sub
                End If
                If SelecaoRpt <> "" Then
                    .Selecao = SelecaoRpt
                End If
                .Formulas "VT_FORMA_CALCULO", MetodologiaCalculo
                .Formulas "AUTORIDADE", Autoridade
                .Formulas "CARGO", Cargo
                .Formulas "vt_prefeitura", Temp.PegaParametro(Bdados, "CLIENTE")
                .Formulas "vt_secretaria", Temp.PegaParametro(Bdados, "SEMFAZ")
                .Titulo = cboTipo.Text
                .Formulas "DOCUMENTO", cboTipo.Text
                .Arvore = False
                .Visualizar
            End With
            
        Case 4 ' LIVBRO
            With Rpt
                If Not .DefinirArquivo(Bdados, App.Path + "\TLivroDATNaoTributario.rpt") Then Exit Sub
                .Selecao = SelecaoRpt
                .Formulas "vt_prefeitura", Temp.PegaParametro(Bdados, "CLIENTE")
                .Formulas "vt_secretaria", Temp.PegaParametro(Bdados, "SEMFAZ")
                .Formulas "AUTORIDADE", Autoridade
                .Formulas "CARGO", Cargo
                .Titulo = cboTipo.Text
                .Arvore = False
                .Visualizar
            End With
        Case Else
            Avisa "Documento não disponível."
            cboTipo.SetFocus
    End Select
    Set Rpt = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    txtIm.SetFocus
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
    Dim Obrig As New Obrigacao
    cabVisual.Exibir Bdados, Me.Name, App.Path
    cboTipo.Preencher Bdados, "SELECT TGE_CODIGO,TGE_NOME FROM VIS_DOC_NAO_TRIBUTARIO ORDER BY TGE_CODIGO", 1
    Obrig.PreencheComboTributo cboImposto, True, etcNaoTributario
End Sub

Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    Dim Ic As String
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        If Len(txtIm) = 10 Or Len(txtIm) = 11 Then
            Ic = Imposto.FormataInscricao(txtIm, InscContrib)
        Else
            Ic = txtIm
        End If
    Else
            Ic = txtIm
    End If
    txtIm = BuscaContribuinte(Ic, txtRazao, txtEndereco)
End Sub

Private Sub txtImovel_LostFocus()
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtImovel.SetFocus
        End If
    End If
End Sub
