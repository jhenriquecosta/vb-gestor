VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIS402 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIS402"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      TabIndex        =   26
      Top             =   7110
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1058
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6390
         TabIndex        =   29
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cboVISUAL cboModoImpressao 
         Height          =   315
         Left            =   1785
         TabIndex        =   20
         Tag             =   "Modo Impressão"
         Top             =   165
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   556
         Caption         =   "Md. Impressão"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   5220
         TabIndex        =   21
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   7545
         TabIndex        =   22
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8700
         TabIndex        =   23
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   -540
      Width           =   375
   End
   Begin VTOcx.grdVISUAL grdContribuinte 
      Height          =   3090
      Left            =   -15
      TabIndex        =   25
      Top             =   4035
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   5450
      CorBorda        =   16711680
      Caption         =   "Contribuintes"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   16711680
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1410
      Left            =   0
      TabIndex        =   28
      Top             =   660
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   2487
      Altura          =   1905
      Caption         =   " Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboTipoCadastro 
         Height          =   510
         Left            =   5850
         TabIndex        =   3
         Top             =   300
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   900
         Caption         =   "Tipo de Cadastro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.txtVISUAL txtIAnterior 
         Height          =   480
         Left            =   3630
         TabIndex        =   2
         Tag             =   "CPF ou  CGC"
         Top             =   330
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   847
         Caption         =   "Inscricão Anterior"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   480
         Left            =   90
         TabIndex        =   0
         Top             =   330
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   847
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Formato         =   8
         Restricao       =   2
         AlinhamentoRotulo=   1
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtFantasia 
         Height          =   480
         Left            =   5880
         TabIndex        =   5
         Top             =   840
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   847
         Caption         =   "Nome Fantasia"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   480
         Left            =   105
         TabIndex        =   4
         Tag             =   "Nome ou Razão Social"
         Top             =   840
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   847
         Caption         =   "Nome ou Razão Social"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtCgc 
         Height          =   480
         Left            =   1545
         TabIndex        =   1
         Tag             =   "CPF ou  CGC"
         Top             =   330
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   847
         Caption         =   "CPF ou CNPJ"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.cboVISUAL cboPonto 
         Height          =   510
         Left            =   5085
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   900
         Caption         =   "Ponto Recepção"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   1935
      Left            =   0
      TabIndex        =   27
      Top             =   2085
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   3413
      Altura          =   1905
      Caption         =   " Localização"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cboVISUAL cboSitCad 
         Height          =   510
         Left            =   5415
         TabIndex        =   16
         Tag             =   "Situação Cadastral"
         Top             =   1350
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   900
         Caption         =   "Situação Cadastral"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.txtVISUAL txtCidade 
         Height          =   480
         Left            =   2100
         TabIndex        =   12
         Tag             =   "Cidade"
         Top             =   840
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   847
         Caption         =   "Cidade"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtNum 
         Height          =   480
         Left            =   5460
         TabIndex        =   9
         Top             =   330
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   847
         Caption         =   "Nº"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.cboVISUAL cboLogr 
         Height          =   510
         Left            =   1710
         TabIndex        =   8
         Tag             =   "Logradouro"
         Top             =   300
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   900
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         Editavel        =   -1  'True
      End
      Begin VTOcx.cboVISUAL cboTipoLogr 
         Height          =   510
         Left            =   75
         TabIndex        =   7
         Tag             =   "Logradouro"
         Top             =   300
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   900
         Caption         =   "Logradouro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboUF 
         Height          =   510
         Left            =   4560
         TabIndex        =   13
         Tag             =   "UF"
         Top             =   825
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
         Caption         =   "UF"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.txtVISUAL txtComplemento 
         Height          =   480
         Left            =   6255
         TabIndex        =   10
         Top             =   330
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   847
         Caption         =   "Complemento"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtbairro 
         Height          =   480
         Left            =   75
         TabIndex        =   11
         Tag             =   "Cidade"
         Top             =   840
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   847
         Caption         =   "Bairro"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.cboVISUAL cboAtivServ 
         Height          =   510
         Left            =   75
         TabIndex        =   15
         Tag             =   "Atividade Principal"
         Top             =   1350
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   900
         Caption         =   "Atividade Principal"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboObrigIss 
         Height          =   510
         Left            =   5400
         TabIndex        =   14
         Tag             =   "Obrigação do ISSQN"
         Top             =   825
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   900
         Caption         =   "Obrigação do ISSQN"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.txtVISUAL txtPeriodo1 
         Height          =   480
         Left            =   5445
         TabIndex        =   17
         Top             =   1980
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   847
         Caption         =   "Periodo Inicial"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtPeriodo2 
         Height          =   480
         Left            =   6810
         TabIndex        =   18
         Top             =   1980
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   847
         Caption         =   "Periodo Final"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtPeriodoAlvara 
         Height          =   480
         Left            =   8625
         TabIndex        =   19
         Top             =   1980
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   847
         Caption         =   "Período Alvará"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1138
      Icone           =   "TCIS402.frx":0000
   End
   Begin VB.Menu mnuGeral 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuCad 
         Caption         =   "Cadastro"
      End
      Begin VB.Menu mnuLanca 
         Caption         =   "Lancamentos"
      End
   End
End
Attribute VB_Name = "TCIS402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Cadastro As VSImposto
Dim Endereco As eEndereco
Dim atividade As atividade
Dim Contribuinte As eContribuinte
Dim CodAtividade As Long
Dim FiltroRpt As String

Private Sub grdContribuinte_dblClick()
If grdContribuinte.ListItems.Count >= 1 Then
        Dim Sql As String
        Dim rs As VSRecordset
        Dim UFM As Double
        
        UFM = Temp.PegaParametro(Bdados, "UFM")
        
        If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
            Sql = "SELECT TFL,TFA,TFAF,TFS,TFOP  "
            Sql = Sql & " FROM VIS_TAXA_ATIVIDADE_COMPLETA "
            Sql = Sql & " WHERE TAE_CAE = '" & Imposto.BuscaCAE(grdContribuinte.SelectedItem.SubItems(11)) & "'"
            If Bdados.AbreTabela(Sql, rs) Then
                grdContribuinte.Mensagem = "TFL - " & Format(rs.Fields("TFL") * UFM, Const_Monetario) & "   TFA - " & Format(rs.Fields("TFA") * UFM, Const_Monetario) & "  TFAF - " & Format(rs.Fields("TFAF") * UFM, Const_Monetario) & "  TFS - " & Format(rs.Fields("TFS"), Const_Monetario) & "  TFOP - " & Format(rs.Fields("TFOP") * UFM, Const_Monetario)
            Else
                grdContribuinte.Mensagem = ""
            End If
        End If
    End If
End Sub

Private Sub mnuLanca_Click()
    Dim ProjObrig As Object
    
    Set ProjObrig = CreateObject("VSTOBRI.Aplicacoes")
        
    Set ProjObrig.Banco = Bdados.Conexao
    ProjObrig.Usuario = AplicacoesVTFuncoes.Usuario
    ProjObrig.Codigo_Municipio = AplicacoesVTFuncoes.Codigo_Municipio
    ProjObrig.municipio = AplicacoesVTFuncoes.municipio
    TempContrib = Trim(grdContribuinte.SelectedItem)
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
        ProjObrig.Abre_Aplicacao "TOBR401", 0, Cod_sis, Sistema, Desc_Form, "C" & Trim(grdContribuinte.SelectedItem)
    Else
        ProjObrig.Abre_Aplicacao "TOBR401", 0, Cod_sis, Sistema, Desc_Form, Trim(grdContribuinte.SelectedItem)
    End If
    
    TempContrib = Trim(grdContribuinte.SelectedItem)
    
End Sub

Private Sub mnuCad_Click()
    Dim ProjObrig As Object
    TCIS103.Show
    Data = Trim(grdContribuinte.SelectedItem.SubItems(19))
    Motivo = Trim(grdContribuinte.SelectedItem.SubItems(20))
    TCIS103.Tag = Trim(grdContribuinte.SelectedItem)
    
End Sub

Private Sub cboTipoLogr_Click()
'    Endereco.PreencherCboRua cboLogr, cboTipoLogr
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo trata
    Dim Iss As String
    
    If cboObrigIss.ListIndex <> -1 Then
        Iss = cboObrigIss.Coluna(1).Valor
    End If
    If Not Contribuinte.BuscarContribuintesHistorico(grdContribuinte, txtIm, txtCgc, txtRazao, txtFantasia, cboTipoLogr, cboLogr, txtBairro, txtNum, txtComplemento, _
                    txtCidade, CStr(cboAtivServ.Coluna(1).Valor), Iss, txtPeriodo1, txtPeriodo2, cboPonto, FiltroRpt, txtIAnterior, CStr(cboTipoCadastro.Coluna(1).Valor), CStr(cboSitCad.Coluna(1).Valor)) Then
        Util.Avisa "Consulta sem resultados."
    End If
    Exit Sub
trata:
    Erro Err.Number & " - " & Err.Description
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{Tab}"
End Sub

Private Sub cmdImprimir_Click()
        If grdContribuinte.ListItems.Count > 0 Then
            Dim i As Integer
            Dim Im As String
            Screen.MousePointer = 11
            If Left(cboModoImpressao, 1) = 1 Then
                If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TCIS402LISTAGEM.rpt") Then Screen.MousePointer = 0: Exit Sub
                Rpt.Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                Rpt.Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
                Rpt.Selecao = FiltroRpt
                Rpt.Arvore = False
                Rpt.Visualizar
                Screen.MousePointer = 0
            ElseIf Left(cboModoImpressao, 1) = 2 Then
                If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TFICHA_CAD_ECONOMICO_HISTORICO.rpt") Then Screen.MousePointer = 0: Exit Sub
                Rpt.Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                Rpt.Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
                Rpt.Selecao = FiltroRpt
                Rpt.Arvore = False
                Rpt.Visualizar
                Screen.MousePointer = 0
            End If
            Set Rpt = Nothing
        Else
            Call Util.Informa("Selecione o(s) Contribuinte no botão buscar.")
        End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    grdContribuinte.ListItems.Clear
    txtIm.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Controle As Control
    Dim i As Byte
    Set Cadastro = New VSImposto
    Set Endereco = New eEndereco
    Set atividade = New atividade
    Set Contribuinte = New eContribuinte
    Endereco.PreencherCboRua cboLogr
    Endereco.PreencherCboTipoLogr cboTipoLogr
    Endereco.PreencherPonto cboPonto
    atividade.PreencherCboAtiv cboAtivServ
    cboUf.PreencherGeral Bdados, "UF"
    cboObrigIss.PreencherGeral Bdados, "TIPO RECOLHIMENTO ISS"
    Contribuinte.PreencherCboSitCad cboSitCad
    cboModoImpressao.AddItem "1 - Listagem"
    If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
        cboModoImpressao.AddItem "2 - Ficha"
    End If
    If AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        txtIAnterior.Visible = False
        txtIm.Formato = formNenhum
    End If
    cboTipoCadastro.PreencherGeral Bdados, "TIPO CADASTRO ECONOMICO"

    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" And AplicacoesVTFuncoes.municipio <> "VERDEJANTE" Then
        txtIm.Formato = formNenhum
    Else
        txtIm.Formato = formDoisDigitos
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set Cadastro = Nothing
    Set Endereco = Nothing
    Set atividade = Nothing
End Sub

Private Sub grdContribuinte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 And grdContribuinte.ListItems.Count > 0 Then
        mnuCad.Caption = "Consultar Cadastro " & grdContribuinte.SelectedItem

        Me.PopupMenu mnuGeral
    End If
End Sub

Private Sub txtcgc_LostFocus()
    If Trim(txtCgc) = "" Then Exit Sub
    If Len(txtCgc) = 11 And IsNumeric(txtCgc) Then
        txtCgc.Formato = formCPF
    ElseIf Len(txtCgc) = 14 And IsNumeric(txtCgc) Then
        txtCgc.Formato = formCGC
    ElseIf Trim(txtCgc) <> "" Then
        Call Util.Informa("Número de CNPJ ou CPF inválido.")
    End If
    txtCgc.Formato = formNenhum
End Sub

