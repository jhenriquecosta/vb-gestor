VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "CABECALHO.OCX"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#2.0#0"; "VTControles.ocx"
Begin VB.Form TCIP201 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   33
      Top             =   5835
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   8100
         TabIndex        =   3
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   6930
         TabIndex        =   2
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9270
         TabIndex        =   4
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   1138
      Icone           =   "TCIP201.frx":0000
   End
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   5025
      Left            =   75
      TabIndex        =   6
      Top             =   660
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   8864
      Caption         =   "Dados Gerais"
      Descricao       =   "Informações gerais do imóvel a ser excluído"
      corFaixa        =   32768
      Icone           =   "TCIP201.frx":031A
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.fraVISUAL fraVISUAL3 
         Height          =   780
         Left            =   105
         TabIndex        =   29
         Top             =   4155
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   1376
         Altura          =   1905
         Caption         =   " Detalhes"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtValor 
            Height          =   285
            Left            =   4875
            TabIndex        =   32
            Top             =   360
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   503
            Caption         =   "Valor Venal(R$)"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   3
         End
         Begin VTOcx.cboVISUAL cboAforado 
            Height          =   315
            Left            =   3510
            TabIndex        =   31
            Top             =   360
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            AutoFocaliza    =   0   'False
            Enabled         =   0   'False
         End
         Begin VTOcx.txtVISUAL txtAnoAq 
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   503
            Caption         =   "Ano de Aquisição"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            MaxLen          =   4
            MinLen          =   4
         End
         Begin VB.Label Label6 
            Caption         =   "Aforado"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2790
            TabIndex        =   39
            Top             =   405
            Width           =   735
         End
      End
      Begin VTOcx.fraVISUAL fraVISUAL2 
         Height          =   1860
         Left            =   105
         TabIndex        =   8
         Top             =   2235
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   3281
         Altura          =   1905
         Caption         =   " Contribuinte"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.cboVISUAL cboSitCad 
            Height          =   315
            Left            =   6015
            TabIndex        =   28
            Top             =   1470
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            AutoFocaliza    =   0   'False
            Enabled         =   0   'False
         End
         Begin VTOcx.cboVISUAL cboUF 
            Height          =   315
            Left            =   5130
            TabIndex        =   27
            Top             =   1470
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            AutoFocaliza    =   0   'False
            Enabled         =   0   'False
         End
         Begin VTOcx.txtVISUAL txtMunic 
            Height          =   480
            Left            =   1725
            TabIndex        =   26
            Top             =   1290
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   847
            Caption         =   "Município"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtCep 
            Height          =   480
            Left            =   105
            TabIndex        =   25
            Top             =   1290
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   847
            Caption         =   "CEP"
            Text            =   ""
            Enabled         =   0   'False
            Formato         =   4
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtBairroContrib 
            Height          =   480
            Left            =   7425
            TabIndex        =   24
            Top             =   780
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   847
            Caption         =   "Bairro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtCompContrib 
            Height          =   480
            Left            =   5115
            TabIndex        =   23
            Top             =   780
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   847
            Caption         =   "Complemento"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNumeroContrib 
            Height          =   480
            Left            =   4470
            TabIndex        =   22
            Top             =   780
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   847
            Caption         =   "Nº"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNomeLogrContrib 
            Height          =   285
            Left            =   1725
            TabIndex        =   21
            Top             =   975
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   503
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
         End
         Begin VTOcx.cboVISUAL cboTipoLogrContrib 
            Height          =   315
            Left            =   105
            TabIndex        =   20
            Top             =   960
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            AutoFocaliza    =   0   'False
            Enabled         =   0   'False
         End
         Begin VTOcx.txtVISUAL txtNomeContrib 
            Height          =   480
            Left            =   1575
            TabIndex        =   19
            Top             =   285
            Width           =   8670
            _ExtentX        =   15293
            _ExtentY        =   847
            Caption         =   "Nome"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtIm 
            Height          =   480
            Left            =   105
            TabIndex        =   18
            Top             =   285
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   847
            Caption         =   "Insc. Municipal"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoRotulo=   1
            Mascara         =   "00000000-00"
         End
         Begin VB.Label Label5 
            Caption         =   "Sit. Cadastral"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6015
            TabIndex        =   38
            Top             =   1275
            Width           =   1230
         End
         Begin VB.Label Label4 
            Caption         =   "UF"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5130
            TabIndex        =   37
            Top             =   1275
            Width           =   435
         End
         Begin VB.Label Label2 
            Caption         =   "Logradouro"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   105
            TabIndex        =   35
            Top             =   765
            Width           =   1635
         End
      End
      Begin VTOcx.fraVISUAL fraVISUAL1 
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   705
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   2566
         Altura          =   1905
         Caption         =   " Imóvel"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.cmdVISUAL cmdOpcao 
            Height          =   330
            Index           =   2
            Left            =   1485
            TabIndex        =   1
            Top             =   480
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   582
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtSecao 
            Height          =   480
            Left            =   6210
            TabIndex        =   17
            Top             =   855
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   847
            Caption         =   "Seção"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtLote 
            Height          =   480
            Left            =   5385
            TabIndex        =   16
            Top             =   855
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   847
            Caption         =   "Lote"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtQuadra 
            Height          =   480
            Left            =   4560
            TabIndex        =   15
            Top             =   855
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   847
            Caption         =   "Quadra"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtLoteamento 
            Height          =   480
            Left            =   3495
            TabIndex        =   14
            Top             =   855
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   847
            Caption         =   "Loteamento"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.cboVISUAL cboBairro 
            Height          =   315
            Left            =   105
            TabIndex        =   13
            Top             =   1035
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            AutoFocaliza    =   0   'False
            Enabled         =   0   'False
         End
         Begin VTOcx.txtVISUAL txtComplemento 
            Height          =   480
            Left            =   6900
            TabIndex        =   12
            Top             =   315
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   847
            Caption         =   "Complemento"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNumero 
            Height          =   480
            Left            =   6150
            TabIndex        =   11
            Top             =   315
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   847
            Caption         =   "Nº"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.cboVISUAL cboLogr 
            Height          =   315
            Left            =   3465
            TabIndex        =   10
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            AutoFocaliza    =   0   'False
            Enabled         =   0   'False
         End
         Begin VTOcx.txtVISUAL txtIc 
            Height          =   480
            Left            =   120
            TabIndex        =   0
            Top             =   315
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   847
            Caption         =   "Insc. Cadastral"
            Text            =   ""
            Formato         =   7
            Restricao       =   2
            AlinhamentoRotulo=   1
            AgruparValores  =   0   'False
         End
         Begin VTOcx.cboVISUAL cboTipoLogr 
            Height          =   315
            Left            =   1875
            TabIndex        =   9
            Top             =   495
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            AutoFocaliza    =   0   'False
            Enabled         =   0   'False
         End
         Begin VB.Label Label3 
            Caption         =   "Bairro"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   36
            Top             =   825
            Width           =   690
         End
         Begin VB.Label Label1 
            Caption         =   "Logradouro"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1905
            TabIndex        =   34
            Top             =   300
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "TCIP201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As VSImposto
Dim Imovel As cImovel
Dim Contribuinte As cContribuinte
Dim Endereco As cEndereco

Private Sub cmdExcluir_Click()
     Screen.MousePointer = 11
            If Util.Confirma("Deseja realmente eliminar o imóvel?") Then
                'Verifica se tem empresa cadastrada no imovel
                If Imovel.TemEmpresaNoImovel(txtIc) Then
                    Call Util.Avisa("Não é possível eliminar este imóvel. Há uma empresa cadastrada neste endereço.")
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                'verifica se existe o imovel
                If Imovel.Buscar(txtIc) Then
                    'exclui o imovel
                    If Imovel.Excluir(txtIc) Then
                        Call Util.Informa("Registro eliminado com sucesso.")
                        cmdLimpar_Click
                    Else
                        Erro "Erro ao excluir."
                    End If
                Else
                    Call Util.Informa("Imóvel não cadastrado.")
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                DoEvents
                Screen.MousePointer = 0
            Else
                Screen.MousePointer = 0
            End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtIc.Enabled = True
    txtIc.SetFocus
End Sub

Private Sub cmdOpcao_Click(Index As Integer)
    Static JaHabilitou As Boolean
    Select Case Index
        Case 2
            txtIc = AplicacoesVTFuncoes.BuscaNoImobiliario
    End Select
End Sub



Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim Controle As Control
    Dim i As Byte
    '----------setando classes
    Set cadastro = New VSImposto
    Set Imovel = New cImovel
    Set Contribuinte = New cContribuinte
    Set Endereco = New cEndereco
    '-----------Preenchendo as combos
    With Endereco
        .PreencherComboLogr cboLogr
        .PreencherComboTipoLogr cboTipoLogr
        .PreencherComboBairro cboBairro
        .PreencherComboTipoLogr cboTipoLogrContrib
    End With
    Contribuinte.PreencherCboSitCad cboSitCad
    cboUF.PreencherGeral Bdados, "UF"
    cboAforado.PreencherGeral Bdados, "SIM OU NÃO"
    
    Screen.MousePointer = 0
    '----------preenchendo cabecalho e rodape
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set Imovel = Nothing
    Set Contribuinte = Nothing
    Set Endereco = Nothing
End Sub

Private Sub txtic_LostFocus()
    Dim RsAux As VSRecordset
    Dim sql As String
    Dim Rs As VSRecordset
    Dim Tipologr As String, Logr As String, Numero As String, Complemento As String, Bairro As String, Loteamento As String, Lote As String, Quadra As String, Secao As String, IM As String, AnoAq As String, Aforado As String, Valor As String
    
    If Trim(txtIc) = "" Then Exit Sub
    If Me.ActiveControl.Name = "cmdSair" Or Me.ActiveControl.Name = "cmdLimpar" Then Exit Sub
    If Imovel.BuscarVisImovel(txtIc, Tipologr, Logr, Numero, Complemento, Bairro, Loteamento, Lote, Quadra, Secao, IM, AnoAq, Aforado, Valor) Then
        'preenchendo dados do imovel
        cboTipoLogr = Tipologr
        cboLogr = Logr
        txtNumero = Numero
        txtComplemento = Complemento
        cboBairro = Bairro
        txtLoteamento = Loteamento
        txtLote = Lote
        txtQuadra = Quadra
        txtSecao = Secao
        'preenchendo dados do contribuinte
        txtIM = IM
        txtIm_LostFocus
        txtAnoAq = AnoAq
        cboAforado.ListIndex = IIf(Aforado = "N", 0, 1)
        txtValor = Format$(Valor, Const_Monetario)
        txtIc.Enabled = False
        'preenche situacao cadastral do imovel
        cboSitCad.ListIndex = Contribuinte.VerificaSitCadastral(IM)
    Else
        Call Util.Avisa("Imóvel não cadastrado.")
        Call Edita.LimpaCampos(Me)
        txtIc.Enabled = True
        txtIc.SetFocus
    End If
    Screen.MousePointer = 0
End Sub

Private Sub txtIm_LostFocus()
    Dim sql As String
    Dim Rs As VSRecordset
    Dim NomeContrib As String, TipoLogrContr As String, LogrContr As String, NumeroContr As String, CompContri As String, BairroContr As String, CepContr As String, MunicContr As String
    
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtIM) <> "" Then
        If Contribuinte.BuscarContribuinte(txtIM, NomeContrib, TipoLogrContr, LogrContr, NumeroContr, CompContri, BairroContr, CepContr, MunicContr) Then
            txtNomeContrib = NomeContrib
            cboTipoLogrContrib.ListIndex = cadastro.BuscaCodLogr(TipoLogrContr) - 1
            txtNomeLogrContrib = LogrContr
            txtNumeroContrib = NumeroContr
            txtCompContrib = CompContri
            txtBairroContrib = BairroContr
            txtCep = CepContr
            txtMunic = MunicContr
            cboUF.ListIndex = 0
        Else
            Call Util.Informa("Contribuinte não encontrado.")
            txtIM.Enabled = True
        End If
    End If
End Sub

