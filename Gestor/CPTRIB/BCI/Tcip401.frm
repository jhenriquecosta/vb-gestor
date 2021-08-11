VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#3.1#0"; "VTControles.ocx"
Begin VB.Form TCIP401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   2265
      Left            =   90
      TabIndex        =   20
      Top             =   765
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   3995
      Altura          =   1905
      Caption         =   " Opções de Busca"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.fraVISUAL fraVISUAL2 
         Height          =   780
         Left            =   75
         TabIndex        =   21
         Top             =   1410
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   1376
         Altura          =   1905
         Caption         =   " Detalhes"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtAnoAq 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   503
            Caption         =   "Ano de Aquisição"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   4
            MinLen          =   4
         End
         Begin VTOcx.cboVISUAL cboAforado 
            Height          =   315
            Left            =   2835
            TabIndex        =   11
            Top             =   360
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   556
            Caption         =   "Aforado"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtValor 
            Height          =   285
            Left            =   4875
            TabIndex        =   12
            Top             =   360
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   503
            Caption         =   "Valor Venal(R$)"
            Text            =   ""
            Restricao       =   3
         End
      End
      Begin VTOcx.txtVISUAL txtSecao 
         Height          =   480
         Left            =   5820
         TabIndex        =   9
         Top             =   855
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   847
         Caption         =   "Seção"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtLote 
         Height          =   480
         Left            =   4950
         TabIndex        =   8
         Top             =   855
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   847
         Caption         =   "Lote"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtQuadra 
         Height          =   480
         Left            =   4095
         TabIndex        =   7
         Top             =   855
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   847
         Caption         =   "Quadra"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtLoteamento 
         Height          =   480
         Left            =   3000
         TabIndex        =   6
         Top             =   855
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   847
         Caption         =   "Loteamento"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.cboVISUAL cboBairro 
         Height          =   510
         Left            =   105
         TabIndex        =   5
         Top             =   825
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   900
         Caption         =   "Bairro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboLogr 
         Height          =   315
         Left            =   7785
         TabIndex        =   4
         Top             =   495
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTipoLogr 
         Height          =   510
         Left            =   6180
         TabIndex        =   3
         Top             =   300
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   900
         Caption         =   "Logradouro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.txtVISUAL txtContrib 
         Height          =   480
         Left            =   2985
         TabIndex        =   2
         Top             =   315
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   847
         Caption         =   "Contribuinte"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   480
         Left            =   1515
         TabIndex        =   1
         Top             =   315
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   847
         Caption         =   "Insc. Municipal"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         Mascara         =   "00000000-00"
      End
      Begin VTOcx.txtVISUAL txtIc 
         Height          =   480
         Left            =   105
         TabIndex        =   0
         Top             =   315
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   847
         Caption         =   "Insc. Cadastral"
         Text            =   ""
         Formato         =   7
         AlinhamentoRotulo=   1
         AgruparValores  =   0   'False
      End
   End
   Begin VTOcx.grdVISUAL Grid 
      Height          =   3540
      Left            =   90
      TabIndex        =   17
      Top             =   3090
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   6244
      CorBorda        =   32768
      Caption         =   "Inscrições Cadastrais"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   19
      Top             =   6720
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   1032
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   0
         Left            =   6750
         TabIndex        =   14
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   5565
         TabIndex        =   13
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   9120
         TabIndex        =   16
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   7935
         TabIndex        =   15
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   1138
      Icone           =   "Tcip401.frx":0000
   End
   Begin VB.Menu mnuPrinc 
      Caption         =   "Principal"
      Visible         =   0   'False
      Begin VB.Menu mnuIPTU 
         Caption         =   "Gerar IPTU"
      End
   End
End
Attribute VB_Name = "TCIP401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As VSImposto
Dim Imovel As cImovel
Dim Endereco As cEndereco

Sub PreencheImovel(RsAux As Object) 'Aki
    cboTipoLogr = RsAux!TTL_NOME
    cboLogr = RsAux!tlg_nome
    cboBairro = RsAux!TBA_NOME
    txtLoteamento = RsAux!Tim_loteamento
    txtLote = RsAux!tim_Lote
    txtQuadra = RsAux!tim_quadra
    txtSecao = RsAux!tim_secao
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim Condicao As String
    Static CondicaoRelatorio As String
    
    Select Case cmd(Index).Caption
        Case "&Buscar"
            Dim Aux As Byte
            Aux = 0
            Condicao = ""
            Screen.MousePointer = 11
            
            Imovel.PreencherGrid Grid, txtIc, txtIm, cboTipoLogr, cboLogr, cboBairro, txtContrib, txtLoteamento, txtQuadra, CStr(cboAforado.Coluna(1).Valor), txtAnoAq, txtValor

            Screen.MousePointer = 0
            DoEvents
        Case "&Imprimir"
            CondicaoRelatorio = ""
            If Trim(txtIc) <> "" Then
                CondicaoRelatorio = CondicaoRelatorio & "  and {TAB_IMOVEL.tim_ic} ='" & txtIc & "'"
            End If
            If Trim(txtIm) <> "" Then
                CondicaoRelatorio = CondicaoRelatorio & "  and {TAB_IMOVEL.tim_tci_im} ='" & txtIm & "'"
            End If
            If Trim(cboTipoLogr) <> "" Then
                CondicaoRelatorio = CondicaoRelatorio & "  and {VIS_BVT.TTL_NOME} ='" & cboTipoLogr & "'"
            End If
            If Trim(cboLogr) <> "" Then
                CondicaoRelatorio = CondicaoRelatorio & "  and {VIS_BVT.tlg_nome} ='" & cboLogr & "'"
            End If
            If Trim(cboBairro) <> "" Then
                CondicaoRelatorio = CondicaoRelatorio & "  and {VIS_BVT.TBA_NOME} ='" & cboBairro & "'"
            End If
            If Trim(txtContrib) <> "" Then
                CondicaoRelatorio = CondicaoRelatorio & "  and {TAB_CONTRIBUINTE.tci_nome} like '" & txtContrib & "%' OR{TAB_CONTRIBUINTE.tci_nome} like '%" & txtContrib & "%'"
            End If
            If Trim(txtLoteamento) <> "" Then
                CondicaoRelatorio = CondicaoRelatorio & "  and mid({TAB_IMOVEL.tim_ic},3,2) ='" & txtLoteamento & "'"
            End If
            If Trim(txtQuadra) <> "" Then
                 CondicaoRelatorio = CondicaoRelatorio & "  and mid({TAB_IMOVEL.tim_ic},5,4) ='" & txtQuadra & "'"
            End If
            
            Screen.MousePointer = 11
            If Rpt.DefinirArquivo(Bdados, App.Path & "\TFC_ResumoImovel.rpt") Then
                Rpt.Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
                Rpt.Selecao = Right(CondicaoRelatorio, Len(CondicaoRelatorio) - 5)
                Rpt.Arvore = False
                Rpt.Visualizar
            End If
            Screen.MousePointer = 0
        Case "Sai&r"
            Unload Me
    End Select
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    Grid.Preencher Bdados, ""
    txtIc.SetFocus
End Sub

'Private Sub Form_DblClick()
'    Dim sql As String
'    Dim Rs As VSRecordset
'
'    sql = "Select IC,tlg_cod_logradouro from vis_gb_pho"
'    If Bdados.AbreTabela(sql, Rs) Then
'        Rs.MoveFirst
'        lbl(0) = "0"
'        Do
'            Bdados.AtualizaDados "TAB_IMOVEL", Bdados.PreparaValor(Rs!tlg_cod_logradouro), "tim_tlg_cod_logradouro", "tim_ic ='" & Rs!Ic & "'"
'            lbl(0) = lbl(0) + 1
'            DoEvents
'            Rs.MoveNext
'        Loop While Not Rs.EOF
'    End If
'End Sub

Private Sub Form_Load()
    Dim Controle As Control
    Dim i As Byte
    'setando
    Set cadastro = New VSImposto
    Set Imovel = New cImovel
    Set Endereco = New cEndereco
    'preenchendo combos
    Endereco.PreencherComboBairro cboBairro
    Endereco.PreencherComboTipoLogr cboTipoLogr
    Endereco.PreencherComboLogr cboLogr
    cboAforado.PreencherGeral Bdados, "SIM OU NÃO"
    
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    AtualizaCabecalho Grid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cadastro = Nothing
    Set Imovel = Nothing
    Set Endereco = Nothing
End Sub

Private Sub grid_DblClick()
    If Not Grid.SelectedItem Is Nothing Then
        TCIP102.Tag = Grid.SelectedItem
        TCIP102.Show vbModal
    End If
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Grid.SelectedItem Is Nothing Then Exit Sub
    If Button = 2 Then
        mnuIPTU.Caption = "Gerar IPTU de " & Grid.SelectedItem
        Me.PopupMenu mnuPrinc
    End If
End Sub
