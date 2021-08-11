VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TAGD101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAGD101"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   38
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TAGD101.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   37
      Top             =   6330
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   8580
         TabIndex        =   9
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar  "
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   7395
         TabIndex        =   8
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9765
         TabIndex        =   10
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   5670
      Left            =   45
      TabIndex        =   13
      Top             =   615
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   10001
      Caption         =   "Contribuinte, Gráfica e Notas Fiscais"
      Descricao       =   "Informações gerais do contribuinte e da gráfica"
      corFaixa        =   32768
      Icone           =   "TAGD101.frx":2123
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.fraVISUAL fraVISUAL2 
         Height          =   705
         Left            =   90
         TabIndex        =   39
         Top             =   690
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   1244
         Altura          =   1905
         Caption         =   " Documento para Impressão"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.cboVISUAL cboDoc 
            Height          =   315
            Left            =   195
            TabIndex        =   40
            Top             =   330
            Width           =   10305
            _ExtentX        =   18177
            _ExtentY        =   556
            Caption         =   "Documento"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Enabled         =   0   'False
         End
      End
      Begin VTOcx.fraVISUAL fraVISUAL1 
         Height          =   960
         Left            =   90
         TabIndex        =   34
         Top             =   4650
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   1693
         Altura          =   1905
         Caption         =   " Sequencia de Notas"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.cboVISUAL cboEspecie 
            Height          =   510
            Left            =   105
            TabIndex        =   4
            Top             =   330
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   900
            Caption         =   "Espécie"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
         End
         Begin VTOcx.cboVISUAL cboSerie 
            Height          =   510
            Left            =   2892
            TabIndex        =   5
            Top             =   330
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   900
            Caption         =   "Série/Sub-Série"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
         End
         Begin VTOcx.cboVISUAL cboQuant 
            Height          =   510
            Left            =   4404
            TabIndex        =   6
            Top             =   330
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   900
            Caption         =   "Notas p/ Bloco"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
         End
         Begin VTOcx.txtVISUAL txtBlocos 
            Height          =   480
            Left            =   5916
            TabIndex        =   7
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   847
            Caption         =   "Total de Blocos"
            Text            =   ""
            Restricao       =   2
            AlinhamentoRotulo=   1
            MaxLen          =   8
         End
         Begin VTOcx.txtVISUAL txtInicio 
            Height          =   480
            Left            =   7563
            TabIndex        =   36
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   847
            Caption         =   "Nota Inicial"
            Text            =   ""
            AlinhamentoRotulo=   1
            MaxLen          =   8
         End
         Begin VTOcx.txtVISUAL txtFim 
            Height          =   480
            Left            =   9210
            TabIndex        =   35
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   847
            Caption         =   "Nota Final"
            Text            =   ""
            AlinhamentoRotulo=   1
         End
      End
      Begin VTOcx.fraVISUAL fra 
         Height          =   3255
         Index           =   0
         Left            =   5490
         TabIndex        =   24
         Top             =   1380
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   5741
         Altura          =   1905
         Caption         =   " Estabelecimento Gráfico"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.cmdVISUAL cmdBuscarGrafica 
            Height          =   330
            Left            =   1665
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   465
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   582
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtCidadeGrafica 
            Height          =   480
            Left            =   75
            TabIndex        =   33
            Top             =   2700
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   847
            Caption         =   "Cidade"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtCompGrafica 
            Height          =   480
            Left            =   75
            TabIndex        =   32
            Top             =   1740
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   847
            Caption         =   "Complemento"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtCgcGrafica 
            Height          =   480
            Left            =   2070
            TabIndex        =   31
            Top             =   300
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   847
            Caption         =   "CNPJ"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtTipoLogrGrafica 
            Height          =   480
            Left            =   75
            TabIndex        =   30
            Top             =   1260
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   847
            Caption         =   "Logradouro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNomeGrafica 
            Height          =   480
            Left            =   75
            TabIndex        =   29
            Top             =   780
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   847
            Caption         =   "Nome"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNumeroGrafica 
            Height          =   480
            Left            =   4680
            TabIndex        =   28
            Top             =   1260
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            Caption         =   "Nº"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtLogrGrafica 
            Height          =   480
            Left            =   1425
            TabIndex        =   27
            Top             =   1260
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   847
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtUFGrafica 
            Height          =   480
            Left            =   4680
            TabIndex        =   26
            Top             =   2700
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            Caption         =   "UF"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtBairroGrafica 
            Height          =   480
            Left            =   75
            TabIndex        =   25
            Top             =   2220
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   847
            Caption         =   "Bairro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtImGrafica 
            Height          =   480
            Left            =   75
            TabIndex        =   2
            Tag             =   "Insc. Municipal"
            Top             =   300
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   847
            Caption         =   "Insc. Municipal"
            Text            =   ""
            Restricao       =   2
            AlinhamentoRotulo=   1
            RetirarMascara  =   0   'False
         End
      End
      Begin VTOcx.fraVISUAL fra 
         Height          =   3255
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   1380
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   5741
         Altura          =   1905
         Caption         =   " Contribuinte"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtLogr 
            Height          =   480
            Left            =   1425
            TabIndex        =   23
            Top             =   1260
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   847
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNumero 
            Height          =   480
            Left            =   4680
            TabIndex        =   22
            Top             =   1260
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            Caption         =   "Nº"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNomeContrib 
            Height          =   480
            Left            =   75
            TabIndex        =   21
            Top             =   780
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   847
            Caption         =   "Nome"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.cmdVISUAL cmdBuscarContr 
            Height          =   330
            Left            =   1500
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   450
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   582
            Caption         =   ""
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtTipoLogr 
            Height          =   480
            Left            =   75
            TabIndex        =   20
            Top             =   1260
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   847
            Caption         =   "Logradouro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtUF 
            Height          =   480
            Left            =   4680
            TabIndex        =   19
            Top             =   2700
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            Caption         =   "UF"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtCgc 
            Height          =   480
            Left            =   1890
            TabIndex        =   18
            Top             =   300
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   847
            Caption         =   "CNPJ"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtComplemento 
            Height          =   480
            Left            =   75
            TabIndex        =   17
            Top             =   1740
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   847
            Caption         =   "Complemento"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtBairro 
            Height          =   480
            Left            =   75
            TabIndex        =   16
            Top             =   2220
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   847
            Caption         =   "Bairro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtCidade 
            Height          =   480
            Left            =   75
            TabIndex        =   15
            Top             =   2700
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   847
            Caption         =   "Cidade"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtIm 
            Height          =   480
            Left            =   75
            TabIndex        =   0
            Tag             =   "Insc. Municipal"
            Top             =   300
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   847
            Caption         =   "Insc. Municipal"
            Text            =   ""
            Restricao       =   2
            AlinhamentoRotulo=   1
            RetirarMascara  =   0   'False
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   1138
      Icone           =   "TAGD101.frx":2DFD
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2070
      TabIndex        =   11
      Top             =   165
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "TAGD101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim Funcoes_Aidf    As AIDF
Dim Contribuinte As cContribuinte
Dim NotaAidf As cNotaAidf
Dim Grafica As cGraficaAidf
Dim NumGreden As String

'Sub ImprimeAidf(NumAidf As Double)
'    Dim a As Byte
'
'    With Rpt
'         If .DefinirArquivo(Bdados, App.Path + "\TAIDF.rpt") Then
'            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
'            .Formulas "Sequencia", Format(txtInicio, "0000") & " a " & Format(txtFim, "0000")
'            .Formulas "Blocos", Format(txtBlocos, "00")
'            .Formulas "Cidade", Aplicacoes.Municipio
'            '.Selecao = "{Tab_Aidf.tai_num_aidf} =" & NumAidf
'            .Formulas "NumAidf", CStr(NumAidf)
'            .Titulo = "Autorização de Impressão de Documentos Fiscais"
'            .Arvore = False
'            .Visualizar
'         End If
'    End With
'    Set Rpt = Nothing
'
'End Sub

Private Sub cmdBuscarContr_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdBuscarGrafica_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtImGrafica
End Sub

Private Sub cmdEnter_Click()
        SendKeys "{Tab}"
End Sub

Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    txtIm.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim valores As String
    Dim campos As String
    Dim Sql As String
    Dim rs As VSRecordset
    Dim condicao As String
    Dim ClsAidf As New AIDF
    Dim NumAidf As Double
    
    If cboDoc.ListIndex = -1 Then
        Avisa "Informe o Documento"
        Exit Sub
    End If
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = vbHourglass
    NumAidf = Grafica.GeraNumero(AIDG)
    If Trim(txtIm) = Trim(txtImGrafica) Then
        Util.Avisa "Contribuinte não pode ser igual a gráfica."
        Exit Sub
    End If
    With NotaAidf
        .NumAidf = NumAidf
        .ImContribuinte = txtIm
        .ImGrafica = txtImGrafica
        .DataAutorizacao = Date
        .NotaInicial = txtInicio
        .NotaFinal = txtFim
        .TotalBlocos = txtBlocos
        .SituacaoAidf = 1
        .CodUsuario = Aplicacoes.Usuario
        .Serie = cboSerie.Coluna(0).VALOR
        .NomeSerie = cboSerie.Coluna(0).VALOR
        .Especie = cboEspecie.Coluna(0).VALOR
        .TipoAidf = cboEspecie.Coluna(1).VALOR
        .Documento = cboDoc.Coluna(1).VALOR
        If .GravarNota Then
            Dim Pos As Integer
            Pos = InStr(cboDoc.Text, " - ")
            Util.Informa Left(cboDoc.Text, Pos - 1) & " emitida com sucesso " & vbCrLf & "Nº gerado: " & NumAidf & "."
            If Util.Confirma("Deseja imprimir AIDG " & NumAidf & "" & "?") Then
                .Imprimir NumAidf
            End If
            Edita.LimpaCampos Me
            Screen.MousePointer = vbNormal
        End If
    End With
End Sub

Private Sub Form_Activate()
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
        txtImGrafica.Formato = formNenhum
    End If
    cboDoc.Preencher Bdados, "select * from vis_tipo_impressão_doc"
    cboDoc.SetarLinha 2, 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = vbKeyTab
End Sub

Private Sub Form_Load()
    '******setando classe
    Set NotaAidf = New cNotaAidf
    Set Contribuinte = New cContribuinte
    Set Funcoes_Aidf = New AIDF
    Set Grafica = New cGraficaAidf
    
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision

    '******preenchendo combo
    With NotaAidf
        .PreencherCboEspecie cboEspecie
        .PreencherCboQtd cboQuant
        .PreencherCboSerie cboSerie
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Funcoes_Aidf = Nothing
    Set Contribuinte = Nothing
    Set NotaAidf = Nothing
    Set Grafica = Nothing
End Sub



Private Sub txtBlocos_Change()
    Dim UltimaNot As Integer
    If Me.ActiveControl.Name = "cmdSair" Then Exit Sub
    If Trim(txtBlocos) = "" Or cboEspecie.ListIndex = -1 Or Trim(txtIm) = "" Then Exit Sub
    If cboQuant.ListIndex = -1 Then
        Util.Informa ("Informe quantidade de notas por bloco.")
        cboQuant.SetFocus
        Exit Sub
    End If
    If NotaAidf.UltimaNota(txtIm, CStr(cboSerie.Coluna(0).VALOR), CStr(cboEspecie.Coluna(1).VALOR), CStr(cboDoc.Coluna(1).VALOR)) > 0 Then
        txtInicio.Enabled = False
    Else
        txtInicio.Enabled = True
    End If
    txtInicio = NotaAidf.UltimaNota(txtIm, CStr(cboSerie.Coluna(0).VALOR), CStr(cboEspecie.Coluna(1).VALOR), CStr(cboDoc.Coluna(1).VALOR)) + 1
    txtFim = CDbl(Nvl(txtInicio, 0)) + CDbl((Nvl(cboQuant, 0)) * CDbl(Nvl(txtBlocos, 0))) - 1
End Sub

Private Sub txtIm_LostFocus()
    Dim NomeContrib As String, TipoLogrContr As String, LogrContr As String, NumeroContr As String, CompContri As String, _
          BairroContr As String, CepContr As String, MunicContr As String, UFContr As String, DocumentoContr As String
    
    If Trim(txtIm) = "" Then LimpaCamposContribuinte: Exit Sub
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtIm = Imposto.FormataInscricao(txtIm, InscContrib)
    End If
    With Contribuinte
        If .BuscarContribuinte(txtIm, NomeContrib, TipoLogrContr, LogrContr, NumeroContr, CompContri, _
            BairroContr, CepContr, MunicContr, UFContr, DocumentoContr) Then
            If DocumentoContr <> "" Then
                txtNomeContrib = NomeContrib
                txtCgc = DocumentoContr
                txtTipoLogr = TipoLogrContr
                txtLogr = LogrContr
                txtNumero = NumeroContr
                txtComplemento = CompContri
                txtBairro = BairroContr
                txtCidade = MunicContr
                txtUF = UFContr
            Else
                Avisa "AIDF não pode ser emitida, Verifique se os dados cadastrais estão corretos." & vbCrLf & "EX: CPF - CNPJ - ATIVIDADE ECONOMICA"
                LimpaCamposContribuinte
                txtIm = ""
                txtIm.SetFocus
            End If
        Else
            Avisa "Contribuinte não Cadastrado."
            LimpaCamposContribuinte
            txtIm.SetFocus
        End If
    End With
End Sub
Private Sub LimpaCamposContribuinte()
    txtNomeContrib = ""
    txtCgc = ""
    txtTipoLogr = ""
    txtLogr = ""
    txtNumero = ""
    txtComplemento = ""
    txtBairro = ""
    txtCidade = ""
    txtUF = ""
End Sub
Private Sub LimpaCamposGrafica()
    txtNomeGrafica = ""
    txtCgcGrafica = ""
    txtTipoLogrGrafica = ""
    txtLogrGrafica = ""
    txtNumeroGrafica = ""
    txtCompGrafica = ""
    txtBairroGrafica = ""
    txtCidadeGrafica = ""
    txtUFGrafica = ""
End Sub

Private Sub txtImGrafica_LostFocus()
Dim NomeGraf As String, TipoLogrGraf As String, LogrGraf As String, NumeroGraf As String, CompGraf As String, _
          BairroGraf As String, CepGraf As String, MunicGraf As String, UFGraf As String, DocumentoGraf As String
    If Trim(txtImGrafica) = "" Then LimpaCamposGrafica:   Exit Sub
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtImGrafica = Imposto.FormataInscricao(txtImGrafica, InscContrib)
    End If
    If Grafica.Buscar(, txtImGrafica) Then
        txtImGrafica = Grafica.Im
        If Grafica.Situacao = 1 Then Util.Avisa "Gráfica descredenciada.": txtImGrafica.SetFocus: Exit Sub
        If Contribuinte.BuscarContribuinte(txtImGrafica, NomeGraf, TipoLogrGraf, LogrGraf, NumeroGraf, CompGraf, _
            BairroGraf, CepGraf, MunicGraf, UFGraf, DocumentoGraf) Then
            txtNomeGrafica = NomeGraf
            txtCgcGrafica = DocumentoGraf
            txtTipoLogrGrafica = TipoLogrGraf
            txtLogrGrafica = LogrGraf
            txtNumeroGrafica = NumeroGraf
            txtCompGrafica = CompGraf
            txtBairroGrafica = BairroGraf
            txtCidadeGrafica = MunicGraf
            txtUFGrafica = UFGraf
            If AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
            If Checa_Obrigacao1_2(txtImGrafica) Then
                If Not Util.Confirma("Contribuinte com obrigações em aberto. Deseja continuar o processo?") Then
                    LimpaCamposGrafica
                    txtImGrafica.SetFocus
                End If
            End If
            End If
        Else
            Util.Avisa "Contribuite não encontrado"
            LimpaCamposContribuinte
            txtImGrafica.SetFocus
        End If
    Else
        Util.Avisa "Estabelecimento não é credenciamento para esta operacão."
         LimpaCamposGrafica
        txtImGrafica.SetFocus
    End If
End Sub

Private Sub txtInicio_Change()
    If Trim(txtInicio) = "" Or cboQuant = "" Or Trim(txtInicio) = "" Then: txtFim = "": Exit Sub
    txtFim = (CDbl(Nvl(cboQuant, 0)) * CDbl(Nvl(txtBlocos, 0))) + (CDbl(txtInicio) - 1)
End Sub

