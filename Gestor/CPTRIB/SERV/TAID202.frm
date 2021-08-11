VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TAID202 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   41
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TAID202.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   37
      Top             =   7140
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   8610
         TabIndex        =   2
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   7425
         TabIndex        =   1
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Cancelar"
         Acao            =   9
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9795
         TabIndex        =   3
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
      Height          =   6465
      Left            =   45
      TabIndex        =   8
      Top             =   645
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   11404
      Caption         =   "Contribuinte, Gráfica e Notas Fiscais"
      Descricao       =   "Informações gerais do contribuinte e da gráfica"
      corFaixa        =   32768
      Icone           =   "TAID202.frx":2123
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.cmdVISUAL CmdConsulta 
         Height          =   330
         Left            =   2310
         TabIndex        =   40
         Top             =   705
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   582
         Caption         =   "Consultar AIDF"
         Acao            =   5
      End
      Begin VB.TextBox txtMotivo 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Text            =   "TAID202.frx":243D
         Top             =   5535
         Width           =   10710
      End
      Begin VTOcx.txtVISUAL txtNumAidf 
         Height          =   285
         Left            =   105
         TabIndex        =   0
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         Caption         =   "Nº AIDF"
         Text            =   ""
      End
      Begin VTOcx.fraVISUAL fra 
         Height          =   3255
         Index           =   1
         Left            =   90
         TabIndex        =   26
         Top             =   1050
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   5741
         Altura          =   1905
         Caption         =   " Contribuinte"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtIm 
            Height          =   480
            Left            =   75
            TabIndex        =   36
            Tag             =   "Insc. Municipal"
            Top             =   300
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   847
            Caption         =   "Insc. Municipal"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoRotulo=   1
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtCidade 
            Height          =   480
            Left            =   75
            TabIndex        =   35
            Top             =   2700
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   847
            Caption         =   "Cidade"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtBairro 
            Height          =   480
            Left            =   75
            TabIndex        =   34
            Top             =   2220
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   847
            Caption         =   "Bairro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtComplemento 
            Height          =   480
            Left            =   75
            TabIndex        =   33
            Top             =   1740
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   847
            Caption         =   "Complemento"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtCgc 
            Height          =   480
            Left            =   1890
            TabIndex        =   32
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
         Begin VTOcx.txtVISUAL txtUF 
            Height          =   480
            Left            =   4680
            TabIndex        =   31
            Top             =   2700
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            Caption         =   "UF"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtTipoLogr 
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
         Begin VTOcx.txtVISUAL txtNomeContrib 
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
         Begin VTOcx.txtVISUAL txtNumero 
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
         Begin VTOcx.txtVISUAL txtLogr 
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
      End
      Begin VTOcx.fraVISUAL fra 
         Height          =   3255
         Index           =   0
         Left            =   5490
         TabIndex        =   15
         Top             =   1065
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   5741
         Altura          =   1905
         Caption         =   " Estabelecimento Gráfico"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtImGrafica 
            Height          =   480
            Left            =   75
            TabIndex        =   25
            Tag             =   "Insc. Municipal"
            Top             =   300
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   847
            Caption         =   "Insc. Municipal"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoRotulo=   1
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtBairroGrafica 
            Height          =   480
            Left            =   75
            TabIndex        =   24
            Top             =   2220
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   847
            Caption         =   "Bairro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtUFGrafica 
            Height          =   480
            Left            =   4680
            TabIndex        =   23
            Top             =   2700
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            Caption         =   "UF"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtLogrGrafica 
            Height          =   480
            Left            =   1425
            TabIndex        =   22
            Top             =   1260
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   847
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNumeroGrafica 
            Height          =   480
            Left            =   4680
            TabIndex        =   21
            Top             =   1260
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            Caption         =   "Nº"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNomeGrafica 
            Height          =   480
            Left            =   75
            TabIndex        =   20
            Top             =   780
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   847
            Caption         =   "Nome"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtTipoLogrGrafica 
            Height          =   480
            Left            =   75
            TabIndex        =   19
            Top             =   1260
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   847
            Caption         =   "Logradouro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtCgcGrafica 
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
            Formato         =   2
            AlinhamentoRotulo=   1
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtCompGrafica 
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
         Begin VTOcx.txtVISUAL txtCidadeGrafica 
            Height          =   480
            Left            =   75
            TabIndex        =   16
            Top             =   2700
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   847
            Caption         =   "Cidade"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
      End
      Begin VTOcx.fraVISUAL fraVISUAL1 
         Height          =   960
         Left            =   90
         TabIndex        =   9
         Top             =   4350
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   1693
         Altura          =   1905
         Caption         =   " Sequencia de Notas"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Enabled         =   0   'False
         Begin VTOcx.txtVISUAL txtNotaBloco 
            Height          =   480
            Left            =   4314
            TabIndex        =   7
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   847
            Caption         =   "Notas p/ Bloco"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtFim 
            Height          =   480
            Left            =   9210
            TabIndex        =   14
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   847
            Caption         =   "Nota Final"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtInicio 
            Height          =   480
            Left            =   7563
            TabIndex        =   13
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   847
            Caption         =   "Nota Inicial"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtBlocos 
            Height          =   480
            Left            =   5916
            TabIndex        =   12
            Top             =   360
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   847
            Caption         =   "Total de Blocos"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.cboVISUAL cboSerie 
            Height          =   510
            Left            =   2847
            TabIndex        =   11
            Top             =   330
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   900
            Caption         =   "Série/Sub-Série"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
            Enabled         =   0   'False
         End
         Begin VTOcx.cboVISUAL cboEspecie 
            Height          =   510
            Left            =   105
            TabIndex        =   10
            Top             =   330
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   900
            Caption         =   "Espécie"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
            Enabled         =   0   'False
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Motivo do Cancelamento"
         Height          =   195
         Left            =   105
         TabIndex        =   39
         Top             =   5340
         Width           =   1770
      End
   End
   Begin VB.PictureBox lbl 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   13
      Left            =   255
      ScaleHeight     =   180
      ScaleWidth      =   825
      TabIndex        =   4
      Top             =   765
      Width           =   885
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TAID202.frx":2465
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "TAID202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim Contribuinte As cContribuinte
Dim Grafica As cGraficaAidf
Dim NotaAidf As cNotaAidf


Private Sub CmdConsulta_Click()
    Load TAID501
    TAID501.Tag = Me.Name
    TipoConsulta = 1
    TAID501.Show 1
    txtNumAidf = Inscri
    txtNumAidf_LostFocus
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtNumAidf.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Motivo As String
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    
    If NotaAidf.Buscar(txtNumAidf) Then
       If NotaAidf.SituacaoAidf = 2 Then
        Util.Avisa "AIDF  já Cancelada."
        Exit Sub
       End If
    End If
    If Util.Confirma("Deseja Cancelar?") Then
        With NotaAidf
            .NumAidf = txtNumAidf
            .DataCancelamento = Date
            .SituacaoAidf = 2
            .CodUsuario = Aplicacoes.Usuario
            .Motivo = txtMotivo
            If .CancelarAidf Then
                Util.Avisa "Transação Realizada com Sucesso."
                cmdLimpar_Click
            Else
                Util.Mensagem "Erro na gravação"
            End If
        End With
    End If
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    Set Grafica = New cGraficaAidf
    Set Contribuinte = New cContribuinte
    Set NotaAidf = New cNotaAidf
    
    NotaAidf.PreencherCboEspecie cboEspecie
    NotaAidf.PreencherCboSerie cboSerie
    If AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
       txtIm.Formato = formNenhum
       txtImGrafica.Formato = formNenhum
    End If
    txtMotivo = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Grafica = Nothing
    Set Contribuinte = Nothing
    Set NotaAidf = Nothing
End Sub

Private Sub txtFim_LostFocus()
    If Trim(txtInicio) <> "" And Trim(txtFim) <> "" Then
        If ((CDbl(txtFim) - CDbl(txtInicio)) / 50) < 1 Then
            Util.Avisa "Quantidade insuficiente de notas."
            txtFim = ""
            Exit Sub
        End If
        txtBlocos = Fix((CDbl(txtFim) - CDbl(txtInicio) + 1) / 50)
    End If
End Sub

Private Sub txtIm_LostFocus()
    Dim NomeContrib As String, TipoLogrContr As String, LogrContr As String, NumeroContr As String, CompContri As String, _
          BairroContr As String, CepContr As String, MunicContr As String, UFContr As String, DocumentoContr As String
    
    If Trim(txtIm) = "" Then Exit Sub
    With Contribuinte
        If .BuscarContribuinte(txtIm, NomeContrib, TipoLogrContr, LogrContr, NumeroContr, CompContri, _
            BairroContr, CepContr, MunicContr, UFContr, DocumentoContr) Then
                txtNomeContrib = NomeContrib
                txtCgc = DocumentoContr
                txtTipoLogr = TipoLogrContr
                txtLogr = LogrContr
                txtNumero = NumeroContr
                txtComplemento = CompContri
                txtBairro = BairroContr
                txtCidade = MunicContr
                txtUF = UFContr
        End If
    End With
End Sub

Private Sub txtImGrafica_LostFocus()
    Dim NomeGraf As String, TipoLogrGraf As String, LogrGraf As String, NumeroGraf As String, CompGraf As String, _
          BairroGraf As String, CepGraf As String, MunicGraf As String, UFGraf As String, DocumentoGraf As String
    If Trim(txtImGrafica) = "" Then Exit Sub
    With Grafica
        LimpaCamposGrafica
        If Contribuinte.BuscarContribuinte(txtIm, NomeGraf, TipoLogrGraf, LogrGraf, NumeroGraf, CompGraf, _
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
        End If
    End With
End Sub

Private Sub txtInicio_LostFocus()
    If Trim(txtInicio) <> "" And Trim(txtFim) <> "" Then
        If ((CDbl(txtFim) - CDbl(txtInicio)) / 50) < 1 Then
            Util.Avisa "Quantidade insuficiente de notas."
            txtInicio = ""
            Exit Sub
        End If
        txtBlocos = Fix((CDbl(txtFim) - CDbl(txtInicio) + 1) / 50)
    End If
End Sub

Private Sub txtNumAidf_LostFocus()
    On Error Resume Next
    If txtNumAidf = "" Then Exit Sub
    With NotaAidf
           LimpaTodosCampos
        If .Buscar(txtNumAidf) Then
'            If .SituacaoAidf = 2 Then
'                Util.Avisa "AIDF  já Cancelada."
'                Exit Sub
'            End If
            txtIm = .ImContribuinte
            txtIm_LostFocus
            txtImGrafica = .ImGrafica
            txtImGrafica_LostFocus
            txtInicio = .NotaInicial
            txtFim = .NotaFinal
            txtBlocos = .TotalBlocos
            cboEspecie.SetarLinha .TipoAidf, 1
            cboSerie.SetarLinha .Serie, 1
            txtNotaBloco = ((txtFim + 1) - txtInicio) / txtBlocos
            txtMotivo = .Motivo
        Else
            Util.Avisa "Número de AIDF não encontrado."
            cmdLimpar_Click
        End If
    End With
End Sub
Private Sub LimpaTodosCampos()
        txtIm = ""
        txtImGrafica = ""
        txtInicio = ""
        txtFim = ""
        txtBlocos = ""
        cboEspecie.ListIndex = -1
        cboSerie.ListIndex = -1
        txtNotaBloco = ""
        LimpaCamposContribuinte
        LimpaCamposGrafica
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
