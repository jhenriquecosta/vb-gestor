VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TAID101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   22
      Top             =   15
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TAID101.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   17
      Top             =   4200
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   4470
         TabIndex        =   12
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
         Left            =   5625
         TabIndex        =   13
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
         Left            =   6780
         TabIndex        =   14
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
      Height          =   3435
      Left            =   60
      TabIndex        =   18
      Top             =   720
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   6059
      Caption         =   "Credenciamento de Gráficas"
      Descricao       =   "Credencia gráficas no sistema"
      corFaixa        =   32768
      Icone           =   "TAID101.frx":2123
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.fraVISUAL fra 
         Height          =   705
         Index           =   0
         Left            =   2580
         TabIndex        =   20
         Top             =   2655
         Visible         =   0   'False
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1244
         Altura          =   1905
         Caption         =   " Informações do Credenciamento"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtValidade 
            Height          =   285
            Left            =   2880
            TabIndex        =   11
            Top             =   345
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   503
            Caption         =   "Validade"
            Text            =   ""
            Formato         =   0
         End
         Begin VTOcx.txtVISUAL txtInicio 
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   345
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   503
            Caption         =   "Data Inicial"
            Text            =   ""
            Formato         =   0
         End
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   510
         Left            =   90
         TabIndex        =   9
         Tag             =   "Tipo"
         Top             =   2820
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   900
         Caption         =   "Tipo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.fraVISUAL fra 
         Height          =   1890
         Index           =   1
         Left            =   75
         TabIndex        =   19
         Top             =   735
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   3334
         Altura          =   1905
         Caption         =   " Informações Gerais"
         CorTexto        =   16777215
         CorFaixa        =   32768
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtLogr 
            Height          =   480
            Left            =   1440
            TabIndex        =   21
            Top             =   825
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   847
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtBairro 
            Height          =   480
            Left            =   90
            TabIndex        =   6
            Top             =   1320
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   847
            Caption         =   "Bairro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtComplemento 
            Height          =   480
            Left            =   4815
            TabIndex        =   5
            Top             =   825
            Width           =   2820
            _ExtentX        =   4974
            _ExtentY        =   847
            Caption         =   "Complemento"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNumero 
            Height          =   480
            Left            =   4170
            TabIndex        =   4
            Top             =   825
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            Caption         =   "Nº"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtIm 
            Height          =   480
            Left            =   90
            TabIndex        =   0
            Tag             =   "Insc. Municipal"
            Top             =   330
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   847
            Caption         =   "Insc. Municipal"
            Text            =   ""
            Restricao       =   2
            AlinhamentoRotulo=   1
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtNomeContrib 
            Height          =   480
            Left            =   1935
            TabIndex        =   2
            Top             =   315
            Width           =   5700
            _ExtentX        =   10054
            _ExtentY        =   847
            Caption         =   "Nome"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.cmdVISUAL cmdPesq 
            Height          =   330
            Index           =   1
            Left            =   1530
            TabIndex        =   1
            Top             =   495
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
            Left            =   90
            TabIndex        =   3
            Top             =   825
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   847
            Caption         =   "Logradouro"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtCidade 
            Height          =   480
            Left            =   3660
            TabIndex        =   7
            Top             =   1320
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   847
            Caption         =   "Cidade"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtUF 
            Height          =   480
            Left            =   6990
            TabIndex        =   8
            Top             =   1320
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            Caption         =   "UF"
            Text            =   ""
            Enabled         =   0   'False
            AlinhamentoRotulo=   1
         End
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   1138
      Icone           =   "TAID101.frx":38E5
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   225
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "TAID101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim Contribuinte As cContribuinte
Dim Grafica As cGraficaAidf

Private Sub cboTipo_Click()
    If cboTipo.ListIndex = 0 Then
        fra(0).Visible = True
        txtInicio = Format(Date, "dd/mm/yyyy")
        txtValidade = DateAdd("d", 365, txtInicio)
    ElseIf cboTipo.ListIndex = 1 Then
        fra(0).Visible = False
    End If
End Sub

Private Sub cmdEnter_Click()
        SendKeys "{Tab}"
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtIm.SetFocus
    fra(0).Visible = False
End Sub

Private Sub cmdPesq_Click(Index As Integer)
    AplicacoesVTFuncoes.BuscaNoEconomico TcoJuridica, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim rs As VSRecordset
    Dim Credenciamento As String
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Credenciamento = Grafica.VerificaCredenciamentoGrafica(txtIm, Credenciar)
    If Credenciamento <> "" Then
        If (cboTipo.Coluna(1).VALOR - 1) = 1 Then
            If Util.Confirma("Gráfica com credenciamento ainda válido até " & Credenciamento & ". Deseja Descredenciar?") Then
                 If Grafica.Descredenciar(txtIm) Then
                    Util.Informa "Gráfica descredenciada com sucesso."
                    cmdLimpar_Click
                End If
            End If
        Else
            Util.Avisa "Gráfica já credenciada, válida até  " & Credenciamento
        End If
    Else
        If (cboTipo.Coluna(1).VALOR - 1) = 0 Then
            If Grafica.Credenciar(txtIm, txtInicio, txtValidade) Then
                Util.Informa "Gráfica Credenciada com Sucesso."
                cmdLimpar_Click
            End If
        Else
            Util.Avisa "Gráfica não credenciada. Não é possivel discredenciar."
        End If
    End If
    Bdados.FechaTabela rs
End Sub

Private Sub Form_Load()

    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    '*********Setando Classes
    Set Contribuinte = New cContribuinte
    Set Grafica = New cGraficaAidf
    '*********Preenchendo combo
    Grafica.PreencherCboTipo cboTipo
    If AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtIm.Formato = formNenhum
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Contribuinte = Nothing
    Set Grafica = Nothing
End Sub

Private Sub txtIm_LostFocus()
    Dim Sql As String
    Dim rs As VSRecordset
    Dim NomeContrib As String, TipoLogrContr As String, LogrContr As String, NumeroContr As String, CompContri As String, _
          BairroContr As String, CepContr As String, MunicContr As String, Uf As String
    If Trim(txtIm) = "" Then Exit Sub
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        txtIm = Imposto.FormataInscricao(txtIm, InscContrib)
    End If

    If Contribuinte.BuscarContribuinte(txtIm, NomeContrib, TipoLogr, LogrContr, NumeroContr, CompContri, BairroContr, CepContr, MunicContr, Uf) Then
        txtNomeContrib = NomeContrib
        txtTipoLogr = TipoLogr
        txtLogr = LogrContr
        txtNumero = NumeroContr
        txtComplemento = CompContri
        txtBairro = BairroContr
        txtCidade = MunicContr
        txtUF = Uf
    Else
        Util.Avisa "Contribuinte não Cadastrado"
        cmdLimpar_Click
    End If
End Sub

