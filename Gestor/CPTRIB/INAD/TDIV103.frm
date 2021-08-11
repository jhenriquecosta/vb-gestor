VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TDIV103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciamento de Gráficas"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDocumento 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1200
      Left            =   1035
      TabIndex        =   20
      Top             =   5115
      Width           =   9825
      Begin VB.Label LblRegistros 
         BackColor       =   &H8000000C&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   2550
         TabIndex        =   26
         Top             =   795
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Total de Registros:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   75
         TabIndex        =   25
         Top             =   765
         Width           =   2430
      End
      Begin VB.Label LblPercentual 
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   2550
         TabIndex        =   24
         Top             =   165
         Width           =   945
      End
      Begin VB.Label LblDocumento 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   2550
         TabIndex        =   23
         Top             =   510
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Documento :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   825
         TabIndex        =   22
         Top             =   435
         Width           =   1680
      End
      Begin VB.Label LblGeracao 
         BackColor       =   &H8000000C&
         Caption         =   "Gerando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   1095
         TabIndex        =   21
         Top             =   150
         Width           =   1410
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TDIV103.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   14
      Top             =   7485
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1085
      Modulo          =   "Divida Ativa"
      Begin VTOcx.cmdVISUAL cmdDivida 
         Height          =   375
         Left            =   7740
         TabIndex        =   8
         Top             =   150
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         Caption         =   "&Gerar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   9240
         TabIndex        =   9
         Top             =   150
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10470
         TabIndex        =   10
         Top             =   150
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
      TabIndex        =   13
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1138
      Icone           =   "TDIV103.frx":2123
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2940
      Index           =   3
      Left            =   0
      TabIndex        =   16
      Top             =   570
      Width           =   11580
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   870
         TabIndex        =   0
         Top             =   150
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   705
         TabIndex        =   2
         Top             =   900
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   529
         Caption         =   "Inscrição"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   315
         Left            =   510
         TabIndex        =   1
         Top             =   525
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   556
         Caption         =   "Documento"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   690
         TabIndex        =   11
         Top             =   1620
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   315
         Left            =   3660
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   900
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtProcesso 
         Height          =   510
         Left            =   1500
         TabIndex        =   7
         Top             =   2340
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   900
         Caption         =   "Termo de Lancamento"
         Text            =   ""
         Requerido       =   0   'False
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VTOcx.txtVISUAL txtPeriodoFinal 
         Height          =   300
         Left            =   3930
         TabIndex        =   5
         Top             =   2010
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   529
         Caption         =   "Periodo Final"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MaxLen          =   4
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   300
         Left            =   270
         TabIndex        =   4
         Top             =   2010
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         Caption         =   "Periodo Inicial"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MaxLen          =   4
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   315
         Left            =   960
         TabIndex        =   27
         Top             =   1245
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   556
         Caption         =   "Razão"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   7650
         TabIndex        =   3
         Top             =   915
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   529
         Caption         =   "Cadastro do Imóvel"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   315
         Left            =   11130
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   900
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtExercicio 
         Height          =   300
         Left            =   8970
         TabIndex        =   6
         Top             =   2010
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         Caption         =   "Exercicio"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MaxLen          =   4
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   18
         Top             =   1590
         Width           =   45
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   12
      Top             =   690
      Width           =   375
   End
   Begin VTOcx.grdVISUAL grdDivida 
      Height          =   3510
      Left            =   0
      TabIndex        =   19
      Top             =   4245
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   6191
      OcultarRodape   =   -1  'True
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Index           =   0
      Left            =   0
      TabIndex        =   29
      Top             =   3435
      Width           =   11580
      Begin VTOcx.txtVISUAL txtAutoridade 
         Height          =   300
         Left            =   525
         TabIndex        =   31
         Top             =   135
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   529
         Caption         =   "Autoridade"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtCargo 
         Height          =   300
         Left            =   930
         TabIndex        =   32
         Top             =   450
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   529
         Caption         =   "Cargo"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   30
         Top             =   1590
         Width           =   45
      End
   End
End
Attribute VB_Name = "TDIV103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim Obrig As New Obrigacao
Dim DAT As New cDividaAtiva
Private Sub cmdDivida_Click()
    FraDocumento.Visible = True
    LblDocumento = cboTipo.Text
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Me.MousePointer = vbHourglass
    Bdados.Conexao.DBConnection.CommandTimeout = 999999
    
    
    If Trim(txtIm) <> "" Then
        If DAT.GeraDocumentacao(Nvl(CStr(cboTipo.Coluna(0).Valor), 0), txtPeriodoInicial, txtPeriodoFinal, txtExercicio, txtIm, InscContrib, CStr(cboImposto.Coluna(0).Valor), txtProcesso.Text, , , , LblPercentual, LblRegistros, edTributaria) Then
            DAT.CarregaDividaGerada grdDivida, txtIm, txtExercicio, txtPeriodoInicial, txtPeriodoFinal, CStr(cboImposto.Coluna(0).Valor), , , , , InscContrib, edTributaria
            Informa "Registro(s) gravado com sucesso."
        Else
            Util.Avisa "Não existe tributos vencidos para o(s) contribuinte(s) ou documento anterior não gerado."
        End If
    Else
        If DAT.GeraDocumentacao(Nvl(CStr(cboTipo.Coluna(0).Valor), 0), txtPeriodoInicial, txtPeriodoFinal, txtExercicio, txtImovel, InscImovel, CStr(cboImposto.Coluna(0).Valor), txtProcesso.Text, , , , LblPercentual, LblRegistros, edTributaria) Then
            DAT.CarregaDividaGerada grdDivida, txtImovel, txtExercicio, txtPeriodoInicial, txtPeriodoFinal, CStr(cboImposto.Coluna(0).Valor), , , , , InscImovel, edTributaria
            Informa "Registro(s) gravado com sucesso."
        Else
            Util.Avisa "Não existe tributos vencidos para o(s) contribuinte(s) ou documento anterior não gerado."
        End If
    End If
    FraDocumento.Visible = False
    Me.MousePointer = 0
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, Me.txtIm, txtRazao
    txtIm_LostFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
    txtImovel_LostFocus
End Sub

Private Sub Form_Activate()
    
    txtCargo = DAT.BuscaParametro("CARGO AUTORIDADE", edTributaria)
    txtAutoridade = DAT.BuscaParametro("AUTORIDADE COMPETENTE", edTributaria)
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    cboTipo.Preencher Bdados, "SELECT TGE_CODIGO,TGE_NOME FROM VIS_DOCUMENTOS_DAT ORDER BY TGE_CODIGO", 1
    Obrig.PreencheComboTributo cboImposto, True, etcTributario
    FraDocumento.Visible = False
    
End Sub

Private Sub txtIm_LostFocus()
    Dim Ic As String
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
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
