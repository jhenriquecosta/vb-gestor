VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIS109 
   Caption         =   "TCIS109"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grdVISUAL1 
      Height          =   3015
      Left            =   330
      TabIndex        =   11
      Top             =   2265
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   5318
   End
   Begin VTOcx.cmdVISUAL cmdBUsca 
      Height          =   315
      Left            =   3615
      TabIndex        =   10
      Top             =   855
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      Caption         =   ""
      Acao            =   5
   End
   Begin VTOcx.cboVISUAL cboAcao 
      Height          =   315
      Left            =   1575
      TabIndex        =   3
      Top             =   1875
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   556
      Caption         =   "Acao"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   9
      Top             =   5775
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   5745
         TabIndex        =   6
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6885
         TabIndex        =   7
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8025
         TabIndex        =   8
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4605
         TabIndex        =   5
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.txtVISUAL txtNomeContrib 
      Height          =   285
      Left            =   330
      TabIndex        =   1
      Top             =   1200
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   503
      Caption         =   "Nome/Razão Social"
      Text            =   ""
   End
   Begin VTOcx.txtVISUAL txtIm 
      Height          =   285
      Left            =   945
      TabIndex        =   0
      Top             =   870
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   503
      Caption         =   "Contribuinte"
      Text            =   ""
      Restricao       =   2
      AgruparValores  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtEndereco 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1530
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   503
      Caption         =   "Endereco"
      Text            =   ""
      Enabled         =   0   'False
   End
   Begin VTOcx.txtVISUAL txtDataEntrada 
      Height          =   285
      Left            =   6660
      TabIndex        =   4
      Top             =   1875
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   503
      Caption         =   "Dt Entrada"
      Text            =   ""
      Formato         =   0
      AgruparValores  =   0   'False
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   1138
      Icone           =   "TCIS109.frx":0000
   End
   Begin VB.Label LblSanitario 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2085
      TabIndex        =   15
      Top             =   5550
      Width           =   75
   End
   Begin VB.Label LblAlvara 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2085
      TabIndex        =   14
      Top             =   5280
      Width           =   75
   End
   Begin VB.Label LblOutras 
      AutoSize        =   -1  'True
      Caption         =   "SITUAÇÃO - TFS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   465
      TabIndex        =   13
      Top             =   5535
      Width           =   1470
   End
   Begin VB.Label LblNomeTFL 
      AutoSize        =   -1  'True
      Caption         =   "SITUAÇÃO - TFL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   5265
      Width           =   1665
   End
End
Attribute VB_Name = "TCIS109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rpt As New VSRelatorio
Public Processo As String

Private Sub cmdBusca_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdExcluir_Click()
    If grdVISUAL1.ListItems.Count >= 1 Then
        If Processo <> "" Then
            If Bdados.DeletaDados("TAB_MOVIMENTO_CONTRIBUINTE", "TMC_CODIGO_MOVIMENTACAO = '" & Processo & "'") Then
                Util.Avisa "Operação concluída com sucesso."
                Preencher
            End If
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    txtDataEntrada = ""
    cboAcao.ListIndex = -1
    cboAcao.SetFocus
    Processo = ""
    LblAlvara = ""
    LblSanitario = ""
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores   As String
    Dim Campos   As String
    Dim Condicao As String
    Dim Conta As New ContaCorrente
    

    If Processo = "" Then
        Processo = Conta.GeraCodPagamento(83)
    End If


    Campos = "TMC_CODIGO_MOVIMENTACAO,TMC_TIPO,TMC_TCI_IM,TMC_DATA"
    Valores = Bdados.PreparaValor(Processo, cboAcao.Coluna(1).Valor, Bdados.Converte(txtIm, tctexto), Bdados.Converte(txtDataEntrada, TCDataHora))
    If Bdados.GravaDados("TAB_MOVIMENTO_CONTRIBUINTE", Valores, Campos, "TMC_CODIGO_MOVIMENTACAO = '" & Processo & "'") Then
        'ALTERO OS CAMPOS TMC_LIBERACAO_ALVARA,TMC_LIBERACAO_SANITARIA
        Campos = " TMC_LIBERACAO_ALVARA,TMC_LIBERACAO_SANITARIA,TMC_STATUS_ALVARA"
        If cboAcao.Coluna(1).Valor = 1 Or cboAcao.Coluna(1).Valor = 3 Then
            Valores = Bdados.PreparaValor(Bdados.Converte("0", tctexto), Bdados.Converte("0", tctexto), etsCreditoOriginalAberto)
        Else
            Valores = Bdados.PreparaValor(Bdados.Converte("1", tctexto), Bdados.Converte("1", tctexto))
        End If
        Bdados.GravaDados "TAB_MOVIMENTO_CONTRIBUINTE", Valores, Campos, "TMC_CODIGO_MOVIMENTACAO = '" & Processo & "'"
        Util.Avisa "Processo Aberto com sucesso" & vbCrLf & "Nº Processo " & Processo
        cmdLimpar_Click
        Preencher
        
    End If
End Sub
Private Sub Preencher()
    Dim Sql As String
            
    If txtIm = "" Then Exit Sub
    
    Sql = "SELECT TMC_CODIGO_MOVIMENTACAO as Código,"
    Sql = Sql & " TMC_TCI_IM as IM,"
    Sql = Sql & " TMC_DATA as Data,"
    Sql = Sql & " TGE_NOME as Ação,"
    Sql = Sql & " TMC_TIPO,TMC_LIBERACAO_ALVARA,TMC_LIBERACAO_SANITARIA"
    Sql = Sql & " From TAB_MOVIMENTO_CONTRIBUINTE, VIS_ASSUNTO"
    Sql = Sql & " Where TGE_CODIGO = TMC_TIPO"
    Sql = Sql & " AND TMC_TCI_IM = '" & txtIm & "'"
        grdVISUAL1.Preencher Bdados, Sql, 1000, 1000, 2000, 2000, 0, 0, 0
End Sub

Private Sub Form_Load()
    cboAcao.PreencherGeral Bdados, "ACAO PROTOCOLO"
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    LimpaCampos Me
    LblNomeTFL = "SITUAÇÃO - " & Imposto.NomeTributo(ttr_ALVARA)
    If UCase(AplicacoesVTFuncoes.municipio) <> "BARRA MANSA" Then
        LblOutras.Visible = False
    Else
        LblOutras.Visible = True
    End If
End Sub

Private Sub grdVISUAL1_DblClick()
    If grdVISUAL1.ListItems.Count >= 1 Then
        txtIm = grdVISUAL1.SelectedItem.SubItems(1)
        txtDataEntrada = grdVISUAL1.SelectedItem.SubItems(2)
        Processo = grdVISUAL1.SelectedItem
        'txtIM_LostFocus
        cboAcao.SetarLinha grdVISUAL1.SelectedItem.SubItems(4), 1
        If grdVISUAL1.SelectedItem.SubItems(5) = 1 Then
            LblAlvara = "GERADO"
        Else
            LblAlvara = "NÃO GERADO"
        End If
        If grdVISUAL1.SelectedItem.SubItems(6) = 1 Then
            LblSanitario = "GERADO"
        Else
            LblSanitario = "NÃO GERADO"
        End If
        If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
            LblSanitario.Visible = True
        Else
            LblSanitario.Visible = False
        End If
    End If
End Sub

Private Sub txtIM_LostFocus()
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
    txtIm = BuscaContribuinte(Ic, txtNomeContrib, txtEndereco)
    Preencher
End Sub
