VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCER104 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCER104"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   855
      Left            =   2520
      TabIndex        =   18
      Top             =   1980
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   1508
      Altura          =   1905
      Caption         =   " Período de Entrega"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtFimEntrega 
         Height          =   300
         Left            =   2160
         TabIndex        =   5
         Tag             =   "Validade"
         Top             =   390
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   529
         Caption         =   "Até"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtEntrega 
         Height          =   300
         Left            =   90
         TabIndex        =   4
         Tag             =   "Validade"
         Top             =   405
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "De"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         MaxLen          =   10
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   60
      ScaleHeight     =   570
      ScaleWidth      =   555
      TabIndex        =   14
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCER104.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1138
      Icone           =   "TCER104.frx":2123
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   6540
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL cmdBuscarContrib 
         Height          =   375
         Left            =   7110
         TabIndex        =   8
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
         Left            =   9390
         TabIndex        =   10
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
         Left            =   10335
         TabIndex        =   11
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
         Left            =   8055
         TabIndex        =   9
         Top             =   75
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
         Caption         =   "Confirmar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.cboVISUAL cboTributo 
      Height          =   315
      Left            =   1845
      TabIndex        =   1
      Top             =   1185
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   556
      Caption         =   "Tributo"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.txtVISUAL txtIm 
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   529
      Caption         =   "Inscricão"
      Text            =   ""
      Restricao       =   2
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
      Height          =   285
      Left            =   4815
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1575
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      Caption         =   ""
      Acao            =   5
   End
   Begin VTOcx.txtVISUAL txtImovel 
      Height          =   300
      Left            =   5910
      TabIndex        =   3
      Top             =   1575
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   529
      Caption         =   "Cadastro do Imóvel"
      Text            =   ""
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.cmdVISUAL cmdVISUAL1 
      Height          =   285
      Left            =   10260
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1575
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      Caption         =   ""
      Acao            =   5
   End
   Begin VTOcx.grdVISUAL grdCPND 
      Height          =   3660
      Left            =   60
      TabIndex        =   17
      Top             =   2880
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6456
      CorBorda        =   32768
      Caption         =   "Certidões emitidas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin VTOcx.cboVISUAL CboDocumento 
      Height          =   315
      Left            =   1710
      TabIndex        =   0
      Top             =   780
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   556
      Caption         =   "Certidão"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   855
      Left            =   6660
      TabIndex        =   19
      Top             =   1980
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1508
      Altura          =   1905
      Caption         =   " Período de Validade"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtAte 
         Height          =   300
         Left            =   2160
         TabIndex        =   7
         Tag             =   "Validade"
         Top             =   390
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   529
         Caption         =   "Até"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtValidade 
         Height          =   300
         Left            =   90
         TabIndex        =   6
         Tag             =   "Validade"
         Top             =   405
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   529
         Caption         =   "De"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         MaxLen          =   10
      End
   End
End
Attribute VB_Name = "TCER104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Certidao As iCertidao

Private Sub CboDocumento_Click()
    Dim sql As String
    sql = "SELECT TCN_COD_NEGATIVA AS Certidão,"
    sql = sql & " TCN_TCI_IM as [Insc.Municipal],"
    sql = sql & " TCN_TIM_IC as [Insc.Cadastral],"
    sql = sql & " TCN_DATA_NEGATIVA as Geração,"
    sql = sql & " TCN_FINALIDADE as Finalidade,"
    sql = sql & " TCN_VALIDADE as Validade,"
    sql = sql & " TCN_TUS_COD_USUARIO as Funcionário,"
    sql = sql & " TCN_TIP_COD_IMPOSTO as Imposto,"
    sql = sql & " TCN_DATA_ENTREGA As Entrega"
    sql = sql & " From TAB_CERTIDAO_NEGATIVA"
    sql = sql & " where tcn_tipo = '" & CboDocumento.Coluna(1).Valor & "'"
    grdCPND.Preencher Bdados, sql
End Sub

Private Sub cmdBuscarContrib_Click()
    Dim sql As String
    sql = "SELECT TCN_COD_NEGATIVA AS Certidão,"
    sql = sql & " TCN_TCI_IM as [Insc.Municipal],"
    sql = sql & " TCN_TIM_IC as [Insc.Cadastral],"
    sql = sql & " TCN_DATA_NEGATIVA as Geração,"
    sql = sql & " TCN_FINALIDADE as Finalidade,"
    sql = sql & " TCN_VALIDADE as Validade,"
    sql = sql & " TCN_TUS_COD_USUARIO as Funcionário,"
    sql = sql & " TCN_TIP_COD_IMPOSTO as Imposto,"
    sql = sql & " TCN_DATA_ENTREGA As Entrega"
    sql = sql & " From TAB_CERTIDAO_NEGATIVA"
    sql = sql & " where tcn_tipo = '" & CboDocumento.Coluna(1).Valor & "'"
    
    If txtEntrega <> "" And txtFimEntrega <> "" Then
        If txtEntrega >= txtFimEntrega Then
            Util.Avisa "Data inválida."
            txtEntrega.SetFocus
            Exit Sub
        End If
    End If
    If txtFimEntrega <> "" And txtEntrega = "" Then
        Util.Avisa "Informe a data de inicio."
        txtEntrega.SetFocus
        Exit Sub
    End If
    If txtValidade <> "" And txtAte <> "" Then
        If txtValidade >= txtAte Then
            Util.Avisa "Data inválida."
            txtValidade.SetFocus
            Exit Sub
        End If
    End If
    If txtAte <> "" And txtValidade = "" Then
        Util.Avisa "Informe a data de inicio."
        txtValidade.SetFocus
        Exit Sub
    End If
    
    If txtIm <> "" Then
        sql = sql & "AND TCN_TCI_IM = '" & txtIm & "'"
    End If
    
    If txtImovel <> "" Then
        sql = sql & " AND TCN_TIM_IC = '" & txtImovel & "'"
    End If
    
    If cboTributo.ListIndex >= 0 Then
        sql = sql & "AND TCN_TIP_COD_IMPOSTO = '" & cboTributo & "'"
    End If
    
    If txtValidade <> "" And txtAte <> "" Then
        sql = sql & " and TCN_VALIDADE >=  " & Bdados.Converte(txtValidade, TCDataHora) & " and TCN_VALIDADE <= " & Bdados.Converte(txtAte, TCDataHora)
    ElseIf txtValidade <> "" And txtAte = "" Then
        sql = sql & " and TCN_VALIDADE >= " & Bdados.Converte(txtValidade, TCDataHora) & " and TCN_VALIDADE <=  " & Bdados.Converte(txtValidade, TCDataHora)
    End If
        
    If txtEntrega <> "" And txtFimEntrega <> "" Then
        sql = sql & " and TCN_DATA_ENTREGA >= " & Bdados.Converte(txtEntrega, TCDataHora) & " and TCN_DATA_ENTREGA <=  " & Bdados.Converte(txtFimEntrega, TCDataHora)
    ElseIf txtEntrega <> "" And txtFimEntrega = "" Then
        sql = sql & " and TCN_DATA_ENTREGA >= " & Bdados.Converte(txtEntrega, TCDataHora) & " and TCN_DATA_ENTREGA <= " & Bdados.Converte(txtEntrega, TCDataHora)
    End If
    
    grdCPND.Preencher Bdados, sql
    If grdCPND.ListItems.Count <= 0 Then
        Util.Avisa "Consulta sem resultados."
    End If
End Sub

Private Sub cmdEmitir_Click()
    If Confirma("Deseja confirmar a entrega da " & CboDocumento.Text & " Número - " & grdCPND.SelectedItem & "?") Then
        If Bdados.GravaDados("tab_certidao_negativa", Bdados.PreparaValor(Date), "TCN_DATA_ENTREGA", "TCN_COD_NEGATIVA = '" & grdCPND.SelectedItem & "'") Then
            Util.Avisa "Entrega confirmada com sucesso."
            cmdBuscarContrib_Click
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdCPND.ListItems.Clear
    CboDocumento.SetFocus
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
   AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    txtValidade = DateAdd("d", Nvl(Temp.PegaParametro(Bdados, "VALIDADE CERTIDAO"), 0), Date)
    Set Certidao = New iCertidao
    Certidao.PreencherCboImposto cboTributo
    CboDocumento.PreencherGeral Bdados, "TIPO CERTIDAO"
End Sub
