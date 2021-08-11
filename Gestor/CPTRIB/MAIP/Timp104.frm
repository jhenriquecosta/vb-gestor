VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TIMP104 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   14
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "Timp104.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.cboVISUAL cboDestinacao 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Tag             =   "11"
      Top             =   4350
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   556
      Caption         =   "Destinação"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.cboVISUAL cboOcupacao 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Tag             =   "1"
      Top             =   3990
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      Caption         =   "Ocupação"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   1138
      Icone           =   "Timp104.frx":2123
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   10
      Top             =   5460
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   979
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   4800
         TabIndex        =   13
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   5775
         TabIndex        =   8
         Top             =   120
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   2955
         TabIndex        =   11
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         Caption         =   "&Novo"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   3825
         TabIndex        =   7
         Top             =   120
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
   End
   Begin VTOcx.txtVISUAL txtCodigo 
      Height          =   300
      Left            =   420
      TabIndex        =   0
      Tag             =   "Codigo"
      Top             =   3630
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      Caption         =   "Codigo"
      Text            =   ""
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtLimInferior 
      Height          =   300
      Left            =   1050
      TabIndex        =   3
      Tag             =   "Lim Inferior"
      Top             =   4740
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   529
      Caption         =   "Limite Inferior"
      Text            =   ""
      Formato         =   5
      Restricao       =   3
   End
   Begin VTOcx.txtVISUAL txtAliqPropria 
      Height          =   300
      Left            =   360
      TabIndex        =   5
      Tag             =   "Aliq Proprio"
      Top             =   5100
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   529
      Caption         =   "Aliquota Rec. Próprios"
      Text            =   ""
      Formato         =   5
      Restricao       =   3
   End
   Begin VTOcx.grdVISUAL grdTransferencia 
      Height          =   3135
      Left            =   60
      TabIndex        =   12
      Top             =   690
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   4339
      CorBorda        =   32768
      Caption         =   "Tipo Transferência"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin VTOcx.txtVISUAL txtLimSuperior 
      Height          =   300
      Left            =   3780
      TabIndex        =   4
      Tag             =   "Lim Superior"
      Top             =   4740
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      Caption         =   "Superior"
      Text            =   ""
      Formato         =   5
      Restricao       =   3
   End
   Begin VTOcx.txtVISUAL txtAliqFinanciado 
      Height          =   300
      Left            =   3630
      TabIndex        =   6
      Tag             =   "Aliq Financiado"
      Top             =   5100
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   529
      Caption         =   "Financiado"
      Text            =   ""
      Formato         =   5
      Restricao       =   3
   End
End
Attribute VB_Name = "TIMP104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conta As New ContaCorrente

Private Sub cmdExcluir_Click()
    If Not grdTransferencia.SelectedItem Is Nothing Then
        If Confirma("Excluir " & grdTransferencia.SelectedItem & " ?") Then
            If Bdados.DeletaDados("TAB_TIPO_TRANSFERENCIA_IMOVEL", "TTT_COD_ALIQUOTA='" & grdTransferencia.SelectedItem & "'") Then
                Edita.LimpaCampos Me
                ExibirTipos
            End If
        End If
    End If
End Sub

Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    txtCodigo.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Campos As String, Valores As String
    If Edita.CriticaCampos(Me) Then
        Campos = "TTT_COD_ALIQUOTA, TTT_TCO_COD_OCUPACAO, TTT_TCO_COD_DESTINACAO, TTT_LIMITE_INFERIOR, TTT_LIMITE_SUPERIOR, TTT_ALIQUOTA_PROPRIO, TTT_ALIQUOTA_FINANCIADO"
        Valores = Bdados.PreparaValor(txtCodigo, cboOcupacao.Coluna(1).Valor, cboDestinacao.Coluna(1).Valor, txtLimInferior, txtLimSuperior, txtAliqPropria, txtAliqFinanciado)
        If Bdados.GravaDados("Tab_Tipo_Transferencia_Imovel", Valores, Campos, "TTT_COD_ALIQUOTA='" & txtCodigo & "'") Then
            'Avisa "Registro gravado com sucesso."
            cmdNovo_Click
            ExibirTipos
        End If
    End If
End Sub

Private Sub ExibirTipos()
    Dim Sql As String
    
    Sql = "SELECT TTT_COD_ALIQUOTA as Codigo, " & _
                " VIS_OCUPACAO.TGE_NOME AS Ocupacao," & _
                " VIS_DESTINACAO.TGE_NOME AS Destinacao," & _
                " TTT_LIMITE_INFERIOR as LimInferior, " & _
                " TTT_LIMITE_SUPERIOR as LimSuperior, " & _
                " TTT_ALIQUOTA_PROPRIO as AliqProprio, " & _
                " TTT_ALIQUOTA_FINANCIADO as AliqFinanciado" & _
            " FROM Tab_Tipo_Transferencia_Imovel, VIS_OCUPACAO,VIS_DESTINACAO " & _
            " WHERE VIS_OCUPACAO.TGE_CODIGO = TTT_TCO_COD_OCUPACAO AND " & _
                " VIS_DESTINACAO.TGE_CODIGO = TTT_TCO_COD_DESTINACAO"
    grdTransferencia.Preencher Bdados, Sql ', (grdTransferencia.Width * 20 / 100), (grdTransferencia.Width * 40 / 100), (grdTransferencia.Width * 20 / 100), (grdTransferencia.Width * 20 / 100)
End Sub
Private Sub Form_Load()
    Dim Controle As Control
    cabVisual.Exibir Bdados, Me.Name, App.Path
    cboDestinacao.PreencherGeral Bdados, "ITBI  DESTINO"
    cboOcupacao.PreencherGeral Bdados, "ITBI OCUPACAO"
    ExibirTipos
End Sub

Private Sub grdTransferencia_DblClick()
    If Not grdTransferencia.SelectedItem Is Nothing Then
        With grdTransferencia.SelectedItem
            txtCodigo = .Text
            cboOcupacao = .SubItems(1)
            cboDestinacao = .SubItems(2)
            txtLimInferior = .SubItems(3)
            txtLimSuperior = .SubItems(4)
            txtAliqPropria = .SubItems(5)
            txtAliqFinanciado = .SubItems(6)
        End With
    End If
End Sub
