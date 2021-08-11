VERSION 5.00
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TAVI102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAVI102"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grdCPND 
      Height          =   2070
      Left            =   30
      TabIndex        =   16
      Top             =   2640
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   3651
      CorBorda        =   32768
      Caption         =   "Avisos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   855
      Left            =   3210
      TabIndex        =   17
      Top             =   720
      Width           =   3915
      _ExtentX        =   6906
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         Picture         =   "TAVI102.frx":0000
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
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1138
      Icone           =   "TAVI102.frx":2123
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   6540
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   873
      Begin VTOcx.txtVISUAL txtDataEntrega 
         Height          =   300
         Left            =   3960
         TabIndex        =   21
         Tag             =   "Validade"
         Top             =   120
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   529
         Caption         =   "Data da Entrega"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         MaxLen          =   10
      End
      Begin VTOcx.cmdVISUAL cmdBuscarContrib 
         Height          =   375
         Left            =   6990
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
         Left            =   9270
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
         Left            =   10215
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
         Left            =   7935
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
   Begin VTOcx.txtVISUAL txtIm 
      Height          =   300
      Left            =   570
      TabIndex        =   1
      Top             =   1170
      Width           =   2235
      _ExtentX        =   3942
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
      Left            =   2835
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1185
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      Caption         =   ""
      Acao            =   5
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   855
      Left            =   7170
      TabIndex        =   18
      Top             =   720
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   1508
      Altura          =   1905
      Caption         =   " Período de Vencimento"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtAte 
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
      Begin VTOcx.txtVISUAL txtValidade 
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
   Begin VTOcx.fraVISUAL fraVISUAL3 
      Height          =   855
      Left            =   7200
      TabIndex        =   20
      Top             =   1620
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   1508
      Altura          =   1905
      Caption         =   " Período de Emissão"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtEmissao 
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
      Begin VTOcx.txtVISUAL txtEmissaoFim 
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
   End
   Begin VTOcx.txtVISUAL txtNotificacao 
      Height          =   300
      Left            =   630
      TabIndex        =   0
      Top             =   795
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   529
      Caption         =   "Nº Aviso"
      Text            =   ""
      Restricao       =   2
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.grdVISUAL GrdItems 
      Height          =   1800
      Left            =   60
      TabIndex        =   19
      Top             =   4740
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   3175
      CorBorda        =   32768
      Caption         =   "Débitos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
End
Attribute VB_Name = "TAVI102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Certidao As iCertidao


Private Sub cmdBuscarContrib_Click()

    If Trim(txtIm) = "" And Trim(txtNotificacao) = "" Then
        Util.Avisa "Informe o Nº do Aviso ou a Inscrição."
        txtNotificacao.SetFocus
        Exit Sub
    End If
    Dim Sql As String
    Sql = Sql & " SELECT TNT_COD_NOTIFICACAO AS Notificação,"
    Sql = Sql & " TNT_INSCRICAO as Inscrição,  TNT_DT_EMISSAO as Emissão, "
    Sql = Sql & " TNT_VENCIMENTO as Vencimento,  TNT_VALOR_NOTIFICACAO as Valor,"
    Sql = Sql & " TNT_TUS_COD_USUARIO As Usuário"
    Sql = Sql & " , TNT_ENTREGA as Entrega"
    Sql = Sql & " From TAB_NOTIFICACAO where 1 = 1 and TNT_TIPO = 2 "
    
    
    If txtEntrega <> "" And txtFimEntrega <> "" Then
        If txtEntrega >= txtFimEntrega Then
            Util.Avisa "Data inválida."
            txtEntrega.SetFocus
            Exit Sub
        End If
    End If
    If txtEmissao = "" And txtEmissaoFim <> "" Then
        Util.Avisa "Informe a data de inicio."
        txtEmissao.SetFocus
        Exit Sub
    End If
    If txtFimEntrega <> "" And txtEntrega = "" Then
        Util.Avisa "Informe a data de inicio."
        txtEntrega.SetFocus
        Exit Sub
    End If
    If txtValidade <> "" And txtAte <> "" Then
        If txtValidade <= txtAte Then
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
    
    If txtNotificacao <> "" Then
        Sql = Sql & " and TNT_COD_NOTIFICACAO = '" & txtNotificacao & "'"
    End If
    
    If txtIm <> "" Then
        Sql = Sql & "AND TNT_INSCRICAO = '" & txtIm & "'"
    End If
    
    
    If txtValidade <> "" And txtAte <> "" Then
        Sql = Sql & " and TNT_VENCIMENTO >=  " & Bdados.Converte(txtValidade, TCDataHora) & " and TNT_VENCIMENTO <= " & Bdados.Converte(txtAte, TCDataHora)
    ElseIf txtValidade <> "" And txtAte = "" Then
        Sql = Sql & " and TNT_VENCIMENTO >= " & Bdados.Converte(txtValidade, TCDataHora) & " and TNT_VENCIMENTO <=  " & Bdados.Converte(txtValidade, TCDataHora)
    End If
        
    If txtEntrega <> "" And txtFimEntrega <> "" Then
        Sql = Sql & " and TNT_ENTREGA >= " & Bdados.Converte(txtEntrega, TCDataHora) & " and TNT_ENTREGA <=  " & Bdados.Converte(txtFimEntrega, TCDataHora)
    ElseIf txtEntrega <> "" And txtFimEntrega = "" Then
        Sql = Sql & " and TNT_ENTREGA >= " & Bdados.Converte(txtEntrega, TCDataHora) & " and TNT_ENTREGA <= " & Bdados.Converte(txtEntrega, TCDataHora)
    End If
        
    If txtEmissao <> "" And txtEmissaoFim <> "" Then
        Sql = Sql & " and TNT_DT_EMISSAO >= " & Bdados.Converte(txtEmissao, TCDataHora) & " and TNT_DT_EMISSAO <=  " & Bdados.Converte(txtEmissaoFim, TCDataHora)
    ElseIf txtEmissao <> "" And txtEmissaoFim = "" Then
        Sql = Sql & " and TNT_DT_EMISSAO >= " & Bdados.Converte(txtEmissao, TCDataHora) & " and TNT_DT_EMISSAO <= " & Bdados.Converte(txtEmissao, TCDataHora)
    End If
    
    
    grdCPND.Preencher Bdados, Sql
    If grdCPND.ListItems.Count <= 0 Then
        Util.Avisa "Consulta sem resultados."
    End If
    GrdItems.ListItems.Clear
End Sub

Private Sub cmdEmitir_Click()
    Dim Data As String
    If Trim(txtIm) = "" And Trim(txtNotificacao) = "" Then Exit Sub
    If grdCPND.ListItems.Count >= 1 Then
        If Confirma("Confirma a entrega do aviso?") Then
            Rem Data = imposto.BuscaDataVencimento(
            If txtDataEntrega = "" Then
                Util.Avisa "Informe data de entrega."
                txtDataEntrega.SetFocus
                Exit Sub
            End If
            If Bdados.GravaDados("tab_notificacao", Bdados.PreparaValor(txtDataEntrega), "tnt_entrega", "tnt_cod_notificacao = '" & grdCPND.SelectedItem & "'") Then
                Util.Avisa "Entrega realizada com sucesso."
                cmdBuscarContrib_Click
            End If
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdCPND.ListItems.Clear
    GrdItems.ListItems.Clear
    txtNotificacao.SetFocus
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
End Sub

Private Sub grdCPND_DblClick()
    Dim Sql As String
    If grdCPND.ListItems.Count >= 1 Then
        Sql = "SELECT TPE_TGT_COD_PAGAMENTO as Pagamento,tip_nome_imposto as Imposto,TPE_SUB_VALOR as Valor"
        Sql = Sql & " From TAB_PAGAMENTO_extrato, tab_imposto"
        Sql = Sql & " Where TPE_TIP_COD_IMPOSTO = tip_cod_imposto"
        Sql = Sql & " and TPE_COD_PAGAMENTO_EXTRATO = '" & grdCPND.SelectedItem & "'"
        GrdItems.Preencher Bdados, Sql
    End If
End Sub
