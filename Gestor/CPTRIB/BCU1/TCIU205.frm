VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU205 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU205"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.cboVISUAL CboBairro 
      Height          =   315
      Left            =   1185
      TabIndex        =   2
      Top             =   1485
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   556
      Caption         =   "Bairro"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.txtVISUAL txtLoteamento 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   765
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   503
      Caption         =   " Nº do Loteamento"
      Text            =   ""
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   12
      Top             =   6225
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   900
      Begin VTOcx.cmdVISUAL CmdBuscar 
         Height          =   375
         Left            =   4500
         TabIndex        =   5
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   0
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   8010
         TabIndex        =   8
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   0
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9165
         TabIndex        =   9
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   0
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   5670
         TabIndex        =   6
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   0
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   6840
         TabIndex        =   7
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   0
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1470
      TabIndex        =   11
      Top             =   -570
      Width           =   375
   End
   Begin VTOcx.grdVISUAL grid 
      Height          =   3795
      Left            =   30
      TabIndex        =   10
      Top             =   2400
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   6694
      CorBorda        =   16711680
      Caption         =   "Loteamentos"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   16711680
   End
   Begin VTOcx.txtVISUAL txtDescricao 
      Height          =   285
      Left            =   885
      TabIndex        =   1
      Top             =   1125
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   503
      Caption         =   "Descrição"
      Text            =   ""
   End
   Begin VTOcx.cboVISUAL CboLogradouro 
      Height          =   315
      Left            =   3405
      TabIndex        =   4
      Top             =   1875
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   556
      Caption         =   "Logradouro"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.cboVISUAL cboTipLogra 
      Height          =   315
      Left            =   930
      TabIndex        =   3
      Top             =   1860
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   556
      Caption         =   "Endereço"
      Text            =   ""
      AutoFocaliza    =   0   'False
      CorFundo        =   -2147483644
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1138
      Icone           =   "TCIU205.frx":0000
   End
End
Attribute VB_Name = "TCIU205"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cadastro As VSImposto

Private Sub cmdBuscar_Click()
    grid.Preencher Bdados, "Select Loteamento as Codigo,Descrição,Bairro,Logradouro,codBairro,CodLogradouro from vis_loteamento order by 1", 1000, 3000, 3000, 3000, 0, 0
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdExcluir_Click()
    If Confirma("Deseja excluir esse registro?", "Aviso") = True Then
        If Bdados.DeletaDados("TAB_LOTEAMENTO", "TLO_COD_LOTEAMENTO= " & txtLoteamento) Then
            Util.Avisa "Registro excluído com sucesso."
            cmdLimpar_Click
            cmdBuscar_Click
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtLoteamento.Enabled = True
    txtLoteamento.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Campos = "TLO_COD_LOTEAMENTO,TLO_DESCRICAO,TLO_TBA_COD_BAIRRO,TLO_TLG_COD_LOGRADOURO,TLO_TTL_COD_TIP_LOGR"
    Valores = Bdados.PreparaValor(txtLoteamento, txtDescricao, cboBairro.Coluna(0).Valor, CboLogradouro.Coluna(0).Valor, cboTipLogra.Coluna(0).Valor)
    If Bdados.GravaDados("TAB_LOTEAMENTO", Valores, Campos, "TLO_COD_LOTEAMENTO= " & txtLoteamento) Then
        Informa "Transação completada."
        cmdLimpar_Click
        cmdBuscar_Click
    End If
End Sub

Private Sub Form_Activate()
    cboBairro.Preencher Bdados, "select tba_cod_bairro,tba_nome from tab_bairro ", 1
    CboLogradouro.Preencher Bdados, "Select * from VIS_LOGRADOURO_COMBO", 1
    cboTipLogra.Preencher Bdados, "SELECT TTL_COD_TIP_LOGR,TTL_NOME FROM TAB_TIPO_LOGR", 1
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path

End Sub
Private Sub grid_DblClick()
    If grid.ListItems.Count >= 1 Then
        txtLoteamento = grid.SelectedItem
        txtLoteamento_LostFocus
    End If
End Sub

Private Sub txtLoteamento_LostFocus()
    Dim Rs As VSRecordset
    Dim Sql As String
    
    If txtLoteamento = "" Then Exit Sub
    
    Sql = "Select * from vis_Loteamento where Loteamento = '" & txtLoteamento & "'"
   If Bdados.AbreTabela(Sql, Rs) Then
        cboTipLogra.SetarLinha "" & Rs.Fields("TLO_TTL_COD_TIP_LOGR")
        txtDescricao = Rs.Fields("Descrição")
        cboBairro.SetarLinha Nvl("" & Rs.Fields("CodBairro"), 0)
        CboLogradouro.SetarLinha Rs.Fields("codlogradouro")
        txtLoteamento.Enabled = False
   Else
        txtLoteamento.Enabled = True
        txtDescricao = ""
        cboBairro.ListIndex = -1
        CboLogradouro.ListIndex = -1
   End If
    
End Sub
