VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU108 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU108"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7695
   StartUpPosition =   1  'CenterOwner
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   3900
      Left            =   90
      TabIndex        =   7
      Top             =   2205
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   6879
      CorBorda        =   16711680
      CorTitulo       =   16711680
      CorCaption      =   -2147483634
      CorDica         =   16711680
   End
   Begin VB.TextBox txtDescricao 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   2250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1140
      Width           =   4905
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   6120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   4215
         TabIndex        =   4
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   3045
         TabIndex        =   3
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6540
         TabIndex        =   6
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   5385
         TabIndex        =   5
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.txtVISUAL txtCodigo 
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   825
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   503
      Caption         =   "Código do Componente"
      Text            =   ""
   End
   Begin VTOcx.cboVISUAL cboCategoria 
      Height          =   315
      Left            =   1365
      TabIndex        =   2
      Top             =   1815
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   556
      Caption         =   "Categoria"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1138
      Icone           =   "TCIU108.frx":0000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   1500
      TabIndex        =   9
      Top             =   1140
      Width           =   720
   End
End
Attribute VB_Name = "TCIU108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcluir_Click()
    If txtCodigo.Enabled = False Then
        If Confirma("Deseja excluir esse grupo?") = True Then
            If Bdados.DeletaDados("tab_grupo_componente_avancado", "tgc_cod_grupo = '" & txtCodigo & "'") Then
                Avisa "Operação concluída com sucesso"
                cmdLimpar_Click
                PreencherGrid
            End If
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me

    txtCodigo.Enabled = True
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub PreencherGrid()
   grdDados.Preencher Bdados, "select tgc_cod_grupo as Código,tgc_nome as Descrição,tgc_categoria as Categoria from tab_grupo_componente_avancado"
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    
    Campos = "tgc_cod_grupo,tgc_nome,tgc_categoria"
    Valores = Bdados.PreparaValor(txtCodigo, Bdados.Converte(txtDescricao, tctexto), cboCategoria.Coluna(1).Valor)
    Condicao = "tgc_cod_grupo = " & txtCodigo
    If Bdados.GravaDados("tab_grupo_componente_avancado", Valores, Campos, Condicao) Then
        Util.Avisa "Operação concluída com sucesso."
        cmdLimpar_Click
        PreencherGrid
    End If
End Sub

Private Sub Form_Load()
    cboCategoria.PreencherGeral Bdados, "CATEGORIA GRUPO IMOVEL"
    PreencherGrid
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
End Sub

Private Sub grdDados_DblClick()
    If grdDados.ListItems.Count >= 1 Then
        txtCodigo = grdDados.SelectedItem
        txtCodigo_LostFocus
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "Select * from tab_grupo_componente_avancado where tgc_cod_grupo = '" & txtCodigo & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        txtDescricao = "" & Rs.Fields("tgc_nome")
        cboCategoria.SetarLinha Rs.Fields("tgc_categoria"), 1
        txtCodigo.Enabled = False
    Else
        txtCodigo.Enabled = True
        txtDescricao = ""
        cboCategoria.ListIndex = -1
    End If
End Sub

