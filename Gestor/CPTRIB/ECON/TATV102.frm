VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TATV102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TATV102.frx":0000
   ScaleHeight     =   8340
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   13
      Top             =   7755
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   1032
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6615
         TabIndex        =   11
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   5460
         TabIndex        =   9
         Top             =   120
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
         Left            =   7770
         TabIndex        =   12
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   7005
      Left            =   75
      TabIndex        =   10
      Top             =   720
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   12356
      Caption         =   "Atividades Estimativas"
      Descricao       =   "Cadastra, realiza manutenção de atividades"
      corFaixa        =   16711680
      Icone           =   "TATV102.frx":0342
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.txtVISUAL txtAno 
         Height          =   285
         Left            =   6540
         TabIndex        =   1
         Top             =   750
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   503
         Caption         =   "Ano"
         Text            =   ""
         Restricao       =   2
         AlinhamentoTexto=   1
         MaxLen          =   10
      End
      Begin VTOcx.txtVISUAL txtUFM 
         Height          =   285
         Left            =   6330
         TabIndex        =   6
         Tag             =   "Valor"
         Top             =   1485
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         Caption         =   "UFM"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtValorExcedente 
         Height          =   285
         Left            =   3750
         TabIndex        =   5
         Top             =   1830
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         Caption         =   "Valor Excedente"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
      End
      Begin VTOcx.cboVISUAL cboFator 
         Height          =   315
         Left            =   105
         TabIndex        =   8
         Top             =   1800
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         Caption         =   "Multiplicar pela base"
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.grdVISUAL grdEstimativo 
         Height          =   2325
         Left            =   60
         TabIndex        =   15
         Top             =   4560
         Width           =   8640
         _ExtentX        =   15240
         _ExtentY        =   4101
         CorBorda        =   16711680
         Caption         =   "Estimativa"
         CorTitulo       =   16711680
         CorCaption      =   16777215
         CorDica         =   16711680
      End
      Begin VTOcx.grdVISUAL grdAtividade 
         Height          =   2325
         Left            =   90
         TabIndex        =   14
         Top             =   2175
         Width           =   8640
         _ExtentX        =   15240
         _ExtentY        =   4101
         CorBorda        =   16711680
         Caption         =   "Atividades"
         CorTitulo       =   16711680
         CorCaption      =   16777215
         CorDica         =   16711680
      End
      Begin VTOcx.txtVISUAL txtLimiteInferior 
         Height          =   285
         Left            =   645
         TabIndex        =   3
         Tag             =   "Limite Inferior"
         Top             =   1470
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   503
         Caption         =   "Limite Inferior"
         Text            =   ""
         Restricao       =   2
      End
      Begin VTOcx.txtVISUAL txtDescAtividade 
         Height          =   285
         Left            =   1020
         TabIndex        =   2
         Top             =   1140
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   503
         Caption         =   "Descrição"
         Text            =   ""
         Enabled         =   0   'False
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   1260
         TabIndex        =   0
         Tag             =   "Código"
         Top             =   795
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   503
         Caption         =   "Código"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   285
         Left            =   6450
         TabIndex        =   7
         Tag             =   "Valor"
         Top             =   1845
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   503
         Caption         =   "R$"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtLimiteSuperior 
         Height          =   285
         Left            =   3840
         TabIndex        =   4
         Tag             =   "Limite Superior"
         Top             =   1470
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         Caption         =   "Limite Superior"
         Text            =   ""
         Restricao       =   2
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   1138
      Icone           =   "TATV102.frx":065C
   End
End
Attribute VB_Name = "TATV102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AtividadeEstimada As eAtividadeEstimada
Dim atividade As atividade

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtCodigo.SetFocus
    grdEstimativo.ListItems.Clear
    AtividadeEstimada.PreencherAtividadesEstimativas grdAtividade
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    If txtValorExcedente = "" Then txtValorExcedente = "0"
    With AtividadeEstimada
        .CodAtividade = txtCodigo
        .LimiteInferior = txtLimiteInferior
        .LimiteSuperior = txtLimiteSuperior
        .ValorUFM = txtUFM
        .LimiteValor = txtValor
        .LimiteFator = CInt(cboFator.Coluna(1).Valor) - 1
        .ValorExcedente = txtValorExcedente
        .Ano = CInt(Nvl(txtAno, 0))
        If .Gravar Then
            Util.Avisa "Transação completada com sucesso."
            .PreencherGrd grdEstimativo, txtCodigo
            LimparDados
            txtLimiteInferior.SetFocus
        End If
        Screen.MousePointer = 0
    End With
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set AtividadeEstimada = New eAtividadeEstimada
    Set atividade = New atividade
    
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    If Temp.PegaParametro(Bdados, "ALVARA ESTIMATIVA GERAL") = "SIM" Then
        txtCodigo.Text = "0"
        txtCodigo.Enabled = False
        txtCodigo_LostFocus
        txtDescAtividade = "(TODAS AS ATIVIDADES ECONÔMICAS)"
        grdAtividade.Visible = False
        grdEstimativo.Top = grdAtividade.Top
        fraFUTURO1.Height = 4575
        Me.Height = 6255
    Else
        AtividadeEstimada.PreencherAtividadesEstimativas grdAtividade
        AtualizaCabecalho grdAtividade
    End If
    AtualizaCabecalho grdEstimativo
    cboFator.PreencherGeral Bdados, "SIM OU NÃO"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set AtividadeEstimada = Nothing
    Set atividade = Nothing
End Sub

Private Sub grdAtividade_DblClick()
    If grdAtividade.ListItems.Count >= 1 Then
        txtCodigo = grdAtividade.SelectedItem
        txtCodigo_LostFocus
    End If
End Sub

Private Sub grdEstimativo_Click()
    If grdEstimativo.SelectedItem Is Nothing Then Exit Sub
    txtLimiteInferior = grdEstimativo.SelectedItem.SubItems(1)
    txtLimiteSuperior = grdEstimativo.SelectedItem.SubItems(2)
    txtUFM = grdEstimativo.SelectedItem.SubItems(3)
    txtValor = grdEstimativo.SelectedItem.SubItems(4)
    cboFator.SetarLinha (grdEstimativo.SelectedItem.SubItems(5)), 0
    txtValorExcedente = "" & grdEstimativo.SelectedItem.SubItems(6)
End Sub

Private Sub grdEstimativo_DblClick()
    If grdEstimativo.SelectedItem Is Nothing Then Exit Sub
    If Confirma("Deseja excluir a " & grdEstimativo.SelectedItem.Index & "ª faixa?") Then
        If AtividadeEstimada.Excluir(grdEstimativo.SelectedItem, grdEstimativo.SelectedItem.SubItems(1)) Then
            Util.Avisa "Faixa Eliminada com Sucesso."
            AtividadeEstimada.PreencherGrd grdEstimativo, txtCodigo
            LimparDados
        End If
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    If Trim(txtCodigo) = "" Then Exit Sub
    If grdAtividade.Visible Then 'ESTIMATIVA POR ATIVIDADE
        If atividade.Buscar(txtCodigo, True) Then
            txtDescAtividade = atividade.Nome
            AtividadeEstimada.PreencherGrd grdEstimativo, txtCodigo, CInt(Nvl(txtAno, 0))
        End If
    Else 'ESTIMATIVA GERAL
        txtDescAtividade = "(TODAS AS ATIVIDADES ECONÔMICAS)"
        AtividadeEstimada.PreencherGrd grdEstimativo, txtCodigo
    End If
End Sub

Private Sub LimparDados()
    txtLimiteInferior = ""
    txtLimiteSuperior = ""
    txtValor = ""
    txtUFM = ""
    cboFator.ListIndex = -1
    txtValorExcedente = ""
End Sub

Private Sub txtUFM_LostFocus()
    If txtUFM = "" Then Exit Sub
    txtValor = Calcula_UFM(txtUFM, Converete_Real)
End Sub

Private Sub txtValor_LostFocus()
    If txtValor = "" Then Exit Sub
    txtUFM = Calcula_UFM(txtValor, Converete_UFM)
End Sub
