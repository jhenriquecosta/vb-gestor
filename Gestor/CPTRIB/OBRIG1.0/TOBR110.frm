VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECA~1.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONT~1.OCX"
Begin VB.Form TOBR110 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TOBR110"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2415
      Index           =   3
      Left            =   15
      TabIndex        =   18
      Top             =   645
      Width           =   8880
      Begin VB.CheckBox chkRemessa 
         Caption         =   "Remessa"
         Height          =   255
         Left            =   7680
         TabIndex        =   25
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1000
      End
      Begin VTOcx.txtVISUAL txtParcela 
         Height          =   300
         Left            =   7575
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "Parcela"
         Text            =   ""
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1185
         TabIndex        =   5
         Top             =   1275
         Width           =   7620
      End
      Begin VTOcx.cboVISUAL cboImposto 
         Height          =   315
         Left            =   540
         TabIndex        =   0
         Tag             =   "Tributo"
         Top             =   150
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   556
         Caption         =   "Tributo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   300
         Left            =   360
         TabIndex        =   2
         Tag             =   "''"
         Top             =   525
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   529
         Caption         =   "Inscricão"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   300
         Left            =   45
         TabIndex        =   4
         Top             =   915
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   529
         Caption         =   "Nome/Razão"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtPeriodoInicial 
         Height          =   300
         Left            =   2235
         TabIndex        =   6
         Tag             =   "Período "
         Top             =   1680
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   529
         Caption         =   "Periodo "
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
         Height          =   300
         Left            =   2940
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   525
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtImovel 
         Height          =   300
         Left            =   4290
         TabIndex        =   3
         Top             =   525
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   529
         Caption         =   "Cadastro do Imóvel"
         Text            =   ""
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   300
         Left            =   8415
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   525
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.txtVISUAL txtVence 
         Height          =   300
         Left            =   3825
         TabIndex        =   7
         Tag             =   "Data Vencimento"
         Top             =   1680
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   529
         Caption         =   "Vencimento"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   300
         Left            =   6000
         TabIndex        =   8
         Tag             =   "Valor Obrigação"
         Top             =   1680
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   529
         Caption         =   "Valor"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoTexto=   1
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtDoc 
         Height          =   300
         Left            =   135
         TabIndex        =   1
         Top             =   525
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   529
         Caption         =   "Número Doc."
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtOrigem 
         Height          =   300
         Left            =   60
         TabIndex        =   11
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         Caption         =   "Doc. Origem"
         Text            =   ""
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtObs 
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   529
         Caption         =   "Observação"
         Text            =   ""
         Requerido       =   0   'False
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   6135
         TabIndex        =   22
         Top             =   1290
         Width           =   45
      End
      Begin VB.Label lblGerado 
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   6420
         TabIndex        =   21
         Top             =   2085
         Width           =   4920
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   17
      Top             =   5445
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   900
      Begin VTOcx.cmdVISUAL cmdVISUAL2 
         Height          =   375
         Left            =   5460
         TabIndex        =   24
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6615
         TabIndex        =   13
         Top             =   90
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
         Left            =   7770
         TabIndex        =   14
         Top             =   90
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
         Left            =   4320
         TabIndex        =   12
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1470
      TabIndex        =   15
      Top             =   -570
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   1138
      Icone           =   "TOBR110.frx":0000
   End
   Begin VTOcx.grdVISUAL lstObrig 
      Height          =   2610
      Left            =   15
      TabIndex        =   23
      Top             =   3090
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4604
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      CheckBox        =   -1  'True
   End
End
Attribute VB_Name = "TOBR110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cadastro As VSImposto
Private Obrig As New Obrigacao
Private String_Taxas  As String
Private Total_Taxas As Double

Private Sub Pega_taxas()
    Dim i As Integer
    Dim Pos As Integer
    String_Taxas = ""
    Total_Taxas = 0
    'For i = 1 To GrdTaxas.ListItems.Count
    '    If GrdTaxas.ListItems(i).Checked Then
    '        Pos = InStr(GrdTaxas.ListItems(i).SubItems(1), "-") - 1
    '        If String_Taxas = "" Then
    '            String_Taxas = String_Taxas & " [ " & Left(GrdTaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(GrdTaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
    '        Else
    '            String_Taxas = String_Taxas & ", [ " & Left(GrdTaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(GrdTaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
    '        End If
    '        Total_Taxas = Total_Taxas + CCur(GrdTaxas.ListItems(i).SubItems(2))
    '    End If
    'Next
End Sub

Private Sub cboImposto_LostFocus()
    
    txtValor = 0
    txtVence = Date
    txtParcela = 0
    txtPeriodoInicial = Format(Month(Date), "00") & Year(Date)
    
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    cboImposto.SetFocus
    
End Sub

Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIM
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Resultado As Boolean
    Dim Qtd As String
    Dim Tipo As TipoInscricaoObrigacao
    Dim InsCad As String, Grupo As Byte 'criado para a utilizacao do grupo de inscruicao
    If txtIM <> "" Then
        Tipo = etiContribuinte
    ElseIf txtImovel <> "" Then
        Tipo = etiImovel
    End If
    
    If cboImposto.ListIndex = -1 Then
        Util.Avisa "Selecione tributo."
        cboImposto.SetFocus
        Exit Sub
    End If
    
    If txtIM = "" And txtImovel = "" Or txtIM <> "" And txtImovel <> "" Then
        Avisa "Informe Inscrição Municipal ou Inscrição Cadastro."
        txtIM.SetFocus
        Exit Sub
    End If
    
    If txtPeriodoInicial = "" Then
        Util.Avisa "Informe " & txtPeriodoInicial.Caption & "."
        txtPeriodoInicial.SetFocus
        Exit Sub
    End If
    
    If txtVence = "" Then
        Util.Avisa "Informe Vencimento."
        txtVence.SetFocus
        Exit Sub
    End If
    
    If txtValor = "" Then
        Util.Avisa "Informe Valor."
        txtValor.SetFocus
        Exit Sub
    End If
    
    
    Screen.MousePointer = 11
    Grupo = IIf(Trim(txtImovel) = "", 0, 1)
'    If cboImposto.Coluna(0).Valor = Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQNFIXO)) Then
'        Obrig.GeraExtratoUnificado = Confirma("Deseja gerar extrato unificado de débitos para o contribuinte?")
'    End If
    If Obrig.CriaObrigacao(CStr(cboImposto.Coluna(0).Valor), txtPeriodoInicial, _
                txtPeriodoInicial, txtIM, txtValor, etsCreditoOriginalAberto, etsCriaNova, _
                txtVence, , , Grupo, , , , txtOrigem, Nvl(txtParcela, 0), txtImovel, Tipo) Then
        Informa "Obrigação gerada com sucesso."
         If Len(Obrig.obCodigoObrigacao) > 0 Then
            Bdados.Executa ("update tab_obrigacao_contribuinte set toc_observacao='" & txtObs & "' where toc_cod_obrigacao=" & Obrig.obCodigoObrigacao)
            If chkRemessa.Value = 0 Then 'false
                Bdados.Executa ("update tab_obrigacao_contribuinte set toc_remessa=99 where toc_cod_obrigacao=" & Obrig.obCodigoObrigacao)
            End If
         End If
         If Not Obrig.MostraObrigacaoGerada(lstObrig, CStr(cboImposto.Coluna(0).Valor), txtIM, _
            , , , , _
             txtPeriodoInicial, txtPeriodoInicial, , txtImovel, , IIf(Temp.PegaParametro(Bdados, "TRAZER SUBDIVIDA") = "SIM", True, False)) Then
            Avisa "Nenhum registro encontrado."
        End If
         If lstObrig.ListItems.Count > 0 Then lstObrig.SelectedItem.Checked = True
         cboImposto.SetFocus
    Else
        Informa "Não foi possivel gerar a(s) obrigacão. Verifique se o contribuinte está sujeito ao tributo, se ja existem pagamentos para neste periodo, ou caso imóvel, verifique valores de infra-estrutura(trecho) e da PGV."
    End If
    
    Screen.MousePointer = 0
'    cmdLimpar_Click
End Sub

Private Sub cmdVISUAL1_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub cmdVISUAL2_Click()
On Error GoTo trata
    
    If lstObrig.ListItems.Count = 0 Then Exit Sub
    
    If lstObrig.ListItems.Count > 1 Then
        If Not Util.Confirma("Confirma impressão de " & lstObrig.ListItems.Count & " obrigações") Then Exit Sub
    Else
        If Not Util.Confirma("Confirma impressão da obrigação") Then Exit Sub
    End If

               
    Screen.MousePointer = 11
    Dim i As Double
    For i = 1 To lstObrig.ListItems.Count
        With lstObrig.ListItems
            .Item(i).Selected = True
            If .Item(i).Checked Then
                Call Pega_taxas
                If Trim(txtImovel) = "" Then
                    ImprimeSelecionado lstObrig, txtRazao, txtEndereco, False, , tdiImpressora, String_Taxas, Total_Taxas, txtIM, txtEndereco
                Else
                    ImprimeSelecionado lstObrig, txtRazao, txtEndereco, False, , tdiImpressora, String_Taxas, Total_Taxas, txtImovel, txtEndereco
                End If
            End If
        End With
        DoEvents
    Next
            
    Avisa "Impressão concluída."
    Screen.MousePointer = 0
    
    Exit Sub
trata:
    Screen.MousePointer = 0
    Erro Err.Description
    
End Sub

Private Sub Form_Activate()
    Obrig.PreencheComboTributo cboImposto, False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
    txtObs = ""
    
End Sub

Private Sub Image1_Click()

End Sub

Private Sub txtIm_LostFocus()
 Dim Ic As String
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(txtIM) = 10 Or Len(txtIM) = 11 Then
            If InStr(1, txtIM, "-") = 0 Then
                Ic = Imposto.FormataInscricao(txtIM, InscContrib)
            Else
                Ic = txtIM
            End If
        Else
            Ic = txtIM
        End If
    Else
            Ic = txtIM
    End If
    txtIM = BuscaContribuinte(Ic, txtRazao, txtEndereco)
End Sub

Private Sub txtImovel_LostFocus()
  Dim Ic As String
  
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtIM.SetFocus
        End If
    End If
End Sub

Private Sub txtParcela_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{TAB}"
'    End If
End Sub
