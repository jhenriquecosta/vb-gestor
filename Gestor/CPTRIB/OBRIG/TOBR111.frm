VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TOBR111 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TOBR111"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   4470
      Left            =   0
      TabIndex        =   20
      Top             =   630
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   7885
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      TagVariant      =   ""
      Tabs            =   "TOBR111.frx":0000
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4080
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   7197
         _Version        =   131082
         TabGuid         =   "TOBR111.frx":007E
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
            Height          =   5445
            Index           =   3
            Left            =   -75
            TabIndex        =   24
            Top             =   -105
            Width           =   11235
            Begin VTOcx.cboVISUAL cboTipo 
               Height          =   315
               Left            =   3060
               TabIndex        =   15
               Tag             =   "Tipo Isenção"
               Top             =   2460
               Width           =   5025
               _ExtentX        =   8864
               _ExtentY        =   556
               Caption         =   "Tipo Isenção"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL CboTributo 
               Height          =   315
               Left            =   675
               TabIndex        =   11
               Top             =   1005
               Width           =   7395
               _ExtentX        =   13044
               _ExtentY        =   556
               Caption         =   "Imposto"
               Text            =   ""
               AutoFocaliza    =   0   'False
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
               Left            =   1440
               TabIndex        =   25
               Top             =   2070
               Width           =   6615
            End
            Begin VTOcx.txtVISUAL txtIm 
               Height          =   300
               Left            =   615
               TabIndex        =   12
               Tag             =   "Inscricão"
               Top             =   1380
               Width           =   2925
               _ExtentX        =   5159
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
               Left            =   300
               TabIndex        =   29
               Top             =   1725
               Width           =   7740
               _ExtentX        =   13653
               _ExtentY        =   529
               Caption         =   "Nome/Razão"
               Text            =   ""
               Enabled         =   0   'False
               Requerido       =   0   'False
            End
            Begin VTOcx.txtVISUAL txtPeriodoInicial 
               Height          =   300
               Left            =   180
               TabIndex        =   14
               Tag             =   "Periodo Inicial"
               Top             =   2460
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   529
               Caption         =   "Periodo Inicial"
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cmdVISUAL cmdPesquisaInscricao 
               Height          =   315
               Left            =   3570
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   1380
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtImovel 
               Height          =   300
               Left            =   3960
               TabIndex        =   13
               Tag             =   "Cadastro do Imóvel"
               Top             =   1395
               Width           =   3690
               _ExtentX        =   6509
               _ExtentY        =   529
               Caption         =   "Cadastro do Imóvel"
               Text            =   ""
               Requerido       =   0   'False
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cmdVISUAL cmdVISUAL1 
               Height          =   315
               Left            =   7695
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   1380
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               Caption         =   ""
               Acao            =   5
            End
            Begin VB.Label LblPercento 
               AutoSize        =   -1  'True
               Height          =   195
               Left            =   6150
               TabIndex        =   28
               Top             =   1290
               Width           =   45
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4080
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   7197
         _Version        =   131082
         TabGuid         =   "TOBR111.frx":00A6
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1920
            Left            =   30
            TabIndex        =   30
            Top             =   15
            Width           =   8325
            _ExtentX        =   14684
            _ExtentY        =   3387
            Altura          =   1905
            Caption         =   " Consultar Por:"
            CorTexto        =   0
            CorFaixa        =   8421504
            CorFundo        =   -2147483626
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtPeriodoFinalConsulta 
               Height          =   300
               Left            =   2565
               TabIndex        =   4
               Top             =   1545
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   529
               Caption         =   "Periodo Inicial"
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.txtVISUAL txtPeriodoInicialConsulta 
               Height          =   300
               Left            =   105
               TabIndex        =   3
               Top             =   1545
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   529
               Caption         =   "Periodo Inicial"
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               MinLen          =   4
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cboVISUAL cboTributoConsulta 
               Height          =   315
               Left            =   585
               TabIndex        =   0
               Top             =   420
               Width           =   7395
               _ExtentX        =   13044
               _ExtentY        =   556
               Caption         =   "Imposto"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cboVISUAL CboTipoConsulta 
               Height          =   315
               Left            =   5055
               TabIndex        =   5
               Top             =   1530
               Width           =   2910
               _ExtentX        =   5133
               _ExtentY        =   556
               Caption         =   "Isenção"
               Text            =   ""
               AutoFocaliza    =   0   'False
            End
            Begin VTOcx.cmdVISUAL cmdVISUAL2 
               Height          =   315
               Left            =   7605
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   810
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtImovelConsulta 
               Height          =   300
               Left            =   3870
               TabIndex        =   2
               Top             =   825
               Width           =   3690
               _ExtentX        =   6509
               _ExtentY        =   529
               Caption         =   "Cadastro do Imóvel"
               Text            =   ""
               Requerido       =   0   'False
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
            Begin VTOcx.cmdVISUAL cmdPesquisaInscricaoConsulta 
               Height          =   315
               Left            =   3480
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   810
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               Caption         =   ""
               Acao            =   5
            End
            Begin VTOcx.txtVISUAL txtRazaoConsulta 
               Height          =   300
               Left            =   210
               TabIndex        =   31
               Top             =   1185
               Width           =   7740
               _ExtentX        =   13653
               _ExtentY        =   529
               Caption         =   "Nome/Razão"
               Text            =   ""
               Enabled         =   0   'False
               Requerido       =   0   'False
            End
            Begin VTOcx.txtVISUAL txtImConsulta 
               Height          =   300
               Left            =   525
               TabIndex        =   1
               Top             =   825
               Width           =   2925
               _ExtentX        =   5159
               _ExtentY        =   529
               Caption         =   "Inscricão"
               Text            =   ""
               Restricao       =   2
               Requerido       =   0   'False
               RetirarMascara  =   0   'False
               AutoTAB         =   -1  'True
            End
         End
         Begin VTOcx.grdVISUAL lstAtv 
            Height          =   2130
            Left            =   30
            TabIndex        =   23
            Top             =   1965
            Width           =   8340
            _ExtentX        =   14711
            _ExtentY        =   3757
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   192
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   19
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TOBR111.frx":00CE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   18
      Top             =   5175
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7245
         TabIndex        =   10
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sair"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   3780
         TabIndex        =   7
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Salvar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   4935
         TabIndex        =   8
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Excluir"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6090
         TabIndex        =   9
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdBusca 
         Height          =   375
         Left            =   2625
         TabIndex        =   6
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   930
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   1138
      Icone           =   "TOBR111.frx":21F1
   End
End
Attribute VB_Name = "TOBR111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
Public Obrig As New OBRIGACAO
Private Sub cmdBusca_Click()
    Dim Sql As String
    Dim RsPref As VSRecordset
    Dim RsCTM As VSRecordset
    Dim Anterior As String
    Sql = "Select TIS_TCI_IM as Contribuinte, TIS_PERIODO as Periodo," & _
        "TIS_TIPO_ISENSAO as Tipo,TIS_TIP_COD_IMPOSTO +' - '+  Tip_Sigla_Imposto as Tributo,tis_tipo_inscricao as Inscricao from Tab_Isento, Tab_Imposto where TIS_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
    
    If Trim(cboTributoConsulta) <> "" Then Sql = Sql & " and TIS_TIP_COD_IMPOSTO ='" & cboTributoConsulta.Coluna(0).Valor & "'"
    If Trim(CboTipoConsulta) <> "" Then Sql = Sql & " and TIS_TIPO_ISENSAO = '" & CboTipoConsulta.Coluna(1).Valor & "'"
    If txtImConsulta <> "" Then
        Sql = Sql & " and TIS_TCI_IM  = '" & txtImConsulta & "'"
    ElseIf txtImovelConsulta <> "" Then
        Sql = Sql & " and TIS_TCI_IM  = '" & txtImovelConsulta & "'"
    End If
    If Trim(txtPeriodoInicialConsulta) <> "" And Trim(txtPeriodoFinalConsulta) <> "" Then
        Sql = Sql & " and TIS_periodo >= " & txtPeriodoInicialConsulta & " and TIS_periodo <= " & txtPeriodoFinalConsulta
    ElseIf txtPeriodoInicialConsulta <> "" And txtPeriodoFinalConsulta = "" Then
        Sql = Sql & " and TIS_periodo >= " & txtPeriodoInicialConsulta & " and TIS_periodo <= " & txtPeriodoInicialConsulta
    End If
    If lstAtv.Preencher(Bdados, Sql, 1300, 2000, 900, 3000, 0) Then
        'lstAtv.Mensagem = "Total de Isenção: R$" & Format(lstAtv.Colunas(4).Soma, Const_Monetario)
    End If
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdExcluir_Click()
    Dim Condicao As String
    Dim Tipo  As TipoInscricaoObrigacao
    Dim Contribuinte As String
    If txtIM <> "" Then
        Contribuinte = txtIM
        Tipo = etiContribuinte
    Else
        Tipo = etiImovel
        Contribuinte = txtImovel
    End If
    Condicao = "tIS_PERIODO = '" & txtPeriodoInicial & "' and TIS_TIP_COD_IMPOSTO =  '" & CboTributo.Coluna(0).Valor & "' and TIS_TCI_IM = '" & Contribuinte & "' and tis_tipo_inscricao = '" & Tipo & "'"
    If Confirma("Confirma a exclusão?", "Aviso") Then
        If CriticaCampos(Me) = False Then Exit Sub
        If Bdados.DeletaDados("tab_isento", Condicao) Then
            Avisa "Operação concluída com sucesso."
            TabDados.Tabs(1).Selected = True
            cmdLimpar_Click
            cmdBusca_Click
            
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    lstAtv.Preencher Bdados, ""
    
End Sub





Private Sub cmdPesquisaInscricao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIM
End Sub

Private Sub cmdPesquisaInscricaoConsulta_Click()
  AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtImConsulta
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub


Private Sub cmdSalvar_Click()
    Dim Valores  As String
    Dim Campos As String
    Dim Condicao As String
    Dim Contribuinte As String
    Dim Tipo As TipoInscricaoObrigacao
    txtIM.Tag = ""
    txtImovel.Tag = ""
    
    If txtIM = "" And txtImovel = "" Or txtIM <> "" And txtImovel <> "" Then
        Util.Avisa "Informe " & txtIM.Caption & " ou " & txtImovel.Caption
    End If
    
    If CriticaCampos(Me) = False Then Exit Sub
    If txtIM <> "" Then
        Contribuinte = txtIM
        Tipo = etiContribuinte
    Else
        Tipo = etiImovel
        Contribuinte = txtImovel
    End If
    Valores = Bdados.PreparaValor(txtPeriodoInicial, Contribuinte, CboTributo.Coluna(0).Valor, cboTipo.Coluna(1).Valor, Bdados.Converte(Date, TCDataHora), AplicacoesVTFuncoes.Usuario, Tipo)
    Campos = "tIS_PERIODO,TIS_TCI_IM,TIS_TIP_COD_IMPOSTO,TIS_TIPO_ISENSAO,TIS_DATA,TIS_TUS_COD_USUARIO,TIS_TIPO_INSCRICAO"
    Condicao = "tIS_PERIODO = '" & txtPeriodoInicial & "' and TIS_TIP_COD_IMPOSTO =  '" & CboTributo.Coluna(0).Valor & "' and TIS_TCI_IM = '" & Contribuinte & "'"
    If Bdados.GravaDados("Tab_isento", Valores, Campos, Condicao) Then
        Util.Avisa "Operação concluída com sucesso."
        cmdLimpar_Click
        cmdBusca_Click
        TabDados.Tabs(1).Selected = True
    End If
End Sub

Private Sub cmdVISUAL1_Click()
AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovel
End Sub

Private Sub cmdVISUAL2_Click()
AplicacoesVTFuncoes.BuscaInscricao InscImovel, txtImovelConsulta
End Sub

Private Sub Form_Activate()
    Dim Sql As String
        
    AtualizaCabecalho lstAtv
    '1 - Isento por Limite da Base
    '2 - Isento de Imposto
    '3 - Isento por Limite Tributo
    '4 - Isento Total
    '5 - Imune
    Obrig.PreencheComboTributo CboTributo, False
    Obrig.PreencheComboTributo cboTributoConsulta, False
End Sub

Private Sub Form_Load()
    On Error Resume Next
    cabVISUAL.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    
    cboTipo.PreencherGeral Bdados, "TIPO ISENCAO"
    CboTipoConsulta.PreencherGeral Bdados, "TIPO ISENCAO"
    
End Sub

Private Sub Timer_Timer()
    On Error Resume Next
    
    
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub


Private Sub Timer1_Timer()

End Sub

Private Sub txtDescAtiv_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtMult_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = Edita.AceitaDig(KeyAscii, Valores)
End Sub


Private Sub txtExercicio1_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtExercicio2_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub lstAtv_DblClick()
    If lstAtv.ListItems.Count >= 1 Then
        txtIM = ""
        txtImovel = ""
        txtPeriodoInicial = lstAtv.SelectedItem.SubItems(1)
        CboTributo.SetarLinha Left(lstAtv.SelectedItem.SubItems(3), 8)
        If lstAtv.SelectedItem.SubItems(4) = 1 Then
            txtImovel = lstAtv.SelectedItem
            txtImovel_LostFocus
        Else
            txtIM = lstAtv.SelectedItem
            txtIm_LostFocus
        End If
        cboTipo.SetarLinha lstAtv.SelectedItem.SubItems(2), 1
        TabDados.Tabs(2).Selected = True
    End If
End Sub

Private Sub txtIm_LostFocus()
 Dim Ic As String
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        If Len(txtIM) = 10 Or Len(txtIM) = 11 Then
            Ic = Imposto.FormataInscricao(txtIM, InscContrib)
        Else
            Ic = txtIM
        End If
    Else
            Ic = txtIM
    End If
    txtIM = BuscaContribuinte(Ic, txtRazao, txtEndereco)
End Sub

Private Sub txtImConsulta_LostFocus()
 Dim Ic As String
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        If Len(txtImConsulta) = 10 Or Len(txtImConsulta) = 11 Then
            Ic = Imposto.FormataInscricao(txtImConsulta, InscContrib)
        Else
            Ic = txtImConsulta
        End If
    Else
            Ic = txtImConsulta
    End If
    txtIM = BuscaContribuinte(Ic, txtRazaoConsulta)
End Sub

Private Sub txtImovel_LostFocus()
    On Error Resume Next
  Dim Ic As String
  
    If Trim(txtImovel) <> "" Then
        txtImovel = BuscaContribuinte(txtImovel, txtRazao, txtEndereco, , etiImovel)
        If Trim(txtImovel) = "" Then
            Avisa "Inscricão não encontrada"
            txtIM.SetFocus
        End If
    End If
End Sub

Private Sub txtImovelConsulta_LostFocus()
 Dim Ic As String
  
    If Trim(txtImovelConsulta) <> "" Then
        txtImovelConsulta = BuscaContribuinte(txtImovelConsulta, txtRazaoConsulta, , , etiImovel)
        If Trim(txtImovelConsulta) = "" Then
            Avisa "Inscricão não encontrada"
            txtImovelConsulta.SetFocus
        End If
    End If
End Sub
