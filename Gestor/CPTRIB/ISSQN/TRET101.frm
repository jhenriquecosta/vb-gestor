VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TRET101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL lstPesq 
      Height          =   1950
      Left            =   60
      TabIndex        =   24
      Top             =   4200
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   3440
      CorBorda        =   32768
      Caption         =   "Lista de Pesquisa"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin VTOcx.fraVISUAL fraDeducao 
      Height          =   1470
      Left            =   45
      TabIndex        =   21
      Top             =   2655
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   2593
      Altura          =   1905
      Caption         =   " Deduções"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483626
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtAliq 
         Height          =   480
         Left            =   8610
         TabIndex        =   26
         Tag             =   "Periodo"
         Top             =   300
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   847
         Caption         =   "Aliquota"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoRotulo=   1
         MaxLen          =   7
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtBase 
         Height          =   480
         Left            =   5865
         TabIndex        =   23
         Top             =   870
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   847
         Caption         =   "Base de Cálculo"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         AlinhamentoRotulo=   1
         MaxLen          =   60
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtNumNota 
         Height          =   480
         Left            =   345
         TabIndex        =   9
         Tag             =   "Nota Fiscal"
         Top             =   345
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   847
         Caption         =   "N° Nota Fiscal"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         MaxLen          =   60
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtDtEmissao 
         Height          =   480
         Left            =   3000
         TabIndex        =   10
         Tag             =   "Data Emissao"
         Top             =   345
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   847
         Caption         =   "Data Emissão"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         AlinhamentoRotulo=   1
         MaxLen          =   60
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtTotalNota 
         Height          =   480
         Left            =   345
         TabIndex        =   12
         Tag             =   "Total Notas"
         Top             =   870
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   847
         Caption         =   "Total da Nota"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoRotulo=   1
         MaxLen          =   60
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtMaterial 
         Height          =   480
         Left            =   3000
         TabIndex        =   13
         Top             =   870
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   847
         Caption         =   "Vl. Material ICMS"
         Text            =   ""
         Formato         =   5
         Restricao       =   3
         AlinhamentoRotulo=   1
         MaxLen          =   30
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtISS 
         Height          =   480
         Left            =   8610
         TabIndex        =   22
         Top             =   870
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   847
         Caption         =   "ISS Devido"
         Text            =   ""
         Enabled         =   0   'False
         Formato         =   5
         AlinhamentoRotulo=   1
         MaxLen          =   30
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtPeriodo 
         Height          =   480
         Left            =   5865
         TabIndex        =   11
         Tag             =   "Periodo"
         Top             =   345
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   847
         Caption         =   "Período Referência"
         Text            =   ""
         AlinhamentoRotulo=   1
         MaxLen          =   7
         RetirarMascara  =   0   'False
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1138
      Icone           =   "TRET101.frx":0000
   End
   Begin VTOcx.fraVISUAL fraPrestador 
      Height          =   1950
      Left            =   45
      TabIndex        =   20
      Top             =   690
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   3440
      Altura          =   1905
      Caption         =   " Prestador de Serviço"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483626
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   465
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   635
         Caption         =   "&Novo"
         Acao            =   6
      End
      Begin VTOcx.cmdVISUAL cmdPesq 
         Height          =   360
         Left            =   9975
         TabIndex        =   2
         Top             =   465
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Caption         =   ""
         Acao            =   5
      End
      Begin VTOcx.cboVISUAL cboAtividade 
         Height          =   510
         Left            =   4620
         TabIndex        =   8
         Top             =   1335
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   900
         Caption         =   "Atividade Economica"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboUF_Rem 
         Height          =   510
         Left            =   9555
         TabIndex        =   6
         Top             =   825
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
         Caption         =   "UF"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.txtVISUAL txtBairro_Rem 
         Height          =   480
         Left            =   4620
         TabIndex        =   4
         Top             =   840
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   847
         Caption         =   "Bairro"
         Text            =   ""
         AlinhamentoRotulo=   1
         MaxLen          =   60
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtCep_Rem 
         Height          =   480
         Left            =   8055
         TabIndex        =   5
         Top             =   840
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   847
         Caption         =   "CEP"
         Text            =   ""
         Formato         =   4
         Restricao       =   2
         AlinhamentoRotulo=   1
         MaxLen          =   10
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtMunicipio_Rem 
         Height          =   480
         Left            =   120
         TabIndex        =   7
         Top             =   1350
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   847
         Caption         =   "Município"
         Text            =   ""
         AlinhamentoRotulo=   1
         MaxLen          =   60
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco_Rem 
         Height          =   480
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   847
         Caption         =   "Endereço"
         Text            =   ""
         AlinhamentoRotulo=   1
         MaxLen          =   80
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtNome_Rem 
         Height          =   480
         Left            =   3240
         TabIndex        =   1
         Tag             =   "Nome"
         Top             =   330
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   847
         Caption         =   "Nome Empresarial"
         Text            =   ""
         AlinhamentoRotulo=   1
         MaxLen          =   80
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtImRem 
         Height          =   480
         Left            =   1095
         TabIndex        =   0
         Tag             =   "CPF/CNPJ ou IM"
         Top             =   330
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   847
         Caption         =   "CPF/CNPJ ou IM"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         MaxLen          =   20
         RetirarMascara  =   0   'False
      End
   End
   Begin Cabecalho.rodVISUAL rod 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   25
      Top             =   6195
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   6915
         TabIndex        =   14
         Top             =   120
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9375
         TabIndex        =   16
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   8205
         TabIndex        =   15
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4410
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   18
      Top             =   45
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "TRET101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NovoRemetente As Boolean
Dim NovoDestino As Boolean
Dim Aliquota As Double
Dim Imposto As New VSImposto
Dim CodPagamento As String
Dim CodImposto As String
Dim NomeImposto As String
Dim NumCGC As String
Dim NumIM As String
Dim Retencao As cRetencao
Dim Nota As cNota
'
'Private Function NumPagamentoRet(Contribuinte As String, Periodo As Long, Nota As String, CodImposto As String) As Double
'    Dim Sql As String
'    Dim Rs As VSRecordset
'    Dim Conta As New ContaCorrente
'    Sql = "Select tna_cod_pagamento  from Tab_Nota_Avulsa " & _
'        " where  tna_periodo=" & Periodo & _
'        " and tna_tca_identidade_remetente='" & Contribuinte & "' and tna_numero_nota =" & Nota
'    If Not Bdados.AbreTabela(Sql, Rs) Then
'        NumPagamentoRet = Conta.GeraCodPagamento(CodImposto)
'    Else
'        NumPagamentoRet = Rs(0)
'    End If
'
'End Function

'Private Function GravaDadosBaixa() As Boolean
'    On Error GoTo trata
'    Dim Valores As String
'    Dim Campos As String
'    Dim Sql As String
'    Dim PeriodoImposto As String
'    Dim Conta As New ContaCorrente
'    Dim Rs As VSRecordset
'    PeriodoImposto = txtPeriodo
'
'
'    GravaDadosBaixa = True
'    Valores = Bdados.PreparaValor(txtimRem, CodImposto, Bdados.Converte(Date, TCDataHora), _
'    IIf(Len(txtPeriodo) = 4, txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2)), Bdados.Converte(Date, TCDataHora), CDbl(txtISS), _
'     Bdados.Converte(txtISS, TCDuplo), Bdados.Converte(Date, TCDataHora), Aplicacoes.Usuario, CodPagamento, Bdados.Converte(txtISS, TCDuplo))
'
'    Campos = "tdr_im,tdr_tip_cod_imposto,tdr_data_vencimento,tdr_periodo," & _
'        "tdr_data_pagamento,tdr_valor_original," & _
'        "tdr_valor_total,tdr_data_entrada,tdr_tus_cod_usuario,tdr_tgt_cod_pagamento,tdr_valor_real_pago"
'    Call Bdados.GravaDados("Tab_Darm_Recebido", Valores, Campos, "tdr_tgt_cod_pagamento=" & CodPagamento)
'    Exit Function
'trata:
'    GravaDadosBaixa = False
'End Function

Sub BuscaAliquota()
    With Retencao
        .BuscaAliquota Date
        Aliquota = .Nota.Aliquota
        CodImposto = .Nota.Cod_Imposto
        NomeImposto = .Nota.Nome_Imposto
    End With
End Sub

Sub HabilitaRemetente(Status As Boolean)
    txtNome_Rem.Enabled = Status
    txtEndereco_Rem.Enabled = Status
    txtBairro_Rem.Enabled = Status
    txtCep_Rem.Enabled = Status
    cboUF_Rem.Enabled = Status
    txtMunicipio_Rem.Enabled = Status
    cboAtividade.Enabled = Status
End Sub

Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    txtImRem.Enabled = True
    txtImRem.SetFocus
End Sub

Private Sub cmdOpcao_Click()
    NovoRemetente = True
    txtImRem = ""
    HabilitaRemetente True
    LimpaCampos Me
    txtImRem.Enabled = True
    txtImRem.SetFocus
    lstPesq.ListItems.Clear
End Sub

Private Sub cmdOpcao_LostFocus()
    txtImRem.SetFocus
End Sub

Private Sub cmdPesq_Click()
    Screen.MousePointer = 11
    DoEvents
    Retencao.PreencheGrid lstPesq, txtNome_Rem
    lstPesq.SetFocus
    Screen.MousePointer = 0
    DoEvents
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If Not CriticaCampos(Me) Then Exit Sub
    Screen.MousePointer = 11
    '==============================================
    DoEvents
    With Retencao
        .Nota.IM_CPF = txtImRem
        .Nota.Nome_Empresa = txtNome_Rem
        .Nota.Endereco.Endereco = txtEndereco_Rem
        .Nota.Endereco.Bairro = txtBairro_Rem
        .Nota.Endereco.CEP = txtCep_Rem
        .Nota.Endereco.UF = CStr(cboUF_Rem.Coluna(0).Valor)
        .Nota.Endereco.Municipio = txtMunicipio_Rem
        .Nota.Atividade = CStr(cboAtividade.Coluna(1).Valor)
        .Nota.Nota_fiscal = txtNumNota
        .Nota.Data_emissao = txtDtEmissao
        .Nota.Periodo_Ref = txtPeriodo
        .Nota.Total_Nota = IIf(txtTotalNota = "", "0", txtTotalNota)
        .Nota.Valor_Material_ICMS = IIf(txtMaterial = "", "0", txtMaterial)
        .Nota.ISS_Devido = txtISS
        .Nota.Usuario = Aplicacoes.Usuario
        .Nota.Cod_Imposto = CodImposto
        .Nota.Aliquota = txtAliq / 100
        .Nota.Cod_Pagamento = CodPagamento
        If .Salvar(NovoRemetente, NumIM, Date) Then
        CodPagamento = .Nota.Cod_Pagamento
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Dim Cobranca As New VSCobranca
            Cobranca.ImprimeDam Rpt, CDbl(CodPagamento), NumIM, txtNome_Rem, NumCGC, txtEndereco_Rem, "", "", _
                                CodImposto, Imposto.NomeTributo(ttr_ISSQN), NomeImposto, _
                                txtPeriodo, 0, 1, UltimoDiaDoMes(Date), txtBase, txtISS, 0, 0, 0, 0, CStr(cboAtividade.Coluna(0).Valor), "Prestador dos Serviços: " & txtNome_Rem, PicBarra, txtNumNota, txtNumNota, txtMaterial
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Screen.MousePointer = 0
            DoEvents
            Util.Informa "Transação Finalizada"
        Else
            Screen.MousePointer = 0
            DoEvents
            Util.Erro "Operação Cancelada: Erro de Gravação!"
        End If
    End With
End Sub

Private Sub Form_Load()
    Set Retencao = New cRetencao
    Set Nota = New cNota
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rod.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    '------------------------preenche combos----------------------------------
    Nota.PreencherCboAtividade cboAtividade
    cboUF_Rem.PreencherGeral Bdados, "UF"
    '-------------------------------------------------------------------------------
    HabilitaRemetente True
    BuscaAliquota
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Retencao = Nothing
    Set Nota = Nothing
End Sub

Private Sub lstPesq_DblClick()
    If lstPesq.SelectedItem Is Nothing Then Exit Sub
    txtImRem = lstPesq.SelectedItem
    Call txtimRem_LostFocus
    txtImRem.Enabled = False
    txtNumNota.SetFocus
End Sub

Private Sub txtBairro_Rem_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBase_Change()
    On Error Resume Next
    txtISS = CStr(CDbl(txtBase) * CDbl(txtAliq) / 100)
End Sub

Private Sub txtDtEmissao_LostFocus()
    If UCase(Me.ActiveControl.Name) = "CMDSALVAR" Or UCase(Me.ActiveControl.Name) = "CMDSAIR" Or UCase(Me.ActiveControl.Name) = "CMDNOVO" Then Exit Sub
    If IsNumeric(txtDtEmissao) Then txtDtEmissao = Edita.FormataTexto(txtDtEmissao, Data)
    If IsDate(txtDtEmissao) Then
        If Len(txtDtEmissao) <> 10 Then Exit Sub
        If CDbl(Right(txtDtEmissao, 4) & Mid(txtDtEmissao, 4, 2) & Left(txtDtEmissao, 2)) > CDbl(Right(Date, 4) & Mid(Date, 4, 2) & Left(Date, 2)) Then
            Avisa "Data de emissão da nota não pode ser superior a atual."
            txtDtEmissao.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtEndereco_Rem_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtimRem_LostFocus()
    If Trim(txtImRem) = "" Then Exit Sub
    If Not AplicacoesVTFuncoes.Municipio = "PETROLINA" Then
        If Len(txtImRem) = 10 And IsNumeric(txtImRem) Then
            txtImRem = Imposto.FormataInscricao(txtImRem, InscContrib)
        ElseIf Len(txtImRem) = 11 And IsNumeric(txtImRem) Then
            txtImRem = Edita.FormataTexto(txtImRem, Cpf)
        ElseIf Len(txtImRem) = 14 Then
            txtImRem = Edita.FormataTexto(txtImRem, Cgc)
        End If
    End If
    If Trim(txtImRem) = "" Then Exit Sub
    '-------------------------------------------------------------------------------
    Screen.MousePointer = 11
    DoEvents
    With Nota
        If .Buscar(txtImRem, txtImRem) Then
            HabilitaRemetente False
            If txtImRem = Const_ImAvulso Then
                txtEndereco_Rem = ""
                txtBairro_Rem = ""
                txtCep_Rem = ""
                cboUF_Rem.ListIndex = -1
                txtMunicipio_Rem = ""
                cboAtividade.ListIndex = -1
            Else
                txtEndereco_Rem = .Endereco.Endereco
                txtBairro_Rem = .Endereco.Bairro
                txtCep_Rem = .Endereco.CEP
                cboUF_Rem.SetarLinha .Endereco.UF, 0
                txtMunicipio_Rem = .Endereco.Municipio
                cboAtividade.SetarLinha .Atividade, 1
            End If
            NumCGC = .NumCGC
            NumIM = .NumIM
            txtNome_Rem = .Nome_Empresa
            NovoRemetente = False
            txtNumNota.SetFocus
        Else
            HabilitaRemetente True
            NovoRemetente = True
            '--------------------------------------------------------
            txtImRem.Tag = txtImRem
            LimpaCampos Me
            txtImRem = txtImRem.Tag
            txtImRem.Tag = "CPF/CNPJ ou IM"
            '--------------------------------------------------------
            txtNome_Rem.SetFocus
            NumCGC = txtImRem
        End If
    End With
    Screen.MousePointer = 0
End Sub

Private Sub txtMaterial_Change()
    On Error Resume Next
    txtBase = CDbl(Nvl(txtTotalNota, 0)) - CDbl(Nvl(txtMaterial, 0))
'    txtBase = Edita.FormataTexto(txtBase, Monetario, True)
End Sub


Private Sub txtMaterial_LostFocus()
    If CDbl(Nvl(txtTotalNota, 0)) < CDbl(Nvl(txtMaterial, 0)) Then
        Avisa "Valor não pode ser maior que Total em notas."
        txtMaterial = "0,00"
        txtMaterial.SetFocus
    End If
End Sub

Private Sub txtMunicipio_Rem_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNome_Rem_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPeriodo_LostFocus()
    If Trim(txtPeriodo) = "" Then Exit Sub
    If IsNumeric(txtPeriodo) Then
        If Len(txtPeriodo) <> 6 Then
            Avisa "Período inválido."
            txtPeriodo = ""
            txtPeriodo.SetFocus
            Exit Sub
        Else
            Aliquota = Imposto.BuscaAliquota(CodImposto, Right(txtPeriodo, 4))
            txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
        End If
    End If
    txtAliq = Aliquota * 100
End Sub

Private Sub txtTotalNota_Change()
    On Error Resume Next
    txtBase = CDbl(Nvl(txtTotalNota, 0)) - CDbl(Nvl(txtMaterial, 0))
End Sub

