VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TMCO201 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleMode       =   0  'User
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   19
      Top             =   6255
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   1032
      Begin VTOcx.cmdVISUAL CmdBuscar 
         Height          =   375
         Left            =   3210
         TabIndex        =   21
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   5790
         TabIndex        =   13
         Top             =   135
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Novo"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   4365
         TabIndex        =   12
         Top             =   135
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         Caption         =   "&Salvar DAM"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6945
         TabIndex        =   14
         Top             =   135
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
      Height          =   5610
      Left            =   0
      TabIndex        =   16
      Top             =   660
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   9895
      Caption         =   "DAM"
      Descricao       =   "Alterações no documento de arrecadação municipal"
      corFaixa        =   16711680
      Icone           =   "TMCO201.frx":0000
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.grdVISUAL Grid 
         Height          =   2490
         Left            =   60
         TabIndex        =   20
         Top             =   3390
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   4392
         CorBorda        =   16711680
         CorFundo        =   0
         CorTitulo       =   16711680
         CorCaption      =   -2147483634
         CorDica         =   16711680
      End
      Begin VTOcx.fraVISUAL fraVISUAL2 
         Height          =   825
         Left            =   60
         TabIndex        =   18
         Top             =   2550
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   1455
         Altura          =   1905
         Caption         =   " Detalhes"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtTaxas 
            Height          =   480
            Left            =   1650
            TabIndex        =   8
            Top             =   300
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   847
            Caption         =   "Taxas Acessórias"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtTotalImposto 
            Height          =   480
            Left            =   6390
            TabIndex        =   11
            Top             =   300
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   847
            Caption         =   "Total a Recolher"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtImposto 
            Height          =   480
            Left            =   75
            TabIndex        =   7
            Top             =   300
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   847
            Caption         =   "Imp. a Recolher"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtMulta 
            Height          =   480
            Left            =   4815
            TabIndex        =   10
            Top             =   300
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   847
            Caption         =   "Multa"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtJuros 
            Height          =   480
            Left            =   3225
            TabIndex        =   9
            Top             =   300
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   847
            Caption         =   "Juros"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoRotulo=   1
         End
      End
      Begin VTOcx.fraVISUAL fraVISUAL1 
         Height          =   1830
         Left            =   60
         TabIndex        =   17
         Top             =   690
         Width           =   7995
         _ExtentX        =   14102
         _ExtentY        =   3228
         Altura          =   1905
         Caption         =   " Informações Gerais"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VTOcx.txtVISUAL txtDam 
            Height          =   480
            Left            =   105
            TabIndex        =   0
            Top             =   285
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   847
            Caption         =   "Nº DAM"
            Text            =   ""
            Restricao       =   2
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtParcela 
            Height          =   480
            Left            =   5325
            TabIndex        =   4
            Top             =   765
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   847
            Caption         =   "Parcela"
            Text            =   ""
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtIm 
            Height          =   480
            Left            =   105
            TabIndex        =   1
            Top             =   765
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   847
            Caption         =   "Insc. Municipal"
            Text            =   ""
            Formato         =   8
            Restricao       =   2
            AlinhamentoRotulo=   1
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtDtVenc 
            Height          =   480
            Left            =   6420
            TabIndex        =   5
            Top             =   765
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   847
            Caption         =   "Data Vencimento"
            Text            =   ""
            Formato         =   0
            AlinhamentoRotulo=   1
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtPeriodo 
            Height          =   480
            Left            =   4260
            TabIndex        =   3
            Top             =   765
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   847
            Caption         =   "Exercício"
            Text            =   ""
            Restricao       =   2
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtIc 
            Height          =   480
            Left            =   2190
            TabIndex        =   2
            Top             =   765
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   847
            Caption         =   "Insc. Cadastral"
            Text            =   ""
            Restricao       =   2
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.cboVISUAL cboReceita 
            Height          =   510
            Left            =   105
            TabIndex        =   6
            Top             =   1260
            Width           =   7830
            _ExtentX        =   13811
            _ExtentY        =   900
            Caption         =   "Tributo"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Alinhamento     =   1
         End
      End
   End
   Begin VB.PictureBox PicBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4485
      ScaleHeight     =   465
      ScaleWidth      =   765
      TabIndex        =   15
      Top             =   1605
      Visible         =   0   'False
      Width           =   795
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   1138
      Icone           =   "TMCO201.frx":031A
   End
End
Attribute VB_Name = "TMCO201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto
Dim eDam As eDam
Dim CodImposto As String
Dim Exercicio As String
Dim Conta As New ContaCorrente
Dim Aliquota As Double
'Variaveis para o Report
Dim InscMuni As String
Dim RazaoSocial As String
Dim Documento As String
Dim Localizacao As String
Dim Data_Vencimento As String
Dim Codigo_Imovel As String
Dim Valor_Imposto As String
Dim CPFCNPJ As String
Dim Endereco As String
Dim Bairro As String
Dim Cod_Atividade As String
Dim Cod_Cidade As String
Dim Cep As String
Dim Uf As String
Dim Cod_Tributo As String
Dim Juro As String
Dim Multa As String
Dim TotalImposto As String
Dim TaxaServico As Double
Dim BaseDeCalculo As String
Dim VetLinhas(0 To 5) As String
Dim Linhas As Byte
Dim ObsAux As String
Dim NomeImposto As String
Dim TributoTaxa As Boolean
Dim TributoTaxaFixa As Double
Dim Tributo As Double
Dim Alvara As Double
Dim PosTraco As Byte
Dim TSU As Double
Dim AreaConstruida As Double
Dim AreaTotal As Double
Dim ValorTerreno As Double
Dim Valoredific As Double
Dim Zona As Integer
Dim ValorMetro As Double
Dim TaxaParcela As Double
Dim Desconto As String
Dim Reducao As String
Dim DtGeracao As String
Dim CodPagamento As Double

Private Sub cmdBuscar_Click()
    Dim Sql As String
    Sql = "Select tgt_im as [IM],"
    Sql = Sql & " tgt_tim_ic as [IC],"
    Sql = Sql & " tgt_periodo as Periodo,"
    Sql = Sql & " tgt_data_vencimento as Vencimento,"
    Sql = Sql & " tip_cod_imposto as [Cod Tributo],  "
    Sql = Sql & " tip_sigla_imposto as Descricao,"
    Sql = Sql & " cast(tgt_valor_tributo as decimal(14,2)) as [Vl Tributo], "
    Sql = Sql & " cast(tgt_taxa_expediente as decimal(14,2)) as [Taxas],  "
    Sql = Sql & " tgt_cod_pagamento as [Num DOC], "
    Sql = Sql & " tgt_cod_pagamento_vinculado as [Doc Vinculo], "
    Sql = Sql & " tgt_cod_pagamento_original as [Doc Origem],"
    Sql = Sql & " tgt_parcela as [Cota]  "
    Sql = Sql & " from  Tab_Geracao_Tributo INNER JOIN Tab_Imposto ON Tab_Geracao_Tributo.tgt_tip_cod_imposto = Tab_Imposto.tip_cod_imposto "
    Sql = Sql & " LEFT OUTER JOIN Tab_Darm_Recebido ON Tab_Geracao_Tributo.tgt_cod_pagamento = Tab_Darm_Recebido.tdr_tgt_cod_pagamento  "
    Sql = Sql & " where  tgt_tip_cod_imposto=tip_cod_imposto "
    Sql = Sql & " and (tgt_ativo =0 or tgt_ativo is null) "
    Sql = Sql & " and tgt_tip_cod_imposto not in ('NOTIFICA','EXTRATO') "
    Sql = Sql & " AND TIP_COD_IMPOSTO not in ('NOTIFICA','EXTRATO') "
    If txtIm <> "" Then
        Sql = Sql & " and tgt_im = " & Bdados.Converte(txtIm, tctexto)
    End If
    If txtIc <> "" Then
        Sql = Sql & " and tgt_tim_ic  = " & Bdados.Converte(txtIc, tctexto)
    End If
    Sql = Sql & " ORDER BY tgt_periodo, tip_sigla_imposto"
    Grid.Preencher Bdados, Sql
                
End Sub

Private Sub cmdSalvar_Click()
    Dim a As Integer
    Dim Valores As String
    Dim Campos As String
    Dim ValorImposto As Double
    Dim RsCob As VSRecordset
    Dim rs As VSRecordset
    Dim Sql As String
    Dim SqlParc As String
    Dim Cobranca As New VSCobranca
    
    CodPagamento = Nvl(txtDam, 0)
    Data_Vencimento = txtDtVenc
    Cod_Tributo = cboReceita.Coluna(1).Valor
    InscMuni = txtIm
    Screen.MousePointer = 11
    If Not Edita.CriticaCampos(Me) Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    Conta.GeraPagamento txtIm, txtIc, CStr(cboReceita.Coluna(1).Valor), CLng(IIf(IsNumeric(txtPeriodo), txtPeriodo, Right(txtPeriodo, 4) & Left(txtPeriodo, 2))), txtDtVenc, CDbl(CDbl(txtImposto)), CDbl(Nvl(txtMulta, 0)), CDbl(Nvl(txtJuros, 0)), CDbl(txtDam), 0, CInt(Nvl(txtParcela, 0)), CDbl(txtTaxas), , IIf(txtParcela = 0, 1, 3)
    Informa "DAM gerado."
    LimpaCampos Me
    txtDam.SetFocus
    Screen.MousePointer = 0
    
    
    
End Sub

Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    txtDam.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

                                
Private Sub Form_Load()
    Set eDam = New eDam
    Dim Controle As Control
    Dim i As Byte
    
    eDam.PreencherCboTributo cboReceita
    Screen.MousePointer = 0
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set eDam = Nothing
End Sub

Private Sub Grid_DblClick()
    If Grid.ListItems.Count >= 1 Then
        txtDam = Grid.SelectedItem.SubItems(8)
        txtDAM_LostFocus
    End If
End Sub

Private Sub txtDAM_LostFocus()
    Dim CodReceita As String, Im As String, Parcela As String, Ic As String, Periodo As String, DtVenc As String, Imposto As String, Taxas As String, Juros As String, Multa As String
    Dim Sql As String
    Dim rs As VSRecordset
    Dim Tabela As String
    Dim i As Byte
    If Trim(txtDam) = "" Then Exit Sub
    txtParcela = Util.Nvl(txtParcela, "0")
    If eDam.BuscaDam(txtDam, CodReceita, Im, Parcela, Ic, Periodo, DtVenc, Imposto, Taxas, Juros, Multa) Then
        cboReceita.SetarLinha CodReceita, 1
        txtIm = Im
        txtParcela = Nvl("" & Parcela, 0)
        txtIc = Ic
        txtPeriodo = IIf(Len("" & Periodo) = 4, "" & Periodo, Right("" & Periodo, 2) & "/" & Left("" & Periodo, 4))
        txtDtVenc = DtVenc
        txtImposto = Imposto
        txtTaxas = Taxas
        txtJuros = Juros
        txtMulta = Multa
        txtTotalImposto = CDbl(Nvl(txtTaxas, 0)) + CDbl(Nvl(txtImposto, 0)) + CDbl(Nvl(txtJuros, 0)) + CDbl(Nvl(txtMulta, 0))
    Else
        Informa "DAM inexistente. Confirme o número."
        txtDam.SetFocus
    End If
End Sub

Private Sub txtImposto_Change()
    On Error Resume Next
    txtTotalImposto = CDbl(Nvl(txtJuros, 0)) + CDbl(Nvl(txtMulta, 0)) + CDbl(Nvl(txtImposto, 0)) + CDbl(Nvl(txtTaxas, 0))
End Sub

Private Sub txtJuros_Change()
    On Error Resume Next
    txtTotalImposto = CDbl(Nvl(txtJuros, 0)) + CDbl(Nvl(txtMulta, 0)) + CDbl(Nvl(txtImposto, 0)) + CDbl(Nvl(txtTaxas, 0))
End Sub

Private Sub txtMulta_Change()
    On Error Resume Next
    txtTotalImposto = CDbl(Nvl(txtJuros, 0)) + CDbl(Nvl(txtMulta, 0)) + CDbl(Nvl(txtImposto, 0)) + CDbl(Nvl(txtTaxas, 0))
End Sub

Private Sub txtTotalNotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    KeyAscii = AceitaDig(KeyAscii, Valores)
End Sub

Private Sub txtTaxas_Change()
    On Error Resume Next
    txtTotalImposto = CDbl(Nvl(txtJuros, 0)) + CDbl(Nvl(txtMulta, 0)) + CDbl(Nvl(txtImposto, 0)) + CDbl(Nvl(txtTaxas, 0))
End Sub

