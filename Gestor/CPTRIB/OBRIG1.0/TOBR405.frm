VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TOBR405 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TOBR405"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   6300
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   2
         Left            =   8565
         TabIndex        =   4
         Top             =   105
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&DAM"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmd 
         Height          =   375
         Index           =   1
         Left            =   9765
         TabIndex        =   3
         Top             =   105
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   1138
      Icone           =   "TOBR405.frx":0000
   End
   Begin VTOcx.grdVISUAL GrdTaxas 
      Height          =   1620
      Left            =   60
      TabIndex        =   5
      Top             =   6345
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   2858
      Caption         =   "Taxas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
      CheckBox        =   -1  'True
   End
   Begin VTOcx.grdVISUAL grdCotas 
      Height          =   4650
      Left            =   45
      TabIndex        =   2
      Top             =   1665
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   8202
      CorTitulo       =   32768
      CorCaption      =   -2147483634
   End
   Begin VTOcx.txtVISUAL txtIm 
      Height          =   300
      Left            =   225
      TabIndex        =   6
      Top             =   795
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   529
      Caption         =   "Inscrição"
      Text            =   ""
      Requerido       =   0   'False
      RetirarMascara  =   0   'False
      AutoTAB         =   -1  'True
   End
   Begin VTOcx.txtVISUAL txtRazao 
      Height          =   300
      Left            =   3195
      TabIndex        =   7
      Top             =   795
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   529
      Caption         =   ""
      Text            =   ""
      Enabled         =   0   'False
      Requerido       =   0   'False
   End
   Begin VTOcx.txtVISUAL txtEndereco 
      Height          =   300
      Left            =   210
      TabIndex        =   8
      Top             =   1125
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   529
      Caption         =   "Endereço"
      Text            =   ""
      Enabled         =   0   'False
      Requerido       =   0   'False
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "Geral"
      Visible         =   0   'False
      Begin VB.Menu mnuReimprime 
         Caption         =   "Reimprimir Documento"
         Index           =   0
      End
      Begin VB.Menu mnuReimprime 
         Caption         =   "Consultar Dados do Pagamento"
         Index           =   1
      End
   End
End
Attribute VB_Name = "TOBR405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim String_Taxas As String
Dim Total_Taxas  As String
Dim NovaData As String

Private Sub cmd_Click(Index As Integer)
    Dim Cobranca As New VSCobranca
 
        Select Case Index
            Case 1
                Unload Me
            Case 2
'                If grdCotas.ListItems.Count <= 0 Then Exit Sub
'                EnderecoImovel = IIf(Trim(txtRazao) = "", txtEndereco, "")
'                EnderecoPessoa = IIf(Trim(EnderecoPessoa) = "", txtEndereco, "")
'                For i = 1 To grdCotas.ListItems.Count
'                    With grdCotas.ListItems
'                        .Item(i).Selected = True
'                        Cobranca.ImprimeDam Rpt, .Item(i), txtIm, txtRazao, "", "", txtIm, "", .Item(i).SubItems(9), .Item(i).SubItems(2), _
'                            .Item(i).SubItems(10), IIf(Len(.Item(i).SubItems(3)) = 4, .Item(i).SubItems(3), Right(.Item(i).SubItems(3), 2) & Left(.Item(i).SubItems(3), 4)), .Item(i).SubItems(5), _
'                            4, .Item(i).SubItems(4), .Item(i).SubItems(5), .Item(i).SubItems(6), 0, .Item(i).SubItems(7), 0, 0, "", "", PicBarra, , , , , , , , , , , tdiImpressora
'                    End With
'                    DoEvents
'                Next
'                Avisa "Impressão concluída."
'
                With grdCotas.SelectedItem
                    NovaData = Imposto.DataVencimentoNova(.SubItems(4))
                    If Trim(NovaData) = "" Then Exit Sub
                End With
                'Pego as taxas
                Call Pega_taxas
                If Me.Caption = "TOBR405 - Cotas de Parcelamento" Then
                    ImprimeSelecionado_Cotas_Parceladas grdCotas, txtRazao, txtEndereco, True, NovaData, tdiTela, String_Taxas, CDbl(Total_Taxas), grdCotas.SelectedItem
                Else
                    ImprimeSelecionado_Cotas_Lancadas grdCotas, txtRazao, txtEndereco, True, NovaData, tdiTela, String_Taxas, CDbl(Total_Taxas), grdCotas.SelectedItem
                End If

        End Select
End Sub

Private Sub Form_Activate()
    Dim Sql As String
    If Me.Caption = "TOBR405 - Cotas de Parcelamento" Then
        Sql = " Select TCp_NUM_COTA AS Documento,tcp_inscricao as Inscrição,"
        Sql = Sql & " TPA_PERIODO AS Ano,TIP_SIGLA_IMPOSTO AS Tributo,TPA_PERIODO AS Periodo,"
        Sql = Sql & " TCp_DATA_VENCIMENTO AS Vencimento,"
        Sql = Sql & " TCp_VALOR_PARCELA As Valor, TCp_VALOR_JUROS As Juros,0 as Multa, "
        Sql = Sql & "Tge_Nome as Situação,0 Taxa,"
        Sql = Sql & " tip_cod_imposto as Imposto,TCp_NUM_PARCELA as Parcela,'' as Obs,'' as Origem,TCP_STATUS_OBRIGACAO_PARCELA"
        Sql = Sql & " From tab_parcelamento, tab_cotas_parcelamento, tab_imposto,vis_status_obrigacao"
        Sql = Sql & " where  tpa_num_parcelamento = TCp_TPA_COD_PARCELAMENTO and tge_codigo = TCP_STATUS_OBRIGACAO_PARCELA and "
        Sql = Sql & " TPA_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
        Sql = Sql & " AND TCP_TPA_COD_PARCELAMENTO = " & Me.Tag & " order by TCP_NUM_PARCELA"
    Else
        Sql = " Select TCO_COD_OBRIGACAo_PARCELA AS Documento,tco_inscricao as Inscrição,"
        Sql = Sql & " tco_periodo as Ano,TIP_SIGLA_IMPOSTO AS Tributo,tco_periodo as Periodo,"
        Sql = Sql & " TCO_DATA_VENCIMENTO AS Vencimento,"
        Sql = Sql & " TCO_VALOR_PARCELA as Valor,TCO_VALOR_JUROS as Juros ,0 as Multa,"
        Sql = Sql & "Tge_Nome as Situação,"
        Sql = Sql & " 0 as Taxa,tip_cod_imposto as Imposto,TCO_NUM_PARCELA as Parcela,'' as Obs,'' as Origem,TCO_STATUS_OBRIGACAO_PARCELA"
        Sql = Sql & " From TAB_COTAS_OBRIGACAO,Tab_imposto,vis_status_obrigacao"
        Sql = Sql & " Where"
        Sql = Sql & " TCO_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
        Sql = Sql & "  and tge_codigo = TCO_STATUS_OBRIGACAO_PARCELA"
        Sql = Sql & " AND TCO_TOC_COD_OBRIGACAO =  '" & Me.Tag & "' order by TCO_NUM_PARCELA"
    End If

    If Not grdCotas.Preencher(Bdados, Sql, 1200, 1500, 0, 1100, 1100, 1100, 1150, 0, 0, 2100, 1000, 0, 1000, 1300, 1500, 0) Then
        Avisa "Nenhum Registro encontrado"
    End If
'    GrdTaxas.Preencher Bdados, "Select * from vis_taxas where ano = '" & Right(Date, 4) & "'"
End Sub
Private Sub Pega_taxas()
    Dim i As Integer
    Dim Pos As Integer
    String_Taxas = ""
    Total_Taxas = 0
    For i = 1 To GrdTaxas.ListItems.Count
        If GrdTaxas.ListItems(i).Checked Then
            Pos = InStr(GrdTaxas.ListItems(i).SubItems(1), "-") - 1
            If String_Taxas = "" Then
                String_Taxas = String_Taxas & " [ " & Left(GrdTaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(GrdTaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            Else
                String_Taxas = String_Taxas & ", [ " & Left(GrdTaxas.ListItems(i).SubItems(1), Pos) & " ]" & " - " & Format(GrdTaxas.ListItems(i).SubItems(2), "###,###,###,##0.00")
            End If
            Total_Taxas = Total_Taxas + CCur(GrdTaxas.ListItems(i).SubItems(2))
        End If
    Next
End Sub

Private Sub grdCotas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not grdCotas.SelectedItem Is Nothing Then
        If Button = 2 Then
            mnuReimprime(0).Caption = "Imprimir DAM da obrigação nº " & grdCotas.SelectedItem.Text
            Me.PopupMenu mnuGeral
        End If
    End If
End Sub

Private Sub mnuReimprime_Click(Index As Integer)
    Dim Cobranca As New VSCobranca
    Select Case Index
        Case 0
            If grdCotas.SelectedItem Is Nothing Then Exit Sub
'            If Not Cobranca.LiberaImpressaoDam(Nvl(grdCotas.SelectedItem.SubItems(1), 0)) Then Exit Sub
            NovaData = Imposto.DataVencimentoNova(grdCotas.SelectedItem.SubItems(5))
            If Trim(NovaData) = "" Then Exit Sub
            
            'Pego as taxas
            Call Pega_taxas
'                ImprimeSelecionado lstObrig, txtRazao, txtEndereco, True, NovaData, tdiTela, String_Taxas, Total_Taxas
            If Trim(txtImovel) = "" Then
                ImprimeSelecionado grdCotas, txtRazao, txtEndereco, False, NovaData, tdiTela, String_Taxas, 0, txtIM, txtEndereco
            Else
                ImprimeSelecionado grdCotas, txtRazao, txtEndereco, False, NovaData, tdiTela, String_Taxas, 0, txtIM, txtEndereco
            End If
        Case 1
            Load TOBR407
            TOBR407.Tag = grdCotas.SelectedItem
            TOBR407.Show 1
    End Select
End Sub
