VERSION 5.00
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TIMP106 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TIMP106.frx":0000
   ScaleHeight     =   6585
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   30
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TIMP106.frx":0342
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   14
      Top             =   6000
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1032
      Begin VTOcx.cmdVISUAL CmdExcluir 
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Top             =   120
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
         Left            =   7650
         TabIndex        =   11
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   5325
         TabIndex        =   9
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8805
         TabIndex        =   12
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1138
      Icone           =   "TIMP106.frx":2465
   End
   Begin VTOcx.txtVISUAL txtUFM 
      Height          =   480
      Left            =   6975
      TabIndex        =   4
      Tag             =   "Valor"
      Top             =   2850
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   847
      Caption         =   "Valor UFM"
      Text            =   ""
      Formato         =   5
      Restricao       =   3
      AlinhamentoRotulo=   1
      AlinhamentoTexto=   1
   End
   Begin VTOcx.grdVISUAL grdEstimativo 
      Height          =   2355
      Left            =   105
      TabIndex        =   8
      Top             =   3900
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   4154
      CorBorda        =   32768
      Caption         =   "Tabela"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin VTOcx.txtVISUAL txtLimiteInferior 
      Height          =   480
      Left            =   4320
      TabIndex        =   2
      Tag             =   "Limite Inferior"
      Top             =   2850
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   847
      Caption         =   "Limite Inferior"
      Text            =   ""
      Restricao       =   2
      AlinhamentoRotulo=   1
   End
   Begin VTOcx.txtVISUAL txtValor 
      Height          =   480
      Left            =   8460
      TabIndex        =   5
      Tag             =   "Valor"
      Top             =   2850
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   847
      Caption         =   "Valor R$"
      Text            =   ""
      Formato         =   5
      Restricao       =   3
      AlinhamentoRotulo=   1
      AlinhamentoTexto=   1
   End
   Begin VTOcx.txtVISUAL txtLimiteSuperior 
      Height          =   480
      Left            =   5625
      TabIndex        =   3
      Tag             =   "Limite Superior"
      Top             =   2850
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   847
      Caption         =   "Limite Superior"
      Text            =   ""
      Restricao       =   2
      AlinhamentoRotulo=   1
   End
   Begin VTOcx.grdVISUAL grdAtividade 
      Height          =   2310
      Left            =   90
      TabIndex        =   0
      Top             =   705
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   4075
      CorBorda        =   32768
      Caption         =   "Publicidades"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
      CheckBox        =   -1  'True
      MarcaUnico      =   -1  'True
   End
   Begin VTOcx.cmdVISUAL CmdEx 
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Top             =   3375
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   661
      Caption         =   ""
      Acao            =   2
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL CmdAd 
      Height          =   375
      Left            =   8565
      TabIndex        =   6
      Top             =   3375
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cboVISUAL CboTipo 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   3030
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   556
      Caption         =   "Tipo"
      Text            =   ""
      AutoFocaliza    =   0   'False
   End
   Begin VTOcx.cboVISUAL cboSubPublicidade 
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   3420
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   556
      Caption         =   "Sub Publicidade"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Editavel        =   -1  'True
   End
End
Attribute VB_Name = "TIMP106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim AtividadeEstimada As eAtividadeEstimada
'Dim Atividade As Atividade

Private Sub CboTipo_Click()
        cboSubPublicidade.Visible = False
        If grdEstimativo.ListItems.Count <= 0 Then
            Monta_Grid
            grdEstimativo.ListItems.Clear
        End If
    If cboTipo.Coluna(1).Valor = 1 Then
        txtLimiteInferior.Text = "0,00"
        txtLimiteSuperior.Text = "0,00"
        txtLimiteInferior.Enabled = True
        txtLimiteSuperior.Enabled = True
    ElseIf cboTipo.Coluna(1).Valor = 2 Then
         txtLimiteInferior.Enabled = False
        txtLimiteSuperior.Enabled = False
        txtLimiteInferior.Text = "0,00"
        txtLimiteSuperior.Text = "0,00"
    Else
        If grdEstimativo.ListItems.Count <= 0 Then
            grdEstimativo.ColumnHeaders.Clear
            grdEstimativo.ColumnHeaders.Add , , "Item", 1000
            grdEstimativo.ColumnHeaders.Add , , "Código", 1000
            grdEstimativo.ColumnHeaders.Add , , "Nome", 4000
            grdEstimativo.ColumnHeaders.Add , , "Descrição", 4000
            grdEstimativo.ColumnHeaders.Add , , "UFM", 1000
            grdEstimativo.ColumnHeaders.Add , , "REAL", 1000
        End If
        txtLimiteInferior.Text = "0,00"
         txtLimiteInferior.Enabled = False
        txtLimiteSuperior.Enabled = False
        txtLimiteSuperior.Text = "0,00"
        cboSubPublicidade.Visible = True
        
    End If
    cboTipo.Enabled = False
End Sub

Private Sub CmdAd_Click()
    Dim Items As ListItem
    Dim Principal As String
    Dim Ilaco As Integer
    Dim Index As Integer
    
    For Ilaco = 1 To grdAtividade.ListItems.Count
        If grdAtividade.ListItems(Ilaco).Checked Then
            Principal = grdAtividade.ListItems(Ilaco)
            Exit For
        End If
    Next
    If cboTipo.Coluna(1).Valor = 1 Then
        If txtLimiteInferior = "" Or txtLimiteInferior = "0,00" Then
            Util.Avisa "Informe " & txtLimiteInferior.Caption
            txtLimiteInferior.SetFocus
            Exit Sub
        End If
        If txtLimiteSuperior = "" Or txtLimiteSuperior = "0,00" Then
            Util.Avisa "Informe " & txtLimiteSuperior.Caption
            txtLimiteSuperior.SetFocus
            Exit Sub
        End If
        
        If txtUFM = "" Or txtUFM = "0,00" Then
            Util.Avisa "Informe " & txtUFM.Caption
            txtUFM.SetFocus
            Exit Sub
        End If
        
        If txtValor = "" Or txtValor = "0,00" Then
            Util.Avisa "Informe " & txtValor.Caption
            txtValor.SetFocus
            Exit Sub
        End If
    ElseIf cboTipo.Coluna(1).Valor = 2 Then
        If txtUFM = "" Or txtUFM = "0,00" Then
            Util.Avisa "Informe " & txtUFM.Caption
            txtUFM.SetFocus
            Exit Sub
        End If
        
        If txtValor = "" Or txtValor = "0,00" Then
            Util.Avisa "Informe o Valor"
            txtValor.SetFocus
            Exit Sub
        End If
    Else
        If cboSubPublicidade.Text = "" Then
            Util.Avisa "Selecione " & cboSubPublicidade.Caption
            cboSubPublicidade.SetFocus
            Exit Sub
        End If
        If txtUFM = "" Or txtUFM = "0,00" Then
            Util.Avisa "Informe " & txtUFM.Caption
            txtUFM.SetFocus
            Exit Sub
        End If
        
        If txtValor = "" Or txtValor = "0,00" Then
            Util.Avisa "Informe o Valor"
            txtValor.SetFocus
            Exit Sub
        End If
    End If
    Index = grdEstimativo.ListItems.Count + 1
    If cboTipo.Coluna(1).Valor <> 3 Then 'Fixo com item
        grdEstimativo.ListItems.Add Index, , Index
        grdEstimativo.ListItems(Index).SubItems(1) = Principal
        grdEstimativo.ListItems(Index).SubItems(2) = grdAtividade.SelectedItem.SubItems("2")
        grdEstimativo.ListItems(Index).SubItems(3) = txtLimiteInferior
        grdEstimativo.ListItems(Index).SubItems(4) = txtLimiteSuperior
        grdEstimativo.ListItems(Index).SubItems(5) = txtUFM
        grdEstimativo.ListItems(Index).SubItems(6) = txtValor
        grdEstimativo.ListItems(Index).SubItems(7) = cboTipo.Coluna(1).Valor
    Else
        grdEstimativo.ListItems.Add Index, , Index
        grdEstimativo.ListItems(Index).SubItems(1) = Principal
        grdEstimativo.ListItems(Index).SubItems(2) = grdAtividade.SelectedItem.SubItems("2")
        grdEstimativo.ListItems(Index).SubItems(3) = cboSubPublicidade.Text
        grdEstimativo.ListItems(Index).SubItems(4) = txtUFM
        grdEstimativo.ListItems(Index).SubItems(5) = txtValor
    End If
    txtLimiteInferior = "0,00"
    txtLimiteSuperior = "0,00"
    txtUFM = "0,00"
    cboSubPublicidade.ListIndex = -1
    txtValor = "0,00"
    txtLimiteInferior.SetFocus
    
End Sub

Private Sub CmdEx_Click()
    Dim i As Integer
    If grdEstimativo.ListItems.Count >= 1 Then
        grdEstimativo.ListItems.Remove grdEstimativo.SelectedItem.Index
        For i = 1 To grdEstimativo.ListItems.Count
            grdEstimativo.ListItems(i) = i
        Next
    End If
    
End Sub

Private Sub cmdExcluir_Click()
    Dim Principal As String
    Dim Ilaco As Integer
    
    For Ilaco = 1 To grdAtividade.ListItems.Count
        If grdAtividade.ListItems(Ilaco).Checked Then
            Principal = grdAtividade.ListItems(Ilaco)
            Exit For
        End If
    Next
    If Principal <> "" Then
        If Confirma("Deseja excluir?") = True Then
            If cboTipo.Coluna(1).Valor = 1 Or cboTipo.Coluna(1).Valor = 2 Then
                If Bdados.DeletaDados("TAB_PARAMETRO_TAXAS", "TPT_TIP_COD_IMPOSTO = " & Bdados.Converte(Principal, tctexto)) Then
                    Util.Avisa "Dados apagados com sucesso."
                    cmdLimpar_Click
                End If
            Else
                If Bdados.DeletaDados("TAB_PARAMETRO_DETALHE", "TPD_TIP_COD_IMPOSTO = " & Bdados.Converte(Principal, tctexto)) Then
                    Util.Avisa "Dados apagados com sucesso."
                    cmdLimpar_Click
                End If
            End If
        End If
    Else
        Util.Avisa "Selecione um anuncio."
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    grdEstimativo.ListItems.Clear
    grdAtividade.Preencher Bdados, "SELECT tip_cod_imposto as Código,tip_sigla_imposto as Sigla,tip_nome_imposto as Nome FROM TAB_IMPOSTO where tip_sigla_imposto  = '" & Imposto.NomeTributo(ttr_PUBLICIDADE) & "'"
    Me.MousePointer = vbNormal
    cboTipo.Enabled = Not cboTipo.Enabled
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores          As String
    Dim Campos         As String
    Dim Condicao        As String
    Dim Seq               As Integer
    Dim Ilaco              As Integer
    Dim Principal        As String
    Dim B As Integer
    
    For Ilaco = 1 To grdAtividade.ListItems.Count
        If grdAtividade.ListItems(Ilaco).Checked Then
            Principal = grdAtividade.ListItems(Ilaco)
            Exit For
        End If
    Next
    Rem TAB_PARAMETRO_TAXAS
    Rem TAB_PARAMETRO_DETALHE
    
    'If Not Edita.CriticaCampos(Me) Then Exit Sub
    If Principal = "" Then Util.Avisa "Selecione um anuncio.": Exit Sub
    Screen.MousePointer = 11
        If cboTipo.Coluna(1).Valor <> 3 Then
            'Deleto para gravar de novo
            If Bdados.DeletaDados("TAB_PARAMETRO_TAXAS", "TPT_TIP_COD_IMPOSTO = " & Bdados.Converte(Principal, tctexto)) Then
                If grdEstimativo.ListItems.Count >= 1 Then
                    For Ilaco = 1 To grdEstimativo.ListItems.Count
                        Valores = Bdados.PreparaValor(grdEstimativo.ListItems(Ilaco).SubItems(1), grdEstimativo.ListItems(Ilaco), grdEstimativo.ListItems(Ilaco).SubItems(3), grdEstimativo.ListItems(Ilaco).SubItems(4), grdEstimativo.ListItems(Ilaco).SubItems(5), grdEstimativo.ListItems(Ilaco).SubItems(6), cboTipo.Coluna(1).Valor)
                        Campos = "TPT_TIP_COD_IMPOSTO, TPT_SEQUENCIAL,TPT_LIMITE_INFERIOR,TPT_LIMITE_SUPERIOR,TPT_VALOR_UFM,TPT_VALOR_REAL,TPT_TIPO"
                        If Bdados.InsereDados("TAB_PARAMETRO_TAXAS", Valores, Campos) Then
                            B = B + 1
                        End If
                    Next
                End If
             End If
        Else
            If Bdados.DeletaDados("TAB_PARAMETRO_DETALHE", "TPD_TIP_COD_IMPOSTO = " & Bdados.Converte(Principal, tctexto)) Then
                If grdEstimativo.ListItems.Count >= 1 Then
                    For Ilaco = 1 To grdEstimativo.ListItems.Count
                        Valores = Bdados.PreparaValor(grdEstimativo.ListItems(Ilaco).SubItems(1), grdEstimativo.ListItems(Ilaco), grdEstimativo.ListItems(Ilaco).SubItems(3), Bdados.Converte(grdEstimativo.ListItems(Ilaco).SubItems(4), TCMonetario), Bdados.Converte(Trim(grdEstimativo.ListItems(Ilaco).SubItems(5)), TCMonetario), cboTipo.Coluna(1).Valor)
                        Campos = "TPD_TIP_COD_IMPOSTO, TPD_ITEM,TPD_DESCRICAO,TPD_VALOR_UFM,TPD_VALOR_REAL,tpd_tipo"
                        If Bdados.InsereDados("TAB_PARAMETRO_DETALHE", Valores, Campos) Then
                            B = B + 1
                        End If
                    Next
                End If
            End If
        End If
            If B = grdEstimativo.ListItems.Count Then
                Util.Avisa "Dados salvos com sucesso."
                cmdLimpar_Click
            Else
                Util.Avisa "Erro ao gravar tabela"
                Exit Sub
                Resume
            End If

         Screen.MousePointer = 0
    
    
End Sub
Private Sub Pega_Tabelas(Codigo As String)
    Dim Sql As String
    Sql = "select TPT_CODIGO AS Código, tip_sigla_imposto + ' - ' + tip_nome_imposto as Imposto,"
    Sql = Sql & " tpt_limite_inferior As Inferior, tpt_limite_superior As Superior, tpt_valor_ufm As UFM, tpt_valor_real As Valor"
    Sql = Sql & " From TAB_PARAMETRO_TAXAS, tab_imposto where 1 = 1"
    
    If Codigo <> "" Then
        Sql = Sql & " and "
    End If
    
End Sub

Private Sub cmdVISUAL1_Click()
  Dim Sql As String
  Dim i As Integer
  Dim Item
  For i = 1 To grdAtividade.ListItems.Count
    If grdAtividade.ListItems(i).Checked Then
        Item = grdAtividade.ListItems(i)
        Exit For
    End If
  Next
        If grdAtividade.ListItems.Count >= 1 Then
            Sql = "select tpt_sequencial as Item,TPT_TIP_COD_IMPOSTO as Imposto, tip_nome_imposto as Imposto,"
            Sql = Sql & " tpt_limite_inferior As Inferior, tpt_limite_superior As Superior, tpt_valor_ufm As UFM, tpt_valor_real As Valor"
            Sql = Sql & " From TAB_PARAMETRO_TAXAS, tab_imposto"
            Sql = Sql & " where TPT_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
            If grdEstimativo.Preencher(Bdados, Sql) = False Then
                Monta_Grid
            End If
        End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    'Set AtividadeEstimada = New eAtividadeEstimada
   ' Set Atividade = New Atividade
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    grdAtividade.Preencher Bdados, "SELECT tip_cod_imposto as Código,tip_sigla_imposto as Sigla,tip_nome_imposto as Nome FROM TAB_IMPOSTO where tip_sigla_imposto  = '" & Imposto.NomeTributo(ttr_PUBLICIDADE) & "'"
   
   Monta_Grid
   cboTipo.PreencherGeral Bdados, "TIPO TABELA"
   cboSubPublicidade.Visible = False
   cboSubPublicidade.Preencher Bdados, "SELECT TIP_COD_IMPOSTO,TIP_NOME_IMPOSTO FROM TAB_IMPOSTO WHERE TIP_SIGLA_IMPOSTO = '" & Imposto.NomeTributo(ttr_PUBLICIDADE) & "'", 1
            
End Sub
Private Sub Monta_Grid()

   grdEstimativo.ColumnHeaders.Clear
   grdEstimativo.ColumnHeaders.Add , , "Item", 1000
   grdEstimativo.ColumnHeaders.Add , , "Imposto", 1000
   grdEstimativo.ColumnHeaders.Add , , "Nome", 4000
   grdEstimativo.ColumnHeaders.Add , , "L.Inferior", 1000
   grdEstimativo.ColumnHeaders.Add , , "L.Superior", 1000
   grdEstimativo.ColumnHeaders.Add , , "UFM", 1000
   grdEstimativo.ColumnHeaders.Add , , "Valor", 1000
   grdEstimativo.ColumnHeaders.Add , , "Tipo", 1000
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Set AtividadeEstimada = Nothing
    'Set Atividade = Nothing
End Sub

Private Sub fraFUTURO1_mudancaStatus()

End Sub

Private Sub grdAtividade_DblClick()
    'txtCodigo = grdAtividade.SelectedItem
   ' txtCodigo_LostFocus
End Sub

Private Sub grdAtividade_ItemCheck(ByVal Item As MSComctlLib.IListItem)
        Dim Sql As String
        If grdAtividade.ListItems.Count >= 1 Then
            Sql = "select tpt_sequencial as Item,TPT_TIP_COD_IMPOSTO as Imposto, tip_nome_imposto as Nome,"
            Sql = Sql & " tpt_limite_inferior As Inferior, tpt_limite_superior As Superior, tpt_valor_ufm As UFM, tpt_valor_real As Valor,TPT_TIPO as Tipo"
            Sql = Sql & " From TAB_PARAMETRO_TAXAS, tab_imposto"
            Sql = Sql & " where TPT_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
            Sql = Sql & " and TPT_TIP_COD_IMPOSTO = '" & Item.Text & "'"
            If grdEstimativo.Preencher(Bdados, Sql) = False Then
                'Pego da tabela de detalhes...
                Sql = "SELECT tpd_item as Item,"
                Sql = Sql & " TPD_TIP_COD_IMPOSTO AS Código,"
                Sql = Sql & " tip_nome_imposto as Nome,"
                Sql = Sql & " tpd_descricao as Descrição,"
                Sql = Sql & " tpd_valor_ufm as V_UFM,"
                Sql = Sql & " tpd_valor_real As V_Real,TPD_TIPO AS Tipo"
                Sql = Sql & " From TAB_PARAMETRO_DETALHE, tab_imposto"
                Sql = Sql & " where TPD_TIP_COD_IMPOSTO = TIP_COD_IMPOSTO"
                Sql = Sql & " and TPD_TIP_COD_IMPOSTO = " & Bdados.Converte(Item.Text, tctexto)
                If grdEstimativo.Preencher(Bdados, Sql) Then
                    If Trim(grdEstimativo.SelectedItem.SubItems(6)) = 3 Then
                        cboTipo.SetarLinha 3, 1
                        cboTipo.Enabled = False
                    End If
                Else
                    cboTipo.Enabled = True
                    txtLimiteInferior.Enabled = True
                    txtLimiteSuperior.Enabled = True
                End If
                'CboTipo.SetarLinha 3, 1
            Else
                If grdEstimativo.SelectedItem.SubItems(7) = 1 Then
                    cboTipo.SetarLinha 1, 1
                    cboTipo.Enabled = False
                ElseIf grdEstimativo.SelectedItem.SubItems(7) = 2 Then
                    cboTipo.SetarLinha 2, 1
                    cboTipo.Enabled = False
                End If
            End If
        End If
        If grdEstimativo.ListItems.Count >= 1 Then
            CboTipo_Click
        End If
End Sub


Private Sub grdEstimativo_DblClick()
    If grdEstimativo.ListItems.Count >= 1 Then
        Dim i As Integer
        If cboTipo.Coluna(1).Valor = 1 Or cboTipo.Coluna(1).Valor = 2 Then
            txtLimiteInferior = grdEstimativo.SelectedItem.SubItems(3)
            txtLimiteSuperior = grdEstimativo.SelectedItem.SubItems(4)
            txtUFM = grdEstimativo.SelectedItem.SubItems(5)
            txtValor = grdEstimativo.SelectedItem.SubItems(6)
        Else
            cboSubPublicidade = grdEstimativo.SelectedItem.SubItems(3)
            txtUFM = grdEstimativo.SelectedItem.SubItems(4)
            txtValor = grdEstimativo.SelectedItem.SubItems(5)
        End If
        
        grdEstimativo.ListItems.Remove grdEstimativo.SelectedItem.Index
        For i = 1 To grdEstimativo.ListItems.Count
            grdEstimativo.ListItems(i) = i
        Next
    End If
End Sub

Private Sub LimparDados()
    txtLimiteInferior = ""
    txtLimiteSuperior = ""
    txtValor = ""
End Sub

Private Sub txtUFM_LostFocus()
    If txtUFM = "" Then Exit Sub
    txtValor = Calcula_UFM(txtUFM, Converete_Real)
End Sub

Private Sub txtValor_LostFocus()
    If txtValor = "" Then Exit Sub
    txtUFM = Calcula_UFM(txtValor, Converete_UFM)
End Sub
