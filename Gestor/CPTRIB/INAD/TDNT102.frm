VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TDNT102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TDNT102"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   20
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TDNT102.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   19
      Top             =   6885
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   953
      Modulo          =   "Divida Ativa"
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10020
         TabIndex        =   8
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
      TabIndex        =   18
      Top             =   0
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   1138
      Icone           =   "TDNT102.frx":2123
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   2790
      TabIndex        =   17
      Top             =   360
      Width           =   375
   End
   Begin ActiveTabs.SSActiveTabs Tab_Dados 
      Height          =   6225
      Left            =   0
      TabIndex        =   21
      Top             =   660
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   10980
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      Tabs            =   "TDNT102.frx":243D
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5835
         Left            =   -99969
         TabIndex        =   22
         Top             =   30
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   10292
         _Version        =   131082
         TabGuid         =   "TDNT102.frx":24BD
         Begin VB.TextBox txtDescricao 
            Appearance      =   0  'Flat
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
            Height          =   2235
            Left            =   1845
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   930
            Width           =   8790
         End
         Begin VTOcx.cboVISUAL cboImposto 
            Height          =   315
            Left            =   1200
            TabIndex        =   11
            Tag             =   "Tributo"
            Top             =   540
            Width           =   9465
            _ExtentX        =   16695
            _ExtentY        =   556
            Caption         =   "Tributo"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
         End
         Begin VTOcx.cmdVISUAL cmdSalvar 
            Height          =   375
            Left            =   7035
            TabIndex        =   14
            Top             =   5445
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   661
            Caption         =   "&Salvar"
            Acao            =   3
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdCancela 
            Height          =   375
            Left            =   9495
            TabIndex        =   16
            Top             =   5445
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   661
            Caption         =   "&Limpar"
            Acao            =   6
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtExecicio 
            Height          =   285
            Left            =   1470
            TabIndex        =   9
            Top             =   180
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   503
            Caption         =   "Ano"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   4
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.txtVISUAL txtExecicioFim 
            Height          =   285
            Left            =   3330
            TabIndex        =   10
            Top             =   180
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   503
            Caption         =   "Até"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   4
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.grdVISUAL GrdDados 
            Height          =   2175
            Left            =   1845
            TabIndex        =   13
            Top             =   3240
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   3836
         End
         Begin VTOcx.cmdVISUAL CmdExcluirTExto 
            Height          =   375
            Left            =   8265
            TabIndex        =   15
            Top             =   5445
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   661
            Caption         =   "&Excluir"
            Acao            =   2
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   1050
            TabIndex        =   23
            Top             =   945
            Width           =   720
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5835
         Left            =   30
         TabIndex        =   24
         Top             =   30
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   10292
         _Version        =   131082
         TabGuid         =   "TDNT102.frx":24E5
         Begin VTOcx.cboVISUAL cboDoc 
            Height          =   315
            Left            =   240
            TabIndex        =   1
            Tag             =   "Tributo"
            Top             =   510
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   556
            Caption         =   "Documento"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
         End
         Begin VTOcx.cboVISUAL CboParametro 
            Height          =   315
            Left            =   330
            TabIndex        =   3
            Top             =   1260
            Width           =   10620
            _ExtentX        =   18733
            _ExtentY        =   556
            Caption         =   "Parametro"
            Text            =   ""
            AutoFocaliza    =   0   'False
            Requerido       =   0   'False
            Editavel        =   -1  'True
         End
         Begin VTOcx.grdVISUAL grdParametro 
            Height          =   3510
            Left            =   1260
            TabIndex        =   4
            Top             =   1800
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   6191
         End
         Begin VTOcx.txtVISUAL txtCodigo 
            Height          =   285
            Left            =   600
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   150
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   503
            Caption         =   "Código"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   4
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.cmdVISUAL CmdSalvarParametro 
            Height          =   375
            Left            =   7395
            TabIndex        =   5
            Top             =   5400
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   661
            Caption         =   "&Salvar"
            Acao            =   3
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.cmdVISUAL cmdLimparParametro 
            Height          =   375
            Left            =   9825
            TabIndex        =   7
            Top             =   5400
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   661
            Caption         =   "&Limpar"
            Acao            =   6
            CorBorda        =   8421504
            CorFrente       =   16384
         End
         Begin VTOcx.txtVISUAL txtOrdem 
            Height          =   285
            Left            =   630
            TabIndex        =   2
            Top             =   930
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   503
            Caption         =   "Ordem"
            Text            =   ""
            Restricao       =   2
            MaxLen          =   4
            AutoTAB         =   -1  'True
         End
         Begin VTOcx.cmdVISUAL CmdExcluir 
            Height          =   375
            Left            =   8610
            TabIndex        =   6
            Top             =   5400
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   661
            Caption         =   "&Excluir"
            Acao            =   2
            CorBorda        =   8421504
            CorFrente       =   16384
         End
      End
   End
   Begin VB.Menu MnuExcluir 
      Caption         =   "MnuExcluir"
      Visible         =   0   'False
      Begin VB.Menu iteexcluir 
         Caption         =   "Excluir Linha"
         Index           =   0
      End
      Begin VB.Menu iteexcluir 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu iteexcluir 
         Caption         =   "Cancelar"
         Index           =   2
      End
   End
End
Attribute VB_Name = "TDNT102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Imposto As New VSImposto

Private Sub cboDoc_Click()
 Pega_dados
End Sub

Private Sub cmdBuscar_Click()
    Dim Sql As String
    
    Sql = "Select Documento,"
    Sql = Sql & " Ano, Imposto, parametro, Descrição"
    Sql = Sql & " from VIS_PARAMETRO_DAT group by ano,Parametro,Imposto,Descrição"
    GrdDados.Preencher Bdados, Sql
End Sub

Private Sub cmdCancela_Click()
   txtExecicio = ""
   cboImposto.ListIndex = -1
   CboParametro.Text = ""
   txtDescricao = ""
   txtExecicio.SetFocus
   
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub


Private Sub cmdMais_Click()
    On Error Resume Next
    Dim ItmX As Object
    
   If cboDoc.ListIndex <= -1 Then
        Util.Avisa "Selecione o docemento."
        cboDoc.SetFocus
        Exit Sub
   End If
   
   If txtExecicio = "" Then
        Util.Avisa "Informe o ano."
        txtExecicio.SetFocus
        Exit Sub
   End If
   If cboImposto.ListIndex <= -1 Then
       Util.Avisa "Selecione Imposto."
       cboImposto.SetFocus
       Exit Sub
   End If
     
   If txtDescricao = "" Then
        Util.Avisa "Informe a descriçao."
        txtDescricao.SetFocus
        Exit Sub
   End If
   
   If CboParametro.ListIndex <= -1 And CboParametro.Text = "" Then
        Util.Avisa "Informe o parametro."
        CboParametro.SetFocus
        Exit Sub
   End If

   Set ItmX = GrdDados.ListItems.Add(GrdDados.ListItems.Count + 1, , GrdDados.ListItems.Count + 1)
   ItmX.SubItems(1) = cboDoc.Coluna(1).Valor & " - " & cboDoc.Text
   ItmX.SubItems(2) = txtExecicio
   ItmX.SubItems(3) = cboImposto.Coluna(0).Valor & " - " & cboImposto.Text
   ItmX.SubItems(4) = CboParametro.Text
   ItmX.SubItems(5) = txtDescricao
   
   'grdDados.ListItems.Add Index, , Index
   'grdDados.ListItems.Item(Index).SubItems(1) = cboDoc.Coluna(1).Valor & " - " & cboDoc.Text
   'grdDados.ListItems.Item(Index).SubItems(2) = txtExecicio
   'grdDados.ListItems.Item(Index).SubItems(3) = cboImposto.Coluna(0).Valor & " - " & cboImposto.Text
   'grdDados.ListItems.Item(Index).SubItems(4) = CboParametro.Text
   'grdDados.ListItems.Item(Index).SubItems(5) = txtDescricao
   
   txtExecicio = ""
   cboImposto.ListIndex = -1
   CboParametro.Text = ""
   txtDescricao = ""
   txtExecicio.SetFocus
   GrdDados.AtualizarQtd
End Sub

Private Sub CmdMenos_Click()

End Sub

Private Sub cmdExcluir_Click()
    Dim condicao As String
    condicao = "TPD_CODIGO = '" & txtCodigo & "' and TPD_TGL_CODIGO = '" & cboDoc.Coluna(0).Valor & "' and  TPD_ITEM = '" & txtOrdem & "'"
    If Confirma("Deseja excluir o item selecionado?") = True Then
        If Bdados.DeletaDados("TAB_PARAMETRO_DOC", condicao) Then
            'Apago os textos...
            condicao = "TPD_CODIGO = '" & txtCodigo & cboDoc.Coluna(0).Valor & txtOrdem & "'"
            If Bdados.DeletaDados("TAB_PARAMETRO_DOC_NAO_TRIBU", condicao) Then
                Util.Avisa "Operação concluída com sucesso."
                grdParametro.Preencher Bdados, "Select * from VIS_PARAMETRO_NAO_TRIB order by 1,2,3"
                cmdLimparParametro_Click
            End If
        End If
    End If
End Sub

Private Sub CmdExcluirTExto_Click()
    Dim condicao As String
   condicao = "TPD_CODIGO = '" & txtCodigo & cboDoc.Coluna(0).Valor & txtOrdem & "' and TPD_ANO = '" & txtExecicio & "' and TPD_IMPOSTO = '" & cboImposto.Coluna(0).Valor & "'"
    If Confirma("Deseja excluir o item selecionado?") = True Then
        If Bdados.DeletaDados("TAB_PARAMETRO_DAT_NAO_TRIBU", condicao) Then
            Util.Avisa "Operação concluída com sucesso."
            txtExecicio = ""
            txtExecicioFim = ""
            cboImposto.ListIndex = -1
            txtDescricao = ""
            txtExecicio.SetFocus
            GrdDados.Preencher Bdados, "select TPD_CODIGO as Código,tpd_ano as Ano,tpd_imposto as Imposto,tpd_descricao as Descrição   from TAB_PARAMETRO_DAT_NAO_TRIBU WHERE TPD_CODIGO = '" & txtCodigo & cboDoc.Coluna(0).Valor & txtOrdem & "'" & " order by 1,2"
        End If
    End If
End Sub

Private Sub cmdLimparParametro_Click()
    LimpaCampos Me
    txtCodigo.SetFocus
    txtCodigo.Enabled = True
    cboDoc.Enabled = True
    txtOrdem.Enabled = True
    Tab_Dados.Tabs(1).Selected = True
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    On Error GoTo Trata
    Dim Valores As String
    Dim campos As String
    Dim condicao As String
    Dim Contador As Integer
    Dim Salvos As Integer
    Dim Pos_Imposto As Integer
    Dim Pos_Doc As Integer
    Dim Imposto As String
    
    If txtExecicioFim = "" Then
        campos = "TPD_CODIGO,TPD_ANO,TPD_IMPOSTO,TPD_DESCRICAO"
        Valores = Bdados.PreparaValor(txtCodigo & cboDoc.Coluna(0).Valor & txtOrdem, txtExecicio, cboImposto.Coluna(0).Valor, txtDescricao)
        condicao = "TPD_CODIGO = '" & txtCodigo & cboDoc.Coluna(0).Valor & txtOrdem & "' and TPD_ANO = '" & txtExecicio & "' and TPD_IMPOSTO = '" & cboImposto.Coluna(0).Valor & "'"
        
        If Bdados.GravaDados("TAB_PARAMETRO_DAT_NAO_TRIBU", Valores, campos, condicao) Then
            Util.Avisa "Operação concluída com sucesso."
            txtExecicio = ""
            txtExecicioFim = ""
            cboImposto.ListIndex = -1
            txtDescricao = ""
            txtExecicio.SetFocus
            GrdDados.Preencher Bdados, "select TPD_CODIGO as Código,tpd_ano as Ano,tpd_imposto as Imposto,tpd_descricao as Descrição   from TAB_PARAMETRO_DAT_NAO_TRIBU WHERE TPD_CODIGO = '" & txtCodigo & cboDoc.Coluna(0).Valor & txtOrdem & "'" & " order by 1,2"
        End If
    ElseIf txtExecicioFim <> "" Then
        Dim i As Integer
        i = txtExecicioFim - txtExecicio
        
        For Contador = 1 To i
            campos = "TPD_CODIGO,TPD_ANO,TPD_IMPOSTO,TPD_DESCRICAO"
            Valores = Bdados.PreparaValor(txtCodigo & cboDoc.Coluna(0).Valor & txtOrdem, txtExecicio + Contador, cboImposto.Coluna(0).Valor, txtDescricao)
            condicao = "TPD_CODIGO = '" & txtCodigo & cboDoc.Coluna(0).Valor & txtOrdem & "' and TPD_ANO = '" & txtExecicio + Contador & "' and TPD_IMPOSTO = '" & cboImposto.Coluna(0).Valor & "'"
            Bdados.GravaDados "TAB_PARAMETRO_DAT_NAO_TRIBU", Valores, campos, condicao
        Next
            Util.Avisa "Operação concluída com sucesso."
            txtExecicio = ""
            txtExecicioFim = ""
            cboImposto.ListIndex = -1
            txtDescricao = ""
            txtExecicio.SetFocus
            GrdDados.Preencher Bdados, "select TPD_CODIGO as Código,tpd_ano as Ano,tpd_imposto as Imposto,tpd_descricao as Descrição   from TAB_PARAMETRO_DAT_NAO_TRIBU WHERE TPD_CODIGO = '" & txtCodigo & cboDoc.Coluna(0).Valor & txtOrdem & "'" & " order by 1,2"
    End If
    Exit Sub
Trata:
    Util.Avisa Err.Number & Err.Description
    Exit Sub
    Resume
End Sub

Private Sub CmdSalvarParametro_Click()
    Dim Valores As String
    Dim campos As String
    Dim condicao As String
    
    If txtCodigo = "" Then
        Util.Avisa "Informe código."
        txtCodigo.SetFocus
        Exit Sub
    End If
    
    If cboDoc.ListIndex = -1 Or cboDoc.Text = "" Then
        Util.Avisa "Selecione documento."
        cboDoc.SetFocus
        Exit Sub
    End If
    If txtOrdem = "" Then
        Util.Avisa "Informe Ordem."
        txtOrdem.SetFocus
        Exit Sub
    End If
    
    campos = "TPD_CODIGO,TPD_TGL_CODIGO,TPD_ITEM,TPD_PARAMETRO,TPD_CODIGO_MONTADO"
    Valores = Bdados.PreparaValor(txtCodigo, cboDoc.Coluna(0).Valor, txtOrdem, CboParametro.Text, txtCodigo & cboDoc.Coluna(0).Valor & txtOrdem)
    condicao = "TPD_CODIGO = '" & txtCodigo & "' and TPD_TGL_CODIGO = '" & cboDoc.Coluna(0).Valor & "' and  TPD_ITEM = '" & txtOrdem & "'"
    If Bdados.GravaDados("TAB_PARAMETRO_DOC_NAO_TRIBU", Valores, campos, condicao) Then
        Util.Avisa "Operação concluída com sucesso."
        grdParametro.Preencher Bdados, "Select * from VIS_PARAMETRO_NAO_TRIB order by 1,2,3"
        cmdLimparParametro_Click
    End If
End Sub

Private Sub Form_Load()
    Dim Obrig As New Obrigacao
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    cboImposto.Preencher Bdados, "select tip_cod_imposto ,tip_nome_imposto from tab_imposto", 1
    cboDoc.Preencher Bdados, "SELECT TGE_CODIGO,TGE_NOME FROM VIS_DOC_NAO_TRIBUTARIO ORDER BY TGE_CODIGO", 1
    grdParametro.Preencher Bdados, "Select Código,Documento,Item,Parametro,[Código Montado]  from VIS_PARAMETRO_NAO_TRIB order by 1,2,3", 0, 1000, 800, 7000, 0
    CboParametro.Preencher Bdados, "SELECT  TPD_PARAMETRO FROM TAB_PARAMETRO_DOC_NAO_TRIBU"
    'grdDados.ColumnHeaders.Clear
    'With grdDados.ColumnHeaders
    '    .Add , , "Item", 0
     '   .Add , , "Documento", 0
     '   .Add , , "Ano"
     '''   .Add , , "Imposto"
    '    .Add , , "parametro"
     '   .Add , , "Descrição"
   ' End With
End Sub


Private Sub txtExercicio_KeyPress(KeyAscii As Integer)
    KeyAscii = AceitaDig(KeyAscii, Numero)
End Sub


Private Sub txtInicio_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtValidade_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtic_LostFocus()
'    CarregaEnderecoImovel txtIc, txtEndereco
End Sub

Private Sub grdDados_ItemClick(ByVal Item As MSComctlLib.IListItem)
    Dim Pos As Integer
    
    If GrdDados.ListItems.Count >= 1 Then
       txtExecicio = GrdDados.SelectedItem.SubItems(1)
       cboImposto.SetarLinha GrdDados.SelectedItem.SubItems(2)
       txtDescricao = GrdDados.SelectedItem.SubItems(3)
    End If
'A Metodologia de Apuração da Base de Cálculo está correta
'PARECER: Há Liquidez e Certeza na Base de Cálculo do crédito tributário.
'CONCLUSÃO: O crédito tributário foi APROVADO na Sub-Apuração da sua Base de Cálculo.
End Sub
Private Function Index(Combo As Object, Texto As String)
    Dim i As Integer
    For i = 0 To Combo.ListCount
        If cboDoc.Coluna(0).Valor = Texto Then
            cboDoc.ListIndex = Index
        End If
    Next
End Function
Private Sub grdDados_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu MnuExcluir
    End If
    
End Sub
Private Sub Pega_dados()
    Dim Sql As String
    Dim Rs As VSRecordset
    If txtCodigo = "" Then Exit Sub
    If cboDoc.Text = "" Or cboDoc.ListIndex = -1 Then Exit Sub
    If txtOrdem = "" Then Exit Sub
    
    Sql = "Select tpd_parametro from TAB_PARAMETRO_DOC_NAO_TRIBU where TPD_CODIGO = '" & txtCodigo & "' and TPD_CODIGO = '" & cboDoc.Coluna(0).Valor & "' and  TPD_ITEM = '" & txtOrdem & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        CboParametro.Text = "" & Rs.Fields("tpd_parametro")
        txtCodigo.Enabled = False
        cboDoc.Enabled = False
        txtOrdem.Enabled = False
    Else
        txtCodigo.Enabled = True
        cboDoc.Enabled = True
        txtOrdem.Enabled = True
        CboParametro.Text = ""
    End If
End Sub

Private Sub grdParametro_DblClick()
    Tab_Dados.Tabs(2).Selected = True
    GrdDados.Preencher Bdados, "select TPD_CODIGO as Código,tpd_ano as Ano,tpd_imposto as Imposto,tpd_descricao as Descrição   from TAB_PARAMETRO_DAT_NAO_TRIBU WHERE TPD_CODIGO = '" & txtCodigo & cboDoc.Coluna(0).Valor & txtOrdem & "'" & " order by 1,2"
End Sub

Private Sub grdParametro_ItemClick(ByVal Item As MSComctlLib.IListItem)
 If grdParametro.ListItems.Count >= 1 Then
        txtCodigo = grdParametro.SelectedItem
        cboDoc.SetarLinha grdParametro.SelectedItem.SubItems(1)
        txtOrdem = grdParametro.SelectedItem.SubItems(2)
        CboParametro.Text = grdParametro.SelectedItem.SubItems(3)
    End If
End Sub



Private Sub txtCodigo_LostFocus()
    Pega_dados
End Sub

'Private Sub txtExecicio_LostFocus()
'    If txtExecicio = "" Then Exit Sub
'    If Len(txtExecicio) <> 4 Then
'        Util.Avisa "Informe o ano com 4 caracteres."
'        txtExecicio.SetFocus
'    End If
'End Sub

Private Sub txtOrdem_LostFocus()
Pega_dados
End Sub
