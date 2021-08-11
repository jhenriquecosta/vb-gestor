VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCER401 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCER401"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   18
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCER401.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.TextBox txtIMProprietario 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   7755
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "Incrição Municipal"
      Top             =   1380
      Width           =   1740
   End
   Begin VTOcx.txtVISUAL txtIC 
      Height          =   300
      Left            =   945
      TabIndex        =   0
      Tag             =   "Inscrição Cadastral"
      Top             =   720
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   529
      Caption         =   "IC"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   20
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   15
      Top             =   6165
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   900
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   5940
         TabIndex        =   8
         Top             =   90
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   661
         Caption         =   "R&eimpressão"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   3885
         TabIndex        =   6
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8490
         TabIndex        =   10
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   7455
         TabIndex        =   9
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Height          =   645
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   1138
      Icone           =   "TCER401.frx":2123
   End
   Begin VTOcx.grdVISUAL grdCertidoes 
      Height          =   4020
      Left            =   90
      TabIndex        =   11
      Top             =   2115
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   4339
      Caption         =   "Certidões Emitidas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VTOcx.txtVISUAL txtDataRequerimento 
      Height          =   300
      Left            =   3555
      TabIndex        =   3
      Tag             =   "Data do Requerimento"
      Top             =   1740
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   529
      Caption         =   "Data"
      Text            =   ""
      Formato         =   0
      Restricao       =   2
      MaxLen          =   10
   End
   Begin VTOcx.txtVISUAL txtProprietário 
      Height          =   300
      Left            =   150
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1380
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   529
      Caption         =   "Proprietário"
      Text            =   ""
      Enabled         =   0   'False
      MaxLen          =   50
   End
   Begin VTOcx.txtVISUAL txtEndereco 
      Height          =   300
      Left            =   180
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1050
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   529
      Caption         =   "Localização"
      Text            =   ""
      Enabled         =   0   'False
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdBuscar 
      Height          =   315
      Left            =   3420
      TabIndex        =   1
      Top             =   705
      Visible         =   0   'False
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   556
      Caption         =   "Buscar Imóvel"
      Acao            =   1
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.txtVISUAL txtProtocolo 
      Height          =   300
      Left            =   6735
      TabIndex        =   5
      Top             =   1740
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   529
      Caption         =   "Protocolo"
      Text            =   ""
      MaxLen          =   10
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtNumRequerimento 
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   1740
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      Caption         =   "Requerimento"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   5
   End
   Begin VTOcx.txtVISUAL txtAno 
      Height          =   300
      Left            =   5520
      TabIndex        =   4
      Top             =   1740
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   529
      Caption         =   "Ano"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   4
      MinLen          =   4
      RetirarMascara  =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IM"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   7500
      TabIndex        =   17
      Top             =   1425
      Width           =   180
   End
   Begin VB.Menu mnuNotifica 
      Caption         =   "."
      Visible         =   0   'False
      Begin VB.Menu mnuEmitir 
         Caption         =   "&Emitir notificação ..."
      End
   End
End
Attribute VB_Name = "TCER401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub cmdBuscar_Click()
Dim sql As String
    If txtIC <> "" Then
        sql = "SELECT TDI_TCO_COD_COMPONENTE FROM TAB_DETALHE_IMOVEL WHERE TDI_TIM_IC = '" & txtIC & "' AND TDI_TGC_COD_GRUPO = 3 AND TDI_TCO_COD_COMPONENTE = 2"
        If Bdados.AbreTabela(sql) Then
            sql = "SELECT * FROM VIS_ENDERECO_IMOVEL WHERE TIM_IC = '" & txtIC & "'"
            If Bdados.AbreTabela(sql) Then
                txtEndereco = Bdados.Tabela!TTL_NOME & " " & Bdados.Tabela!tlg_nome & ", " & Bdados.Tabela!tim_numero & " " & Bdados.Tabela!tim_complemento & " - " & Bdados.Tabela!TBA_NOME
            End If
            sql = "SELECT TCI_NOME, TCI_IM FROM TAB_CONTRIBUINTE, TAB_IMOVEL WHERE TIM_TCI_IM = TCI_IM AND TIM_IC = '" & txtIC & "'"
            If Bdados.AbreTabela(sql) Then
                txtProprietário = Bdados.Tabela(0)
                txtIMProprietario = Bdados.Tabela(1)
                If txtIMProprietario = "" Then
                    Avisa "Não foi possível encontrar a Inscrição Municipal referente ao imóvel '" & txtIC & "'. Verifique o cadastro do imóvel e tente novamente."
                    Edita.LimpaCampos Me
                    txtIC.SetFocus
                Else
                    txtNumRequerimento.SetFocus
                End If
            End If
            Bdados.FechaTabela
        Else
            Util.Informa "Imóvel indicado não é isento de imposto. Não será possível emitir a certidão para o mesmo."
            Edita.LimpaCampos Me
            txtIC.SetFocus
        End If
    End If
End Sub

Private Sub cmdCancela_Click()
    Edita.LimpaCampos Me
    MostraIsencoes
    txtIC.SetFocus
End Sub

Private Sub cmdExcluir_Click()
    If grdCertidoes.SelectedItem Is Nothing Then Exit Sub
    If Util.Confirma("Deseja excluir " & Trim(grdCertidoes.SelectedItem) & "/" & grdCertidoes.SelectedItem.SubItems(1) & "?") Then
        If Bdados.DeletaDados("TAB_ISENCAO_IPTU", "TII_NUM_REQUERIMENTO = " & grdCertidoes.SelectedItem & " AND TII_ANO = " & grdCertidoes.SelectedItem.SubItems(1)) Then
            Util.Informa "Dados excluídos com sucesso."
            Edita.LimpaCampos Me
            txtIC.SetFocus
            MostraIsencoes
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo Trata
    Dim rs As VSRecordset
    Dim sql As String
    Dim NUM As String
    If grdCertidoes.SelectedItem Is Nothing Then Exit Sub
    NUM = Trim(grdCertidoes.SelectedItem.Text)
    
    Screen.MousePointer = 11
 
    Call ImprimeIsencao(Format(NUM, "0000"))

    Screen.MousePointer = 0
    Exit Sub
    
Trata:
    Screen.MousePointer = 0
    Erro Err.Description
    Bdados.FechaTabela rs
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Ordem As String
    Dim sql As String
    Dim rs As VSRecordset
    If Not Edita.CriticaCampos(Me) Then
        Informa "Digite a Inscrição Cadastral do imóvel e clique em Buscar Imóvel."
        Exit Sub
    End If
    If txtNumRequerimento = "" Then
        sql = "SELECT MAX(TII_NUM_REQUERIMENTO)  AS Ordem FROM TAB_ISENCAO_IPTU"
        If Bdados.AbreTabela(sql) Then
            Ordem = Format(CDbl(Nvl("" & Bdados.Tabela(0), 0)) + 1, "0000")
        End If
    Else
        Ordem = Format(txtNumRequerimento, "0000")
    End If
    
    'criticanto o num de edificacoes. SÓ PODE TER 1 PARA IMPRIMIR
    sql = "SELECT EDIFICACOES FROM VIS_NUM_EDIFICACOES WHERE IC='" & txtIC & "'"
    If Bdados.AbreTabela(sql, rs) Then
        If rs(0) & "" = 1 Then
            Bdados.FechaTabela rs
            
            'criticando a destinacao. SÓ PODE SER RESIDENCIAL PARA IMPRIMIR
            ' era pra ser o tdi_componente mas foi o tdi_valor_item devido erros no BD....
            
            sql = "select tdi_valor_item  from TAB_DETALHE_IMOVEL where TDI_TGC_COD_GRUPO = 11 AND TDI_TIM_IC='" & txtIC & "'"
            If Bdados.AbreTabela(sql, rs) Then
                If rs(0) & "" = 1 Then
                    Grava Ordem, txtAno
                Else
                    Avisa "A destinação deste imóvel não é residencial."
                End If
            Else
                Avisa "Não foi possível verificar a destinação."
            End If
        Else
            Avisa "Este imóvel possui mais de uma edificação."
        End If
    Else
        Avisa "Não foi possível verificar o número de edificações."
    End If
    Bdados.FechaTabela rs
    
    
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    Call MostraIsencoes
End Sub

Private Sub MostraIsencoes()
    grdCertidoes.Preencher Bdados, "select TII_NUM_REQUERIMENTO as Requerimento, TII_ANO AS Ano, TII_NUM_PROTOCOLO as Protocolo, TII_TIM_IC as IC , TII_TCI_IM as IM, TII_DATA_REQUERIMENTO as Data , TII_DATA_EMISSAO as Emissão from TAB_ISENCAO_IPTU"
End Sub

Private Sub Grava(Ordem As String, Ano As String)
    Dim sValores As String
    Dim sCampos As String
    Dim Obrig As New Obrigacao
    Dim sql As String
    Dim rs As VSRecordset
    sValores = Bdados.PreparaValor(Trim(Bdados.Converte(Ordem, VSClass.tctexto)), txtProtocolo, Bdados.Converte(txtIC, VSClass.tctexto), Bdados.Converte(txtIMProprietario, VSClass.tctexto), txtDataRequerimento, Format(Date, "dd/mm/yyyy"), CInt(txtAno))
    sCampos = "TII_NUM_REQUERIMENTO,TII_NUM_PROTOCOLO,TII_TIM_IC,TII_TCI_IM,TII_DATA_REQUERIMENTO,TII_DATA_EMISSAO, TII_ANO"
'    If Bdados.AbreTabela("SELECT * FROM TAB_ISENCAO_IPTU WHERE TII_NUM_REQUERIMENTO=" & O & " and TII_ANO = " & Ano) Then
'        If Util.Confirma("Número " & O & " já cadastrado no ano de " & Ano & ". Deseja alterar?") Then
'            sValores = Bdados.PreparaValor(Trim(Bdados.Converte(O, vsclass.TCTexto)), txtProtocolo, Bdados.Converte(txtIC, vsclass.TCTexto), Bdados.Converte(txtIMProprietario, vsclass.TCTexto), txtDataRequerimento, CInt(txtAno))
'            sCampos = "TII_NUM_REQUERIMENTO,TII_NUM_PROTOCOLO,TII_TIM_IC,TII_TCI_IM,TII_DATA_REQUERIMENTO"
'            If Bdados.GravaDados("TAB_ISENCAO_IPTU", sValores, sCampos, "TII_NUM_REQUERIMENTO = " & txtNumRequerimento & " AND TII_ANO = " & txtAno) Then
'                Call MostraIsencoes
'                Util.Informa "Certidão " & O & " gravada com sucesso."
'            Else
'                Util.Erro "Houve erro ao gravar as informações da certidão."
'            End If
'        Else
'            Exit Sub
'        End If
'    Elsendif
    sql = "SELECT TOC_COD_OBRIGACAO FROM TAB_OBRIGACAO_CONTRIBUINTE WHERE TOC_INSCRICAO ='" & txtIC & _
        "' AND TOC_PERIODO =" & txtAno & " AND TOC_TIP_COD_IMPOSTO ='" & Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_IPTU)) & "'"
    If Bdados.AbreTabela(sql, rs) Then
        Call Obrig.TrocaSitObrigacao(rs!TOC_COD_OBRIGACAO, etsCreditoIsento)
    End If
    If Util.Confirma("Deseja gravar " & Ordem & "/" & Ano & "?") Then
        If Bdados.GravaDados("TAB_ISENCAO_IPTU", sValores, sCampos, "TII_NUM_REQUERIMENTO = " & Ordem & " AND TII_ANO = " & Ano) Then
             Call MostraIsencoes
              If Util.Confirma("Imprimir a certidão?") Then
                 Imprime Ordem
             Else
                 Util.Informa "Certidão " & Ordem & " gravada com sucesso."
             End If
             txtIC.SetFocus
             Edita.LimpaCampos Me
         Else
             Util.Erro "Houve erro ao gravar as informações da certidão."
         End If
    End If
    Bdados.FechaTabela
    Screen.MousePointer = 0
End Sub

Private Sub Imprime(O As String)
    On Error GoTo Trata
    Dim i As Long
    Dim Ok As Boolean
    Ok = False
    For i = grdCertidoes.ListItems.Count To 1 Step -1
        If Trim(grdCertidoes.ListItems.Item(i).Text) = Trim(O) Then
            grdCertidoes.ListItems.Item(i).Selected = True
            Ok = True
            Exit For
        End If
    Next
    If Ok Then
        cmdImprimir_Click
    Else
        Avisa "Não foi possível encontrar o item '" & O & "' na grade abaixo."
    End If
    
    Exit Sub
Trata:
    Erro Err.Description
End Sub

Private Sub ImprimeIsencao(Ordem As String)
    On Error GoTo Trata
    Dim Rpt As New VSRelatorio
    Set Rpt = Nothing
    Set Rpt = New VSRelatorio
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path & "\TCertidaoIsento.rpt") Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Selecao = "{TAB_ISENCAO_IPTU.TII_NUM_REQUERIMENTO} = '" & Ordem & "'"
        .Titulo = "Certidão de Isenção de IPTU"
        .Arvore = False
        .Visualizar
    End With
    Set Rpt = Nothing
    
    Exit Sub
Trata:
    Erro Err.Description
End Sub

Private Sub grdCertidoes_DblClick()
    If grdCertidoes.SelectedItem Is Nothing Then Exit Sub
    txtIC = grdCertidoes.SelectedItem.SubItems(3)
    Call txtic_LostFocus
    txtProtocolo = grdCertidoes.SelectedItem.SubItems(2)
    txtDataRequerimento = grdCertidoes.SelectedItem.SubItems(5)
    txtNumRequerimento = grdCertidoes.SelectedItem
    txtAno = grdCertidoes.SelectedItem.SubItems(1)
    
End Sub

Private Sub txtDataRequerimento_LostFocus()
    If Trim$(txtDataRequerimento) <> "" Then txtAno = Year(txtDataRequerimento)
End Sub

Private Sub txtic_LostFocus()
    Dim RetNome As String
    Dim Doc As String
    
    If Trim(txtIC) = "" Then Exit Sub
    txtIC = BuscaContribuinte(txtIC, txtProprietário, txtEndereco, , etiImovel)
    cmdBuscar_Click
End Sub

