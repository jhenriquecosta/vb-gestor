VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TMPU701 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TMPU701"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   ControlBox      =   0   'False
   Icon            =   "TMPU701.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   24
      Top             =   6225
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   7440
         TabIndex        =   22
         ToolTipText     =   "Sair"
         Top             =   105
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   330
         Left            =   4650
         TabIndex        =   19
         ToolTipText     =   "Salvar"
         Top             =   105
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   345
         Left            =   2610
         TabIndex        =   17
         ToolTipText     =   "Imprimir"
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   330
         Left            =   5580
         TabIndex        =   20
         ToolTipText     =   "Novo"
         Top             =   105
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         Caption         =   "&Novo"
         Acao            =   1
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   330
         Left            =   6510
         TabIndex        =   21
         ToolTipText     =   "Excluir"
         Top             =   105
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   3720
         TabIndex        =   18
         ToolTipText     =   "Buscar"
         Top             =   105
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.grdVISUAL grdLogradouro 
      Height          =   3075
      Left            =   120
      TabIndex        =   23
      Top             =   3060
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5424
      Caption         =   "Logradouros"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VTOcx.cboVISUAL cboRelatorio 
      Height          =   510
      Left            =   5025
      TabIndex        =   16
      Top             =   2475
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   900
      Caption         =   "Relatório"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Alinhamento     =   1
   End
   Begin ActiveTabs.SSActiveTabs tabRep 
      Height          =   1740
      Left            =   105
      TabIndex        =   25
      Tag             =   "Documento gerencial"
      Top             =   720
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3069
      _Version        =   131082
      TabCount        =   3
      TabOrientation  =   2
      BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tabs            =   "TMPU701.frx":08CA
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel13 
         Height          =   1350
         Left            =   -99969
         TabIndex        =   26
         Top             =   30
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   2381
         _Version        =   131082
         TabGuid         =   "TMPU701.frx":0976
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            Caption         =   "Fim"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   105
            TabIndex        =   34
            Top             =   0
            Width           =   7875
            Begin VTOcx.txtVISUAL txtCodLogrFinal 
               Height          =   315
               Left            =   135
               TabIndex        =   11
               Top             =   330
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   556
               Caption         =   "Logradouro"
               Text            =   ""
               Restricao       =   2
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCodBairroFinal 
               Height          =   315
               Left            =   570
               TabIndex        =   13
               Top             =   675
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   556
               Caption         =   "Bairro"
               Text            =   ""
               Restricao       =   2
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtLogrFinal 
               Height          =   315
               Left            =   2520
               TabIndex        =   12
               Top             =   330
               Width           =   5145
               _ExtentX        =   9075
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtBairroFinal 
               Height          =   315
               Left            =   2520
               TabIndex        =   14
               Top             =   675
               Width           =   5145
               _ExtentX        =   9075
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               RetirarMascara  =   0   'False
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel7 
         Height          =   1350
         Left            =   30
         TabIndex        =   27
         Top             =   30
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   2381
         _Version        =   131082
         TabGuid         =   "TMPU701.frx":099E
         Begin VTOcx.txtVISUAL txtCodigo 
            Height          =   315
            Left            =   405
            TabIndex        =   0
            Tag             =   "Codigo"
            Top             =   90
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            Caption         =   "Codigo"
            Text            =   ""
            Restricao       =   2
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtDescricao 
            Height          =   315
            Left            =   2325
            TabIndex        =   3
            Tag             =   "Descricao"
            Top             =   465
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.cboVISUAL cboTipo 
            Height          =   315
            Left            =   45
            TabIndex        =   2
            Tag             =   "Tipo"
            Top             =   465
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            Caption         =   "Logradouro"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtBairro 
            Height          =   315
            Left            =   495
            TabIndex        =   5
            Tag             =   "Bairro"
            Top             =   855
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            Caption         =   "Bairro"
            Text            =   ""
            Restricao       =   2
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtSetor 
            Height          =   315
            Left            =   6990
            TabIndex        =   1
            Top             =   90
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            Caption         =   "Setor"
            Text            =   ""
            Restricao       =   2
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.cboVISUAL cboBairro 
            Height          =   315
            Left            =   2310
            TabIndex        =   6
            Top             =   870
            Width           =   5760
            _ExtentX        =   10160
            _ExtentY        =   556
            Caption         =   ""
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtCep 
            Height          =   315
            Left            =   6450
            TabIndex        =   4
            Tag             =   "CEP"
            Top             =   465
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            Caption         =   "CEP"
            Text            =   ""
            Formato         =   4
            Restricao       =   2
            MaxLen          =   10
            RetirarMascara  =   0   'False
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel5 
         Height          =   915
         Left            =   -99969
         TabIndex        =   28
         Top             =   30
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   1614
         _Version        =   131082
         TabGuid         =   "TMPU701.frx":09C6
         Begin VTOcx.fraVISUAL fraVISUAL5 
            Height          =   855
            Left            =   0
            TabIndex        =   29
            Top             =   15
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   1508
            Altura          =   1905
            Caption         =   " Dados Responsável"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel6 
         Height          =   3270
         Left            =   -99969
         TabIndex        =   30
         Top             =   30
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   5768
         _Version        =   131082
         TabGuid         =   "TMPU701.frx":09EE
         Begin VTOcx.fraVISUAL fraVISUAL6 
            Height          =   3285
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   5794
            Altura          =   1905
            Caption         =   " Livro Fiscal (Modelos Diferentes)"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483633
            Ocultavel       =   0   'False
            Borda           =   0
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel9 
         Height          =   1350
         Left            =   -99969
         TabIndex        =   32
         Top             =   30
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   2381
         _Version        =   131082
         TabGuid         =   "TMPU701.frx":0A16
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            Caption         =   "Início"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   120
            TabIndex        =   33
            Top             =   -15
            Width           =   7830
            Begin VTOcx.txtVISUAL txtCodLogrInicial 
               Height          =   315
               Left            =   135
               TabIndex        =   7
               Top             =   330
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   556
               Caption         =   "Logradouro"
               Text            =   ""
               Restricao       =   2
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtCodBairroInicial 
               Height          =   315
               Left            =   585
               TabIndex        =   9
               Top             =   675
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   556
               Caption         =   "Bairro"
               Text            =   ""
               Restricao       =   2
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtLogrInicial 
               Height          =   315
               Left            =   2505
               TabIndex        =   8
               Top             =   330
               Width           =   5190
               _ExtentX        =   9155
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               RetirarMascara  =   0   'False
            End
            Begin VTOcx.txtVISUAL txtBairroInicial 
               Height          =   315
               Left            =   2505
               TabIndex        =   10
               Top             =   675
               Width           =   5190
               _ExtentX        =   9155
               _ExtentY        =   556
               Caption         =   ""
               Text            =   ""
               Enabled         =   0   'False
               RetirarMascara  =   0   'False
            End
         End
      End
   End
   Begin VTOcx.cboVISUAL cboCaracterizacao 
      Height          =   510
      Left            =   150
      TabIndex        =   15
      Top             =   2475
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   900
      Caption         =   "Caracterização"
      Text            =   ""
      AutoFocaliza    =   0   'False
      Alinhamento     =   1
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   1138
      Icone           =   "TMPU701.frx":0A3E
   End
End
Attribute VB_Name = "TMPU701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim CodLoteamento As String, CodDivisa As String
Dim CodAntigo As String

Private Sub cboBairro_Click()
    If cboBairro.ListCount > 0 Then
        txtBairro = cboBairro.Coluna(1).Valor
    End If
End Sub

Private Sub cmdBuscar_Click()
    consultarLogradouros txtCodigo, txtSetor, txtBairro, cboCaracterizacao, txtDescricao, cboTipo
End Sub

Private Sub cmdExcluir_Click()
    If CodAntigo <> "" Then
        If Util.Confirma("Excluir logradouro " & CodAntigo & " ?") Then
            If Bdados.AbreTabela("SELECT * FROM TAB_IMOVEL WHERE tim_tlg_cod_logradouro='" & CodAntigo & "'") Then
                Erro "O logradouro possui imóveis. Impossível excluir."
            Else
                Bdados.DeletaDados "TAB_TRECHO", "TTC_TLG_COD_LOGRADOURO='" & CodAntigo & "'"
                Bdados.DeletaDados "TAB_DETALHE_LOGRADOURO", "tdl_tlg_cod_logradouro='" & CodAntigo & "'"
                Bdados.DeletaDados "TAB_VALOR_TERRENO", "tvl_tlg_cod_logradouro='" & CodAntigo & "'"
                Bdados.DeletaDados "TAB_LOGRADOURO", "tlg_cod_logradouro='" & CodAntigo & "'"
                Avisa "Logradouro apagado com sucesso."
                cmdNovo_Click
                grdLogradouro.ListItems.Clear
            End If
        End If
    End If
End Sub

Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    CodAntigo = ""
    tabRep.Tabs(1).Selected = True
    txtCodigo.SetFocus
    'txtCodLogr.SetFocus
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo trata
    Dim Filtro As String
    Dim Titulo As String
    
    Screen.MousePointer = vbArrowHourglass
    Set Rpt = New VSRelatorio
    
    '1. Arquivo
    Select Case cboRelatorio.ListIndex
        Case 0 'Cadastro
            If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TCadLog.rpt") Then Exit Sub
            Titulo = "Cadastro "
        Case 1 'Relacao
            If Not Rpt.DefinirArquivo(Bdados, App.Path & "\TRelLog.rpt") Then Exit Sub
            Titulo = "Relação "
        Case Else
            Screen.MousePointer = vbNormal
            Erro "Relatório não definido."
            cboRelatorio.SetFocus
            Set Rpt = Nothing
            Exit Sub
    End Select
    
    Filtro = ""
    
    '2. Selecao
    If Trim$(txtSetor) <> "" Then
        'Filtro = "{TAB_TRECHO.TTC_SETOR} = '" & txtSetor & "'"
    End If
    
    Select Case cboCaracterizacao.ListIndex
        Case 0
            Filtro = Filtro & IIf(Filtro = "", "", " and ") & "mid({VIS_BVT.tlg_cod_logradouro},1," & Len(CodLoteamento) & ") < '" & CodLoteamento & "'"
            Rpt.Formulas "VTTituloRel", Titulo & " de  Logradouros"
        Case 1
            Filtro = Filtro & IIf(Filtro = "", "", " and ") & "Tonumber(mid({VIS_BVT.tlg_cod_logradouro},1," & Len(CodLoteamento) & ")) >= " & CodLoteamento & " and Tonumber(mid({VIS_BVT.tlg_cod_logradouro},1," & Len(CodLoteamento) & ")) <= " & CodDivisa
            Rpt.Formulas "VTTituloRel", Titulo & " de  Loteamentos"
        Case 2
            Filtro = Filtro & IIf(Filtro = "", "", " and ") & "mid({VIS_BVT.tlg_cod_logradouro},1," & Len(CodDivisa) & ") = '" & CodDivisa & "'"
            Rpt.Formulas "VTTituloRel", Titulo & " de  Divisas"
    End Select
    
    
    Rpt.Selecao = Filtro
    
    '3. Formula
    If Trim$(txtSetor) <> "" Then
        Rpt.Formulas "Setor", "Setor " & txtSetor
    End If


    '4. Cabecalho/Rodape
    Rpt.Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
    Select Case cboRelatorio.ListIndex
        Case 0 'Cadastro
            Rpt.Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), "TCadLog", Aplicacoes.Usuario, Horizontal
        
        Case 1 'Relacao
            Rpt.Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), "TRelLog", Aplicacoes.Usuario
    
    End Select
    
    
    
    
    Rpt.Visualizar
    
    Set Rpt = Nothing
'    Screen.MousePointer = 11
'    Rpt.DefinirArquivo Bdados, App.Path + "\TLogradouros.rpt"
'    ''rpt.Connect = Bdados.BDSistema.Connect
'    With Rpt
'        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
'        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
'        .Arvore = False
'        .Visualizar
'    End With
'     Screen.MousePointer = 0
    Screen.MousePointer = vbNormal
    
    Exit Sub
trata:
    Screen.MousePointer = vbNormal
    Erro Err.Description
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Campos As String
    Dim Valores As String
    Dim CodLogra As String
    Dim rl As VSRecordset
    
    'BCP
    If Len(txtCodigo) = 0 Then
        Dim bcpCodLogra As String
        bcpCodLogra = Temp.PegaParametro(Bdados, "MUNICIPIO")
        If Bdados.AbreTabela("SELECT COUNT(tlg_cod_logradouro) AS TOTAL FROM TAB_LOGRADOURO WHERE tlg_tba_cod_bairro =  " & txtBairro, rl) Then
            bcpCodLogra = bcpCodLogra & txtBairro & rl(0)
        Else
            bcpCodLogra = bcpCodLogra & txtBairro
        End If
        txtCodigo = bcpCodLogra
    End If
    'FIM BCP
    CodLogra = txtCodigo
    If Edita.CriticaCampos(Me) Then
        If CodAntigo <> "" Then
            If CodAntigo <> txtCodigo Then
                atualizarLogradouro CodAntigo, txtCodigo
            End If
        End If
        
        Campos = "tlg_tmu_cod_municipio, tlg_tba_cod_bairro, tlg_ttl_cod_tip_logr, " & _
                " tlg_cod_logradouro, tlg_nome, " & _
                " tlg_cod_logradouro_inicial, tlg_cod_logradouro_final, " & _
                " tlg_cod_bairro_inicial, tlg_cod_bairro_final,tlg_cep"
        Valores = Bdados.PreparaValor(Temp.PegaParametro(Bdados, "MUNICIPIO"), txtBairro, cboTipo.Coluna(1).Valor, _
                    txtCodigo, txtDescricao, txtCodLogrInicial, txtCodLogrFinal, _
                    txtCodBairroInicial, txtCodBairroFinal, txtCep)
        If Bdados.GravaDados("TAB_LOGRADOURO", Valores, Campos, "tlg_cod_logradouro='" & txtCodigo & "'") Then
            Mensagem "Logradouro " & txtCodigo & " - " & txtDescricao & " atualizado com SUCESSO!"
            cmdNovo_Click
            txtCodigo = CodLogra
            cmdBuscar_Click
           
            LimpaCampos Me
        End If
        
    End If
End Sub

Private Sub atualizarLogradouro(Antigo As String, Novo As String)
    Dim Sql As String
    
    '1. Logradouro
    Sql = "UPDATE TAB_LOGRADOURO SET tlg_cod_logradouro='" & Novo & "' WHERE tlg_cod_logradouro='" & Antigo & "'"
    Bdados.Executa Sql
    Sql = "UPDATE TAB_LOGRADOURO SET tlg_cod_logradouro_inicial='" & Novo & "' WHERE tlg_cod_logradouro_inicial='" & Antigo & "'"
    Bdados.Executa Sql
    Sql = "UPDATE TAB_LOGRADOURO SET tlg_cod_logradouro_final='" & Novo & "' WHERE tlg_cod_logradouro_final='" & Antigo & "'"
    Bdados.Executa Sql
    '2. Trecho
    Sql = "UPDATE TAB_TRECHO SET TTC_TLG_COD_LOGRADOURO='" & Novo & "' WHERE TTC_TLG_COD_LOGRADOURO='" & Antigo & "'"
    Bdados.Executa Sql
    Sql = "UPDATE TAB_TRECHO SET TTC_LOGR_INICIAL='" & Novo & "' WHERE TTC_LOGR_INICIAL='" & Antigo & "'"
    Bdados.Executa Sql
    Sql = "UPDATE TAB_TRECHO SET TTC_LOGR_FINAL='" & Novo & "' WHERE TTC_LOGR_FINAL='" & Antigo & "'"
    Bdados.Executa Sql
    '3. Detalhe logradouro
    Sql = "UPDATE TAB_DETALHE_LOGRADOURO SET tdl_tlg_cod_logradouro='" & Novo & "' WHERE tdl_tlg_cod_logradouro='" & Antigo & "'"
    Bdados.Executa Sql
    '4. Imovel
    Sql = "UPDATE TAB_IMOVEL SET tim_tlg_cod_logradouro='" & Novo & "' WHERE tim_tlg_cod_logradouro='" & Antigo & "'"
    Bdados.Executa Sql
    '5. Valor terreno
    Sql = "UPDATE TAB_VALOR_TERRENO SET tvl_tlg_cod_logradouro='" & Novo & "' WHERE tvl_tlg_cod_logradouro='" & Antigo & "'"
    Bdados.Executa Sql
    Sql = "UPDATE TAB_VALOR_TERRENO SET tvl_logr_inicial='" & Novo & "' WHERE tvl_logr_inicial='" & Antigo & "'"
    Bdados.Executa Sql
    Sql = "UPDATE TAB_VALOR_TERRENO SET tvl_logr_final='" & Novo & "' WHERE tvl_logr_final='" & Antigo & "'"
    Bdados.Executa Sql
    '6. Contribuinte
    Sql = "UPDATE TAB_CONTRIBUINTE SET tci_cod_logradouro='" & Novo & "' WHERE tci_cod_logradouro='" & Antigo & "'"
    Bdados.Executa Sql
End Sub
Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name
    
    prepararControles
    CodLoteamento = Temp.PegaParametro(Bdados, "CODIGO LOTEAMENTO")
    CodDivisa = Temp.PegaParametro(Bdados, "CODIGO DIVISA")
    cboCaracterizacao.Visible = (CodLoteamento <> "")
End Sub

Private Sub prepararControles()
    Dim Sql As String
    
    Sql = "SELECT TTL_NOME, TTL_COD_TIP_LOGR FROM TAB_TIPO_LOGR ORDER BY TTL_NOME"
    cboTipo.Preencher Bdados, Sql
    
    Sql = "SELECT TBA_NOME, TBA_COD_BAIRRO FROM TAB_BAIRRO WHERE TBA_TMU_COD_MUNICIPIO=" & Temp.PegaParametro(Bdados, "MUNICIPIO") & " ORDER BY TBA_NOME"
    cboBairro.Preencher Bdados, Sql
    
    cboCaracterizacao.Clear
    cboCaracterizacao.AddItem "LOGRADOURO"
    cboCaracterizacao.AddItem "LOTEAMENTO"
    cboCaracterizacao.AddItem "DIVISA"
    
    cboRelatorio.Clear
    cboRelatorio.AddItem "CADASTRO"
    cboRelatorio.AddItem "RELAÇÃO"
End Sub

Private Function buscarBairro(ByRef Codigo As Object) As String
    Dim Sql As String
    
    If Nvl(Trim$(Codigo), 0) > 0 Then
        Sql = "SELECT TBA_NOME FROM TAB_BAIRRO WHERE TBA_COD_BAIRRO=" & Codigo
        If Bdados.AbreTabela(Sql) Then
            buscarBairro = Bdados.Tabela.Fields(0).Value
        Else
            Erro "Bairro não encontrado."
            Codigo = ""
        End If
        Bdados.FechaTabela
    End If
End Function

Private Function buscarLogradouro(ByRef Codigo As Object) As String
    Dim Sql As String
    
    If Nvl(Trim$(Codigo), 0) > 0 Then
        Sql = "SELECT tlg_nome FROM TAB_LOGRADOURO WHERE tlg_tmu_cod_municipio=" & Temp.PegaParametro(Bdados, "MUNICIPIO") & " AND tlg_cod_logradouro='" & Codigo & "'"
        If Bdados.AbreTabela(Sql) Then
            buscarLogradouro = Bdados.Tabela.Fields(0).Value
        Else
            Erro "Logradouro não encontrado."
            Codigo = ""
        End If
        Bdados.FechaTabela
    End If
End Function

Private Sub grdLogradouro_Click()
    If Not grdLogradouro.SelectedItem Is Nothing Then
        With grdLogradouro.SelectedItem
            CodAntigo = .Text
            txtCodigo = .Text
            cboTipo.SetarLinha .SubItems(3), 1
            txtDescricao = .SubItems(4)
            txtBairro = .SubItems(5)
                txtBairro_LostFocus
                
            txtCodLogrInicial = .SubItems(6)
            txtCodLogrInicial_LostFocus
            txtCodLogrInicial = .SubItems(6) 'Nao quero que apague se o codigo nao for encontrado.
            txtCodBairroInicial = .SubItems(7)
            txtCodBairroInicial_LostFocus
            
            txtCodLogrFinal = .SubItems(8)
            txtCodLogrFinal_LostFocus
            txtCodLogrFinal = .SubItems(8)
            txtCodBairroFinal = .SubItems(9)
            txtCodBairroFinal_LostFocus
            txtCep = .SubItems(10)
        End With
    End If
End Sub

Private Sub txtBairro_LostFocus()
    cboBairro = buscarBairro(txtBairro)
End Sub

Private Sub txtCodBairroFinal_LostFocus()
    txtBairroFinal = buscarBairro(txtCodBairroFinal)
End Sub

Private Sub txtCodBairroInicial_LostFocus()
    txtBairroInicial = buscarBairro(txtCodBairroInicial)
End Sub

Private Sub txtCodigo_LostFocus()
    Dim Codigo As String
    
    If Trim$(txtCodigo) <> "" Then
        Codigo = txtCodigo
        consultarLogradouros txtCodigo, "", "", "", "", ""
        'Edita.LimpaCampos Me
        txtCodigo = Codigo
        If grdLogradouro.ListItems.Count > 0 Then
            grdLogradouro.ListItems(1).Selected = True
            grdLogradouro_Click
        End If
    End If
End Sub

Private Sub txtCodLogrFinal_LostFocus()
    txtLogrFinal = buscarLogradouro(txtCodLogrFinal)
End Sub

Private Sub txtCodLogrInicial_LostFocus()
    txtLogrInicial = buscarLogradouro(txtCodLogrInicial)
End Sub

Private Sub consultarLogradouros(Logradouro As String, Setor As String, Bairro As String, Definicao As String, Nome As String, Tipo As String)
    Dim Sql As String, where As String
    
    where = ""
    
    If Trim$(Logradouro) <> "" Then where = where & " and TAB_LOGRADOURO.tlg_cod_logradouro='" & Logradouro & "'"
    If Trim$(Setor) <> "" Then where = where & " and TTC_SETOR=" & Setor
    If Trim$(Bairro) <> "" Then where = where & " and TAB_LOGRADOURO.tlg_tba_cod_bairro=" & Bairro
    If Trim$(Nome) <> "" Then where = where & " and TAB_LOGRADOURO.tlg_nome like '%" & Nome & "%'"
    If Trim$(Tipo) <> "" Then
        If Trim$(Setor) <> "" Then
            where = where & " and TAB_LOGRADOURO.tlg_nome like '%" & Tipo & "%'"
        Else
            where = where & " and TTL_NOME='" & Tipo & "'"
        End If
    End If
    
    
    Select Case Definicao
        Case "LOGRADOURO"
            where = where & " and CAST(" & Bdados.ParteTexto("TAB_LOGRADOURO.tlg_cod_logradouro", MidVs, 1, Len(CodLoteamento), True) & " AS smallint) <" & CodLoteamento
            grdLogradouro.CabecalhoTitulo = "RELAÇÃO LOGRADOUROS (SETOR " & Setor & ")"
        
        Case "LOTEAMENTO"
            where = where & " and cast(" & Bdados.ParteTexto("TAB_LOGRADOURO.tlg_cod_logradouro", MidVs, 1, Len(CodLoteamento), True) & " AS smallint) >= " & CodLoteamento & "  and cast(" & Bdados.ParteTexto("TAB_LOGRADOURO.tlg_cod_logradouro", MidVs, 1, Len(CodLoteamento), True) & " as smallint) <" & CodDivisa
            grdLogradouro.CabecalhoTitulo = "RELAÇÃO LOTEAMENTOS (SETOR " & Setor & ")"
    
        Case "DIVISA"
            where = where & " and " & Bdados.ParteTexto("TAB_LOGRADOURO.tlg_cod_logradouro", MidVs, 1, Len(CodDivisa), True) & "='" & CodDivisa & "'"
            grdLogradouro.CabecalhoTitulo = "RELAÇÃO DIVISAS (SETOR " & Setor & ")"
    End Select
    
    
    If Trim$(Setor) <> "" Then
        Sql = "SELECT distinct TAB_LOGRADOURO.tlg_cod_logradouro AS Codigo," & _
                    " LOGRADOURO as Logradouro," & _
                    " TBA_NOME as Bairro, " & _
                    " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr, " & _
                    " TAB_LOGRADOURO.tlg_nome, " & _
                    " TAB_LOGRADOURO.tlg_tba_cod_bairro," & _
                    " TAB_LOGRADOURO.tlg_cod_logradouro_inicial, " & _
                    " TAB_LOGRADOURO.tlg_cod_bairro_inicial, " & _
                    " TAB_LOGRADOURO.tlg_cod_logradouro_final, " & _
                    " TAB_LOGRADOURO.tlg_cod_bairro_final, " & _
                    " TAB_LOGRADOURO.tlg_cep as Cep " & _
                " FROM VIS_INFRA, TAB_BAIRRO, TAB_LOGRADOURO" & _
                " WHERE VIS_INFRA.tlg_tba_cod_bairro = tba_cod_bairro" & _
                    " AND VIS_INFRA.tlg_cod_logradouro = TAB_LOGRADOURO.tlg_cod_logradouro AND tlg_tmu_cod_municipio=" & Temp.PegaParametro(Bdados, "MUNICIPIO")
    Else
        Sql = "SELECT TAB_LOGRADOURO.tlg_cod_logradouro AS Codigo," & _
                    " TTL_NOME" & Bdados.Concatena & "' '" & Bdados.Concatena & "tlg_nome as Logradouro," & _
                    " TBA_NOME as Bairro, " & _
                    " TAB_LOGRADOURO.tlg_ttl_cod_tip_logr, " & _
                    " TAB_LOGRADOURO.tlg_nome, " & _
                    " TAB_LOGRADOURO.tlg_tba_cod_bairro," & _
                    " TAB_LOGRADOURO.tlg_cod_logradouro_inicial, " & _
                    " TAB_LOGRADOURO.tlg_cod_bairro_inicial, " & _
                    " TAB_LOGRADOURO.tlg_cod_logradouro_final, " & _
                    " TAB_LOGRADOURO.tlg_cod_bairro_final, " & _
                    " TAB_LOGRADOURO.tlg_cep as Cep " & _
                " FROM TAB_BAIRRO, TAB_LOGRADOURO, TAB_TIPO_LOGR" & _
                " WHERE tlg_tba_cod_bairro = tba_cod_bairro" & _
                    " AND tlg_ttl_cod_tip_logr=ttl_cod_tip_logr "
    End If
    Sql = Sql & where
    Sql = Sql & " ORDER BY TAB_LOGRADOURO.tlg_cod_logradouro"
    If grdLogradouro.Preencher(Bdados, Sql, 15, 50, 35, 0, 0, 0, 0, 0, 0, 0, 20) Then
        grdLogradouro.Mensagem = ""
    Else
        grdLogradouro.Mensagem = "Nenhum registro encontrado."
    End If
End Sub
