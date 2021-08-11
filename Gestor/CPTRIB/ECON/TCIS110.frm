VERSION 5.00
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIS110 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIS110"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TCIS110.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs TabDados 
      Align           =   2  'Align Bottom
      Height          =   5715
      Left            =   0
      TabIndex        =   10
      Top             =   690
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   10081
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      TagVariant      =   ""
      Tabs            =   "TCIS110.frx":0342
      Images          =   "TCIS110.frx":03BD
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5325
         Left            =   -99969
         TabIndex        =   14
         Top             =   30
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   9393
         _Version        =   131082
         TabGuid         =   "TCIS110.frx":0A05
         Begin VTOcx.txtVISUAL txtPercentualGeral 
            Height          =   285
            Left            =   6195
            TabIndex        =   18
            Tag             =   "Valor"
            Top             =   345
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   503
            Caption         =   "Valor %"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoTexto=   1
         End
         Begin VTOcx.grdVISUAL grdValoresGeral 
            Height          =   4470
            Left            =   195
            TabIndex        =   15
            Top             =   750
            Width           =   8445
            _ExtentX        =   14896
            _ExtentY        =   7885
            CorBorda        =   16711680
            Caption         =   "Valores"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
         Begin VTOcx.txtVISUAL txtLimiteInferiorGeral 
            Height          =   285
            Left            =   285
            TabIndex        =   16
            Tag             =   "Limite Inferior"
            Top             =   345
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   503
            Caption         =   "Limite Inferior"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
         End
         Begin VTOcx.txtVISUAL txtLimiteSuperiorGeral 
            Height          =   285
            Left            =   3360
            TabIndex        =   17
            Tag             =   "Limite Superior"
            Top             =   345
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   503
            Caption         =   "Limite Superior"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5325
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   9393
         _Version        =   131082
         TabGuid         =   "TCIS110.frx":0A2D
         Begin VTOcx.txtVISUAL txtPercentual 
            Height          =   285
            Left            =   6360
            TabIndex        =   4
            Tag             =   "Valor"
            Top             =   840
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   503
            Caption         =   "Valor %"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
            AlinhamentoTexto=   1
         End
         Begin VTOcx.grdVISUAL GrdValores 
            Height          =   1800
            Left            =   285
            TabIndex        =   12
            Top             =   3510
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3175
            CorBorda        =   16711680
            Caption         =   "Valores"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
         Begin VTOcx.grdVISUAL GrdBairros 
            Height          =   2265
            Left            =   270
            TabIndex        =   13
            Top             =   1200
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3995
            CorBorda        =   16711680
            Caption         =   "Bairros"
            CorTitulo       =   16711680
            CorCaption      =   16777215
            CorDica         =   16711680
         End
         Begin VTOcx.txtVISUAL txtLimiteInferior 
            Height          =   285
            Left            =   675
            TabIndex        =   2
            Tag             =   "Limite Inferior"
            Top             =   825
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   503
            Caption         =   "Limite Inferior"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
         End
         Begin VTOcx.txtVISUAL txtDescAtividade 
            Height          =   285
            Left            =   1050
            TabIndex        =   1
            Top             =   495
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   503
            Caption         =   "Descrição"
            Text            =   ""
            Enabled         =   0   'False
         End
         Begin VTOcx.txtVISUAL txtCodigo 
            Height          =   285
            Left            =   1290
            TabIndex        =   0
            Tag             =   "Código"
            Top             =   150
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   503
            Caption         =   "Código"
            Text            =   ""
         End
         Begin VTOcx.txtVISUAL txtLimiteSuperior 
            Height          =   285
            Left            =   3675
            TabIndex        =   3
            Tag             =   "Limite Superior"
            Top             =   825
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   503
            Caption         =   "Limite Superior"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   9
      Top             =   6405
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   1032
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   5490
         TabIndex        =   6
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   6630
         TabIndex        =   7
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
         Left            =   4350
         TabIndex        =   5
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
         TabIndex        =   8
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
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   1138
      Icone           =   "TCIS110.frx":0A55
   End
End
Attribute VB_Name = "TCIS110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExcluir_Click()
    Dim Condicao As String
    If TabDados.Tabs(1).Selected Then
        If GrdValores.ListItems.Count >= 1 Then
           If Confirma("Deseja excluir o registro selecionado?", "Mensagem") Then
                Condicao = "TDA_BAIRRO = " & txtCodigo & " aND TDA_VALOR_INICIAL = " & Bdados.Converte(txtLimiteInferior, TCMonetario)
                If Bdados.DeletaDados("tab_desconto_alvara", Condicao) Then
                    Avisa "Valor excluído com sucesso."
                    cmdLimpar_Click
                    PreencherGrid
                End If
           End If
        End If
    Else
        If grdValoresGeral.ListItems.Count >= 1 Then
           If Confirma("Deseja excluir o registro selecionado?", "Mensagem") Then
                Condicao = "TDG_VALOR_INICIAL = " & Bdados.Converte(txtLimiteInferiorGeral, TCMonetario)
                If Bdados.DeletaDados("tab_desconto_alvara_geral", Condicao) Then
                    Avisa "Valor excluído com sucesso."
                    cmdLimpar_Click
                    PreencherGrid
                End If
           End If
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    
    txtLimiteInferior = ""
    txtLimiteSuperior = ""
    txtPercentual = ""
    
    txtLimiteInferiorGeral = ""
    txtLimiteSuperiorGeral = ""
    txtPercentualGeral = ""
    If TabDados.Tabs(1).Selected Then
        txtLimiteInferior.SetFocus
    Else
        txtLimiteInferiorGeral.SetFocus
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub
Private Sub PreencherGrid()
    Dim Sql As String
    
    If TabDados.Tabs(1).Selected Then
        Sql = "select tda_bairro as Bairro,"
        Sql = Sql & " tda_valor_inicial as Limite_Inferior,"
        Sql = Sql & " tda_Valor_final As Limite_Superior,"
        Sql = Sql & " tda_percentual As VALOR"
        Sql = Sql & " From tab_desconto_alvara"
        Sql = Sql & " where tda_bairro = " & txtCodigo
        Sql = Sql & " order by 1,2"
        GrdValores.Preencher Bdados, Sql
    Else
        Sql = "select tdG_valor_inicial as Limite_Inferior,"
        Sql = Sql & " tdG_Valor_final As Limite_Superior,"
        Sql = Sql & " tdG_percentual As VALOR"
        Sql = Sql & " From tab_desconto_alvara_GERAL"
        Sql = Sql & " order by 1,2"
        grdValoresGeral.Preencher Bdados, Sql
    End If
    
End Sub
Private Sub cmdSalvar_Click()
    Dim Valores  As String
    Dim Campos   As String
    Dim Condicao As String
    
    
    If TabDados.Tabs(1).Selected Then
        Campos = "TDA_BAIRRO,TDA_VALOR_INICIAL,TDA_VALOR_FINAL,TDA_PERCENTUAL"
        Valores = Bdados.PreparaValor(txtCodigo, Bdados.Converte(txtLimiteInferior, TCMonetario), Bdados.Converte(txtLimiteSuperior, TCMonetario), Bdados.Converte(txtPercentual, TCMonetario))
        Condicao = "TDA_BAIRRO = " & txtCodigo & " aND TDA_VALOR_INICIAL = " & Bdados.Converte(txtLimiteInferior, TCMonetario)
        If Bdados.GravaDados("TAB_DESCONTO_ALVARA", Valores, Campos, Condicao) Then
            Avisa "Operação concluída com sucesso."
            cmdLimpar_Click
            PreencherGrid
        End If
    Else
        Campos = "TDG_VALOR_INICIAL,TDG_VALOR_FINAL,TDG_PERCENTUAL"
        Valores = Bdados.PreparaValor(Bdados.Converte(txtLimiteInferiorGeral, TCMonetario), Bdados.Converte(txtLimiteSuperiorGeral, TCMonetario), Bdados.Converte(txtPercentualGeral, TCMonetario))
        Condicao = "TDG_VALOR_INICIAL = " & Bdados.Converte(txtLimiteInferiorGeral, TCMonetario)
        If Bdados.GravaDados("TAB_DESCONTO_ALVARA_geral", Valores, Campos, Condicao) Then
            Avisa "Operação concluída com sucesso."
            cmdLimpar_Click
            PreencherGrid
        End If
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    GrdBairros.Preencher Bdados, "SELECT TBA_COD_BAIRRO as Código,TBA_NOME  as Nome FROM tab_bairro ORDER BY 2"
End Sub

Private Sub GrdBairros_DblClick()
    If GrdBairros.ListItems.Count >= 1 Then
        txtCodigo = GrdBairros.SelectedItem
        txtDescAtividade = GrdBairros.SelectedItem.SubItems(1)
        txtCodigo_LostFocus
    End If
End Sub

Private Sub GrdValores_DblClick()
    If GrdValores.ListItems.Count >= 1 Then
        txtLimiteInferior = GrdValores.SelectedItem.SubItems(1)
        txtLimiteSuperior = GrdValores.SelectedItem.SubItems(2)
        txtPercentual = GrdValores.SelectedItem.SubItems(3)
    End If
End Sub

Private Sub grdValoresGeral_DblClick()
    If grdValoresGeral.ListItems.Count >= 1 Then
        txtLimiteInferiorGeral = grdValoresGeral.SelectedItem
        txtLimiteSuperiorGeral = grdValoresGeral.SelectedItem.SubItems(1)
        txtPercentualGeral = grdValoresGeral.SelectedItem.SubItems(2)
    End If
End Sub

Private Sub TabDados_Click()
    If TabDados.Tabs(1).Selected Then
        txtCodigo.SetFocus
    Else
        txtLimiteInferiorGeral.SetFocus
        PreencherGrid
    End If
End Sub

Private Sub txtCodigo_LostFocus()
    PreencherGrid
End Sub
