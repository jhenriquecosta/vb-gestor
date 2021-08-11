VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU402 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU402"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   1138
      Icone           =   "TCIU402.frx":0000
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   2655
      Left            =   15
      TabIndex        =   29
      Top             =   675
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   4683
      Altura          =   1905
      Caption         =   " Consultar Por:"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VB.Frame Frame1 
         Caption         =   "Data Alteração"
         Height          =   600
         Left            =   7455
         TabIndex        =   37
         Top             =   930
         Width           =   3675
         Begin VTOcx.txtVISUAL txtInicio 
            Height          =   285
            Left            =   270
            TabIndex        =   12
            Top             =   210
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   503
            Caption         =   "De"
            Text            =   ""
            Formato         =   0
         End
         Begin VTOcx.txtVISUAL txtAte 
            Height          =   285
            Left            =   1965
            TabIndex        =   13
            Top             =   210
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   503
            Caption         =   "Até"
            Text            =   ""
            Formato         =   0
         End
      End
      Begin VTOcx.txtVISUAL txtUsuario 
         Height          =   480
         Left            =   9045
         TabIndex        =   3
         Top             =   360
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   847
         Caption         =   "Atendente/Usuário"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VTOcx.txtVISUAL txtQuadraInsc 
         Height          =   480
         Left            =   6840
         TabIndex        =   8
         Top             =   -525
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   847
         Caption         =   "Quadra"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VTOcx.txtVISUAL txtBloco 
         Height          =   480
         Left            =   8010
         TabIndex        =   9
         Top             =   -525
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   847
         Caption         =   "Bloco"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VTOcx.txtVISUAL txtSala 
         Height          =   480
         Left            =   5910
         TabIndex        =   11
         Top             =   945
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   847
         Caption         =   "Sala/Loja"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtApto 
         Height          =   480
         Left            =   9060
         TabIndex        =   10
         Top             =   -525
         Visible         =   0   'False
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   847
         Caption         =   "Apto"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VTOcx.cboVISUAL cboCond 
         Height          =   510
         Left            =   165
         TabIndex        =   18
         Top             =   2040
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   900
         Caption         =   "Condominio"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.txtVISUAL txtICAnt 
         Height          =   480
         Left            =   4305
         TabIndex        =   6
         Top             =   945
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   847
         Caption         =   "Insc. Anterior"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VTOcx.txtVISUAL txtLote 
         Height          =   495
         Left            =   10215
         TabIndex        =   21
         Top             =   2055
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   "Lote:"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtNumero 
         Height          =   480
         Left            =   5925
         TabIndex        =   16
         Top             =   1470
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   847
         Caption         =   "Número"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VTOcx.cboVISUAL cboBairro 
         Height          =   510
         Left            =   6870
         TabIndex        =   17
         Top             =   1425
         Visible         =   0   'False
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   900
         Caption         =   "Bairro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboLoteamento 
         Height          =   510
         Left            =   4875
         TabIndex        =   19
         Top             =   2040
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   900
         Caption         =   "Loteamento"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboLogr 
         Height          =   315
         Left            =   1770
         TabIndex        =   15
         Top             =   1635
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTipoLogr 
         Height          =   510
         Left            =   150
         TabIndex        =   14
         Top             =   1440
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   900
         Caption         =   "Logradouro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         Editavel        =   -1  'True
      End
      Begin VB.TextBox txtIM 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   135
         MaxLength       =   11
         TabIndex        =   0
         Top             =   540
         Width           =   1680
      End
      Begin VB.TextBox txtContrib 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3765
         TabIndex        =   2
         Top             =   555
         Width           =   5175
      End
      Begin VB.TextBox txtIC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   150
         TabIndex        =   4
         Top             =   1140
         Width           =   1680
      End
      Begin VB.TextBox txtICAnterior 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1890
         TabIndex        =   5
         Top             =   1140
         Width           =   2310
      End
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9300
         MaxLength       =   5
         TabIndex        =   20
         Top             =   2265
         Width           =   810
      End
      Begin VB.TextBox txtSetor 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5925
         MaxLength       =   5
         TabIndex        =   7
         Top             =   -330
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox txtImanterior 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1890
         MaxLength       =   11
         TabIndex        =   1
         Top             =   555
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razão Social"
         Height          =   195
         Index           =   5
         Left            =   3795
         TabIndex        =   36
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cad. Contribuinte"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   35
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cad.Imobiliário"
         Height          =   195
         Index           =   25
         Left            =   135
         TabIndex        =   34
         Top             =   915
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cad. Auxiliar Imovel"
         Height          =   195
         Index           =   0
         Left            =   1875
         TabIndex        =   33
         Top             =   930
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra"
         Height          =   195
         Index           =   7
         Left            =   9300
         TabIndex        =   32
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Setor"
         Height          =   195
         Index           =   6
         Left            =   5910
         TabIndex        =   31
         Top             =   -540
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cad. Auxiliar Contribuinte"
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   30
         Top             =   330
         Width           =   1845
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   26
      Top             =   7155
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   979
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   405
         Left            =   6480
         TabIndex        =   22
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "&Listagem"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdCancelar 
         Height          =   405
         Left            =   7680
         TabIndex        =   23
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
         CorFoco         =   -2147483626
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   405
         Left            =   8880
         TabIndex        =   24
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   405
         Left            =   10080
         TabIndex        =   25
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.grdVISUAL grid 
      Height          =   3795
      Left            =   0
      TabIndex        =   27
      Top             =   3360
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   6694
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   345
      Left            =   4890
      TabIndex        =   28
      Top             =   60
      Width           =   855
   End
   Begin VB.Menu MnuDados 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu IteCad 
         Caption         =   "Consultar Cadastro"
         Index           =   1
      End
      Begin VB.Menu IteCad 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu IteCad 
         Caption         =   "Cancelar"
         Index           =   3
      End
   End
End
Attribute VB_Name = "TCIU402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelecaoRpt As String
Dim SelecaoRptSub As String
Private Sub cmdCancelar_Click()
    Dim Sql As String
    Dim Setor As String
    Dim Elo As String
    
    
    If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        If Bdados.Conexao.FormatoBanco = SQLServer Then
            Setor = "substring(tim_ic_auxiliar,3,2) AS Setor"
        ElseIf Bdados.Conexao.FormatoBanco = oracle Then
            If AplicacoesVTFuncoes.municipio = "BARRA MANSA" Then
                Setor = "SUBSTR(tim_ic_auxiliar,5,2) AS Setor"
            End If
        End If
    Else
        If Bdados.Conexao.FormatoBanco = SQLServer Then
            Setor = "substring(tim_ic,3,2) AS Setor"
        ElseIf Bdados.Conexao.FormatoBanco = oracle Then
            If AplicacoesVTFuncoes.municipio = "BARRA MANSA" Then
                Setor = "SUBSTR(tim_ic_auxiliar,5,2) AS Setor"
            End If
        End If
    End If
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "select tim_ic as  [Cad Imobiliária],tim_ic_auxiliar as [Insc Imobiliária],"
        Sql = Sql & " tim_ic_anterior as [Insc Anterior], tim_tci_im as [Cad Contribuinte],"
        Sql = Sql & "tci_nome as Contribuinte,TTL_NOME as Logradouro ,tlg_nome as Endereco,TBA_NOME as Bairro,tim_numero as Número,"
        Sql = Sql & Setor & ", tim_complemento as Compl ,TED_DESCRICAO AS Edificio,TIM_BLOCO as Bloco, TIM_APTO as Apto,TIM_SALA_LOJA AS [Sl/Lj], TLO_DESCRICAO as Loteamento,tim_lote as Lote,tim_quadra as Quadra"
        Sql = Sql & ", tim_valor_edific + tim_valor_terreno AS [Valor Venal], tim_valor_terreno as [Vl Terreno],tim_valor_edific as [Vl Edific]"
        Sql = Sql & " From vis_imovel_HISTORICO where 1 = 1"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "select tim_ic as Cad_Imobiliária,tim_ic_auxiliar as Insc_Imobiliária,"
        Sql = Sql & " tim_ic_anterior as Insc_Anterior, tim_tci_im as Cad_Contribuinte,"
        Sql = Sql & " tci_nome as Contribuinte,TTL_NOME as Logradouro ,tlg_nome as Endereco,tim_numero as Número,"
        Sql = Sql & Setor & ", tim_complemento as Compl ,TED_DESCRICAO AS Edificio,TIM_BLOCO as Bloco, TIM_APTO as Apto,TIM_SALA_LOJA AS Sl_Lj, TLO_DESCRICAO as Loteamento,tim_lote as Lote,tim_quadra as Quadra"
        Sql = Sql & ", tim_valor_edific + tim_valor_terreno AS Valor_Venal, tim_valor_terreno as Vl_Terreno,tim_valor_edific as Vl_Edific,TIM_MOTIVO AS MOTIVO,TIM_TUS_USUARIO as Usuário,TIM_DATA AS Data"
        Sql = Sql & " From vis_imovel_HISTORICO where 1 = 1"
    End If
    If txtImanterior <> "" Then
        Sql = Sql & " and tim_ic_anterior = '" & txtImanterior & "'"
    End If
    
    If txtSetor <> "" Then
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            If Bdados.Conexao.FormatoBanco = SQLServer Then
                Sql = Sql & " and substring(tim_ic_auxiliar,3,2) = '" & txtSetor & "'"
            ElseIf Bdados.Conexao.FormatoBanco = oracle Then
                If AplicacoesVTFuncoes.municipio = "BARRA MANSA" Then
                    Sql = Sql & " AND  SUBSTR(tim_ic_auxiliar,5,2) = '" & txtSetor & "'"
                End If
            End If
        Else
            If Bdados.Conexao.FormatoBanco = SQLServer Then
                Sql = Sql & " and substring(tim_ic,3,2) = '" & txtSetor & "'"
            ElseIf Bdados.Conexao.FormatoBanco = oracle Then
                If AplicacoesVTFuncoes.municipio = "BARRA MANSA" Then
                    Sql = Sql & " and SUBSTR(tim_ic,5,2) = '" & txtSetor & "'"
                End If
            End If
        End If
    End If
    
    If txtQuadraInsc <> "" Then
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            If Bdados.Conexao.FormatoBanco = SQLServer Then
                Sql = Sql & " and substring(tim_ic_auxiliar,5," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ") = '" & txtQuadraInsc & "'"
            ElseIf Bdados.Conexao.FormatoBanco = oracle Then
                If AplicacoesVTFuncoes.municipio = "BARRA MANSA" Then
                    Sql = Sql & " and SUBSTR(tim_ic_auxiliar,7," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ") = '" & Format(txtQuadraInsc, "000") & "'"
                End If
            End If
        Else
            If Bdados.Conexao.FormatoBanco = SQLServer Then
                Sql = Sql & " and substring(tim_ic,5," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ") = '" & txtQuadraInsc & "'"
            ElseIf Bdados.Conexao.FormatoBanco = oracle Then
                If AplicacoesVTFuncoes.municipio = "BARRA MANSA" Then
                    Sql = Sql & " and SUBSTR(tim_ic,7," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ") = '" & txtQuadraInsc & "'"
                End If
            End If
        End If
    End If
    
    If txtInicio <> "" Then
        Sql = Sql & " and tim_data >= '" & txtInicio & "'"
    End If
    
    If txtAte <> "" Then
        Sql = Sql & " and tim_data <= '" & txtAte & "'"
    End If
    
    If txtUsuario <> "" Then
        Sql = Sql & " and TIM_TUS_USUARIO like '" & txtUsuario & "%'"
    End If
    
    If txtQuadra <> "" Then
        Sql = Sql & " and tim_quadra  = '" & txtQuadra & "'"
    End If
    If txtComplemento <> "" Then
        Sql = Sql & " and tim_complemento like '" & txtComplemento & "%'"
    End If
    
    If txtICAnt <> "" Then
        Sql = Sql & " and tim_ic_anterior  =  '" & Edita.TiraTudo(txtICAnt) & "'"
    End If
    If txtIM <> "" Then
        Sql = Sql & " and tim_tci_im  =  '" & txtIM & "'"
    End If
    
    If txtContrib <> "" Then
        Sql = Sql & " and tci_nome like '" & txtContrib & "%'"
    End If
    
    If txtic <> "" Then
        Sql = Sql & " and tim_ic  = '" & UCase(txtic) & "'"
    End If
    
    If txtICAnterior <> "" Then
        Sql = Sql & " and tim_ic_auxiliar  = '" & Edita.TiraTudo(txtICAnterior) & "'"
    End If
    
    If cboTipoLogr.ListIndex >= 0 Then
        Sql = Sql & " and TTL_NOME like '" & UCase(cboTipoLogr.Text) & "'"
    End If
    
    If cboLogr.ListIndex >= 0 Then
        Sql = Sql & " and tlg_nome like '" & UCase(cboLogr) & "%'"
    End If
    
    If txtNumero <> "" Then
        Sql = Sql & " and     tim_numero  = '" & UCase(txtNumero) & "'"
    End If
    If cboBairro.ListIndex >= 0 Then
        Sql = Sql & " and TBA_NOME like '" & UCase(cboBairro.Text) & "'"
    End If
    
    If cboLoteamento.ListIndex >= 0 Then
        Sql = Sql & " and TLO_DESCRICAO like '" & UCase(cboLoteamento.Text) & "%'"
    End If
    If cboCond.ListIndex >= 0 Then
        Sql = Sql & " and TIM_TED_COD_EDIFICIO= '" & cboCond.Coluna(0).Valor & "'"
    End If
    If txtLote <> "" Then
        Sql = Sql & " and tim_lote  = '" & txtLote & "'"
    End If
    
    If Trim(txtBloco) <> "" Then
        Sql = Sql & " and tim_bloco  = '" & txtBloco & "'"
    End If
    If Trim(txtApto) <> "" Then
        Sql = Sql & " and tim_apto  = '" & txtApto & "'"
    End If
    If Trim(txtSala) <> "" Then
        Sql = Sql & " and TIM_SALA_LOJA  = '" & txtSala & "'"
    End If
    '***********************************************************************************************
    'SelecaoRpt = "{TAB_BAIRRO.TBA_TMU_COD_MUNICIPIO}= '" & Aplicacoes.Codigo_Municipio & "'"
    'SelecaoRpt = SelecaoRpt & " and {TAB_LOGRADOURO.tlg_tmu_cod_municipio}="  & Aplicacoes.Codigo_Municipio
    
    'IC
    SelecaoRpt = " 1 = 1"
    If Trim(txtic) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.tim_ic}='" & txtic & "'"
    End If
    If Trim(txtNumero) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.tim_numero}='" & txtNumero & "'"
    End If
    'IM
    If cboLoteamento.ListIndex >= 0 Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.tim_loteamento}= '" & cboLoteamento.Coluna(0).Valor & "'"
    End If
    If cboCond.ListIndex >= 0 Then
        If Bdados.Conexao.FormatoBanco = SQLServer Then
            SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.TIM_TED_COD_EDIFICIO}= '" & cboCond.Coluna(0).Valor & "'"
        ElseIf Bdados.Conexao.FormatoBanco = oracle Then
            SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.TIM_TED_COD_EDIFICIO}= '" & cboCond.Coluna(0).Valor & "'"
        End If
    End If
    
    
    If txtInicio <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.tim_data} >= '" & txtInicio & "'"
    End If
    
    If txtAte <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.tim_data} <= '" & txtAte & "'"
    End If
    
    If Trim(txtIM) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.tim_tci_im}='" & txtIM & "'"
    End If
    'Tipo Logradouro
    If Trim(cboTipoLogr) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.TTL_NOME}='" & cboTipoLogr & "'"
    End If
    If txtUsuario <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.TIM_TUS_USUARIO} like '" & txtUsuario & "*'"
    End If
    'Logradouro
    If Trim(cboLogr) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.tlg_nome}= '" & cboLogr & "'"
    End If
    'Bairro
    If Trim(cboBairro) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.TBA_NOME}= '" & cboBairro & "'"
    End If
    If Trim(txtBloco) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.tim_bloco}  = '" & txtBloco & "'"
    End If
    If Trim(txtApto) <> "" Then
        SelecaoRpt = SelecaoRpt & "  and {VIS_IMOVEL_HISTORICO.tim_apto}  = '" & txtApto & "'"
    End If
    If Trim(txtSala) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.TIM_SALA_LOJA}  = '" & txtSala & "'"
    End If
    'Razao Social
    If Trim(txtContrib) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.tci_nome} like '" & txtContrib & "*'"
    End If
    'Quadra
    If Trim(txtSetor) <> "" Then
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            SelecaoRpt = SelecaoRpt & " and Mid({VIS_IMOVEL_HISTORICO.tim_ic_AUXILIAR},3,2)='" & txtSetor & "'"
        Else
            SelecaoRpt = SelecaoRpt & " and Mid({VIS_IMOVEL_HISTORICO.tim_ic},3,2)='" & txtSetor & "'"
        End If
    End If
    
    If txtQuadra <> "" Then
        SelecaoRpt = SelecaoRpt & " and {VIS_IMOVEL_HISTORICO.tim_quadra}  = '" & txtQuadra & "'"
    End If
    'Quadra
    If Trim(txtQuadraInsc) <> "" Then
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            SelecaoRpt = SelecaoRpt & " and Mid({VIS_IMOVEL_HISTORICO.tim_ic_AUXILIAR},5," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ")='" & txtQuadraInsc & "'"
        Else
            SelecaoRpt = SelecaoRpt & " and Mid({VIS_IMOVEL_HISTORICO.tim_ic},5," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ")='" & txtQuadraInsc & "'"
        End If
    End If
    
    '**********************************************************************************************
     If Not grid.Preencher(Bdados, Sql) Then
        Util.Avisa "Consulta sem resultados."
     End If
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{tab}"
End Sub

Private Sub cmdImprimir_Click()
    
    Screen.MousePointer = 11
    If grid.ListItems.Count > 0 Then
        With Rpt
            If Not .DefinirArquivo(Bdados, App.Path & "\TCIU402.rpt") Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
            .Selecao = SelecaoRpt
            .Titulo = "Ficha Cadastral"
            .Arvore = False
            .Visualizar
            DoEvents
        End With
    End If
    Set Rpt = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grid.ListItems.Clear
    txtIM.SetFocus
End Sub
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    
    Screen.MousePointer = 11
    If grid.ListItems.Count > 0 Then
        With Rpt
            If Not .DefinirArquivo(Bdados, App.Path & "\TCIU402.rpt") Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
            .Selecao = SelecaoRpt
            .Titulo = "Listagem de Imóveis - Histórico"
            .Arvore = False
            .Visualizar
            DoEvents
        End With
    End If
    Set Rpt = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    cboLogr.Preencher Bdados, "Select tlg_cod_logradouro,(tlg_nome)   From Tab_Logradouro order by tlg_nome", 1
    cboTipoLogr.Preencher Bdados, "Select TTL_COD_TIP_LOGR ,(ttl_nome) From Tab_Tipo_Logr", 1
    cboBairro.Preencher Bdados, "Select DISTINCT(tba_nome),tba_cod_bairro From Tab_Bairro "
    cboLoteamento.Preencher Bdados, "Select TLO_COD_LOTEAMENTO,TLO_DESCRICAO from TAB_LOTEAMENTO ORDER BY TLO_DESCRICAO", 1
    cboCond.Preencher Bdados, "Select TED_COD_EDIFICIO,TED_DESCRICAO from TAB_EDIFICIO ORDER BY TED_DESCRICAO", 1
   ' cabVisual.Exibir Bdados, Me.Name, App.Path
    If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
        cmdVISUAL1.Enabled = True
    Else
        cmdVISUAL1.Enabled = False
    End If
End Sub

Private Sub grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu MnuDados
    End If
    
End Sub

Private Sub IteCad_Click(Index As Integer)
    If grid.ListItems.Count <= 0 Then Exit Sub
    Select Case Index
        Case 1
            TCIU402a.Tag = Trim(grid.SelectedItem) & Format(Trim(grid.SelectedItem.SubItems(8)), "00000")
            TCIU402a.Show
    End Select
End Sub

Private Sub txtContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtic_KeyPress(KeyAscii As Integer)
KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtICAnt_LostFocus()
    If Len(Trim(txtICAnt)) <> 18 Then Exit Sub
    txtICAnt = Edita.TiraTudo(txtICAnt)
    txtICAnt = Left(txtICAnt, 1) & "." & Mid(txtICAnt, 2, 4) & "." & Mid(txtICAnt, 6, 3) & "." & Mid(txtICAnt, 9, 2) & "." & Mid(txtICAnt, 11, 4) & "." & Right(txtICAnt, 4)
End Sub

Private Sub txtIcAnterior_KeyPress(KeyAscii As Integer)
KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtICAnterior_LostFocus()
    If Len(Trim(txtICAnterior)) <> 14 Then Exit Sub
    txtICAnterior = Edita.TiraTudo(txtICAnterior)
    txtICAnterior = Left(txtICAnterior, 2) & "." & Mid(txtICAnterior, 3, 2) & "." & Mid(txtICAnterior, 5, 3) & "." & Mid(txtICAnterior, 9, 4) & "." & Right(txtICAnterior, 3)
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    txtIM = Imposto.FormataInscricao(txtIM, InscContrib)
End Sub

Private Sub txtQuadra_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.Maiuscula(KeyAscii)
End Sub


Private Sub txtSetor_KeyPress(KeyAscii As Integer)
KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub
