VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIU404 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCIU404"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11220
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
   ScaleHeight     =   7695
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   3465
      Left            =   0
      TabIndex        =   31
      Top             =   660
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   6112
      Altura          =   1905
      Caption         =   " Consultar Por:"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtOcupante 
         Height          =   510
         Left            =   2400
         TabIndex        =   42
         Top             =   960
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   900
         Caption         =   "Nome"
         Text            =   ""
         Enabled         =   0   'False
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtCPFOcupante 
         Height          =   510
         Left            =   8520
         TabIndex        =   41
         Top             =   960
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   900
         Caption         =   "CPF/CNPJ"
         Text            =   ""
         Enabled         =   0   'False
         Restricao       =   2
         AlinhamentoRotulo=   1
         CorFundo        =   -2147483644
      End
      Begin VTOcx.cmdVISUAL CmdConsultaContribuinteOcupante 
         Height          =   315
         Left            =   2040
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1150
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   556
         Caption         =   ""
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.txtVISUAL txtInscMunicipalOcupante 
         Height          =   510
         Left            =   165
         TabIndex        =   4
         Top             =   960
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   900
         Caption         =   "IM Ocupante"
         Text            =   ""
         Restricao       =   2
         AlinhamentoRotulo=   1
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.cboVISUAL cboUsuario 
         Height          =   510
         Left            =   8970
         TabIndex        =   21
         Top             =   2580
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   900
         Caption         =   "Cadastrador"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         ValorPadrao     =   "TIPO LOTE"
      End
      Begin VTOcx.cboVISUAL cboTipo 
         Height          =   510
         Left            =   8580
         TabIndex        =   3
         Top             =   420
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   900
         Caption         =   "Tipo"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
         ValorPadrao     =   "TIPO LOTE"
      End
      Begin VTOcx.txtVISUAL txtQuadraInsc 
         Height          =   480
         Left            =   6705
         TabIndex        =   9
         Top             =   1590
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   847
         Caption         =   "Quadra"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VTOcx.txtVISUAL txtBloco 
         Height          =   480
         Left            =   7935
         TabIndex        =   10
         Top             =   1590
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   847
         Caption         =   "Bloco"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VTOcx.txtVISUAL txtSala 
         Height          =   495
         Left            =   10065
         TabIndex        =   12
         Top             =   1575
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Caption         =   "Sala/Loja"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtApto 
         Height          =   480
         Left            =   9015
         TabIndex        =   11
         Top             =   1590
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   847
         Caption         =   "Apto"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VTOcx.cboVISUAL cboCond 
         Height          =   510
         Left            =   165
         TabIndex        =   19
         Top             =   2595
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
         Left            =   4320
         TabIndex        =   7
         Top             =   1590
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   847
         Caption         =   "Insc. Anterior"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
         RetirarMascara  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtLote 
         Height          =   495
         Left            =   10095
         TabIndex        =   18
         Top             =   2085
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   873
         Caption         =   "Lote:"
         Text            =   ""
         AlinhamentoRotulo=   1
      End
      Begin VTOcx.txtVISUAL txtNumero 
         Height          =   480
         Left            =   4860
         TabIndex        =   15
         Top             =   2085
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   847
         Caption         =   "Número"
         Text            =   ""
         AlinhamentoRotulo=   1
         AlinhamentoRotuloVertical=   0
      End
      Begin VTOcx.cboVISUAL cboBairro 
         Height          =   510
         Left            =   5595
         TabIndex        =   16
         Top             =   2070
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   900
         Caption         =   "Bairro"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboLoteamento 
         Height          =   510
         Left            =   4875
         TabIndex        =   20
         Top             =   2595
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   900
         Caption         =   "Loteamento"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
      Begin VTOcx.cboVISUAL cboLogr 
         Height          =   315
         Left            =   1785
         TabIndex        =   14
         Top             =   2265
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         Caption         =   ""
         Text            =   ""
         AutoFocaliza    =   0   'False
      End
      Begin VTOcx.cboVISUAL cboTipoLogr 
         Height          =   510
         Left            =   165
         TabIndex        =   13
         Top             =   2085
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
         Left            =   165
         MaxLength       =   11
         TabIndex        =   0
         Top             =   645
         Width           =   1680
      End
      Begin VB.TextBox txtContrib 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3840
         TabIndex        =   2
         Top             =   645
         Width           =   4605
      End
      Begin VB.TextBox txtIC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   165
         TabIndex        =   5
         Top             =   1785
         Width           =   1680
      End
      Begin VB.TextBox txtICAnterior 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1905
         TabIndex        =   6
         Top             =   1785
         Width           =   2310
      End
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9030
         MaxLength       =   5
         TabIndex        =   17
         Top             =   2295
         Width           =   810
      End
      Begin VB.TextBox txtSetor 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5940
         MaxLength       =   5
         TabIndex        =   8
         Top             =   1785
         Width           =   630
      End
      Begin VB.TextBox txtImanterior 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1920
         MaxLength       =   11
         TabIndex        =   1
         Top             =   645
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razão Social do Prorietário"
         Height          =   195
         Index           =   5
         Left            =   3855
         TabIndex        =   38
         Top             =   435
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra"
         Height          =   195
         Index           =   7
         Left            =   9030
         TabIndex        =   37
         Top             =   2085
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Setor"
         Height          =   195
         Index           =   6
         Left            =   5955
         TabIndex        =   36
         Top             =   1590
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição Proprietário"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   35
         Top             =   420
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inscri.Imobiliária"
         Height          =   195
         Index           =   0
         Left            =   1890
         TabIndex        =   34
         Top             =   1590
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cad.Imobiliário"
         Height          =   195
         Index           =   25
         Left            =   165
         TabIndex        =   33
         Top             =   1590
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cad. Contribuinte"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   32
         Top             =   420
         Width           =   1275
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   28
      Top             =   7140
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   979
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL CmdResumo 
         Height          =   405
         Left            =   3960
         TabIndex        =   22
         Top             =   60
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "&Cálculo"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdImprimirListagem 
         Height          =   405
         Left            =   6360
         TabIndex        =   24
         Top             =   60
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "&Listagem"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   405
         Left            =   5160
         TabIndex        =   23
         Top             =   60
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
         Caption         =   "&Ficha"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdCancelar 
         Height          =   405
         Left            =   7560
         TabIndex        =   25
         Top             =   60
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
         Left            =   8760
         TabIndex        =   26
         Top             =   60
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
         Left            =   9990
         TabIndex        =   27
         Top             =   60
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
      Height          =   3030
      Left            =   -15
      TabIndex        =   29
      Top             =   4140
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   5345
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   345
      Left            =   4890
      TabIndex        =   30
      Top             =   -375
      Visible         =   0   'False
      Width           =   855
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   1138
      Icone           =   "TCIU404.frx":0000
   End
   Begin VB.Menu MnuDados 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu IteCad 
         Caption         =   "Consultar Lançamentos"
         Index           =   0
      End
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
Attribute VB_Name = "TCIU404"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelecaoRpt As String
Dim SelecaoRptListagem As String

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
            Else
                Setor = "substr(tim_ic_auxiliar,3,2) AS Setor"
            End If
        End If
    Else
        If Bdados.Conexao.FormatoBanco = SQLServer Then
            If AplicacoesVTFuncoes.municipio = "SANTA MARIA DA BOA VISTA" Then
                Setor = "substring(tim_ic,2,4) AS Setor"
            ElseIf AplicacoesVTFuncoes.municipio <> "COLINAS" Then
                Setor = "substring(tim_ic,3,2) AS Setor"
            Else
                Setor = "SUBSTRING(tim_ic_anterior,3,2) AS Setor"
            End If
        ElseIf Bdados.Conexao.FormatoBanco = oracle Then
            If AplicacoesVTFuncoes.municipio = "BARRA MANSA" Then
                Setor = "SUBSTR(tim_ic_auxiliar,5,2) AS Setor"
            End If
        End If
    End If
    If Bdados.Conexao.FormatoBanco = SQLServer Then
        Sql = "select tim_ic as  [Cad Imobiliária],tim_ic_auxiliar as [Insc Imobiliária],"
        Sql = Sql & " tim_ic_anterior as [Insc Anterior], tim_tci_im as [Cad Contribuinte],"
        Sql = Sql & "tci_nome as Contribuinte,tim_ocupante as Ocupante,TTL_NOME as Logradouro ,tlg_nome as Endereco,TBA_NOME as Bairro,tim_numero as Número,"
        Sql = Sql & Setor & ", tim_complemento as Compl ,TED_DESCRICAO AS Edificio,TIM_BLOCO as Bloco, TIM_APTO as Apto,TIM_SALA_LOJA AS [Sl/Lj], TLO_DESCRICAO as Loteamento,tim_lote as Lote,tim_quadra as Quadra"
        Sql = Sql & ", tim_valor_edific + tim_valor_terreno AS [Valor Venal], tim_valor_terreno as [Vl Terreno],tim_valor_edific as [Vl Edific],tim_tus_cod_usuario as Cadastrador"
        Sql = Sql & " From vis_imovel where 1 = 1"
    ElseIf Bdados.Conexao.FormatoBanco = oracle Then
        Sql = "select tim_ic as Cad_Imobiliária,tim_ic_auxiliar as Insc_Imobiliária,"
        Sql = Sql & " tim_ic_anterior as Insc_Anterior, tim_tci_im as Cad_Contribuinte,"
        Sql = Sql & "tci_nome as Contribuinte,TTL_NOME as Logradouro ,tlg_nome as Endereco,TBA_NOME as Bairro,tim_numero as Número,"
        Sql = Sql & Setor & ", tim_complemento as Compl ,TED_DESCRICAO AS Edificio,TIM_BLOCO as Bloco, TIM_APTO as Apto,TIM_SALA_LOJA AS Sl_Lj, TLO_DESCRICAO as Loteamento,tim_lote as Lote,tim_quadra as Quadra"
        Sql = Sql & ", tim_valor_edific + tim_valor_terreno AS Valor_Venal, tim_valor_terreno as Vl_Terreno,tim_valor_edific as Vl_Edific, tim_tus_cod_usuario as Cadastrador"
        Sql = Sql & " From vis_imovel where 1 = 1"
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
                If AplicacoesVTFuncoes.municipio = "SANTA MARIA DA BOA VISTA" Then
                    Sql = Sql & " and substring(tim_ic,2,4) = '" & txtSetor & "'"
                ElseIf AplicacoesVTFuncoes.municipio <> "COLINAS" Then
                    Sql = Sql & " and substring(tim_ic,3,2) = '" & txtSetor & "'"
                Else
                    Sql = Sql & " and SUBSTRING(tim_ic_anterior,3,2) = '" & txtSetor & "'"
                End If
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
                If AplicacoesVTFuncoes.municipio = "SANTA MARIA DA BOA VISTA" Then
                    Sql = Sql & " and substring(tim_ic,6," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ") = '" & txtQuadraInsc & "'"
                ElseIf AplicacoesVTFuncoes.municipio <> "COLINAS" Then
                    Sql = Sql & " and substring(tim_ic,5," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ") = '" & txtQuadraInsc & "'"
                Else
                    Sql = Sql & " and substring(tim_ic_anterior,5," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ") = '" & txtQuadraInsc & "'"
                End If
            ElseIf Bdados.Conexao.FormatoBanco = oracle Then
                If AplicacoesVTFuncoes.municipio = "BARRA MANSA" Then
                    Sql = Sql & " and SUBSTR(tim_ic,7," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ") = '" & txtQuadraInsc & "'"
                End If
            End If
        End If
    End If
    
    If txtQuadra <> "" Then
        Sql = Sql & " and tim_quadra  = '" & txtQuadra & "'"
    End If
    If cboUsuario <> "" Then
        Sql = Sql & " and TIM_TUS_COD_USUARIO  = '" & cboUsuario & "'"
    End If
    If txtComplemento <> "" Then
        Sql = Sql & " and tim_complemento like '" & txtComplemento & "%'"
    End If
    
    If txtICAnt <> "" Then
        Sql = Sql & " and tim_ic_anterior  =  '" & txtICAnt & "'"
    End If
    If txtIM <> "" Then
        Sql = Sql & " and tim_tci_im  =  '" & txtIM & "'"
    End If
    
    If txtContrib <> "" Then
        Sql = Sql & " and tci_nome like '" & txtContrib & "%'"
    End If
    
    If txtIC <> "" Then
        Sql = Sql & " and tim_ic  = '" & UCase(txtIC) & "'"
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
        Sql = Sql & " and TIM_TED_COD_EDIFICIO= '" & cboCond.coluna(0).Valor & "'"
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
    
    If Trim(txtInscMunicipalOcupante) <> "" Then
        Sql = Sql & " and TIM_IM_OCUPANTE  like '%" & txtInscMunicipalOcupante & "%'"
    End If
    
    If cboTipo.ListIndex <> -1 Then
        Sql = Sql & " AND TIM_TIPO_IMOVEL = '" & cboTipo.coluna(1).Valor & "'"
    End If
    '***********************************************************************************************
    SelecaoRpt = "{TAB_BAIRRO.TBA_TMU_COD_MUNICIPIO}= " & Aplicacoes.Codigo_Municipio
    SelecaoRptListagem = " 1 = 1"
    'SelecaoRpt = SelecaoRpt & " and {TAB_LOGRADOURO.tlg_tmu_cod_municipio}="  & Aplicacoes.Codigo_Municipio
    
    'IC
    If Trim(txtIC) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_ic}='" & txtIC & "'"
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.tim_ic}='" & txtIC & "'"
    End If
    If Trim(txtNumero) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_numero}='" & txtNumero & "'"
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.tim_numero}='" & txtNumero & "'"
    End If
    If Trim(cboUsuario) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.TIM_TUS_COD_USUARIO}='" & cboUsuario & "'"
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.TIM_TUS_COD_USUARIO}='" & cboUsuario & "'"
    End If
    'IM
    If cboLoteamento.ListIndex >= 0 Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_loteamento}= " & cboLoteamento.coluna(0).Valor
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.tim_loteamento}= " & cboLoteamento.coluna(0).Valor
    End If
    If cboCond.ListIndex >= 0 Then
        If Bdados.Conexao.FormatoBanco = SQLServer Then
            SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.TIM_TED_COD_EDIFICIO}= " & cboCond.coluna(0).Valor
            SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.TIM_TED_COD_EDIFICIO}= " & cboCond.coluna(0).Valor
        ElseIf Bdados.Conexao.FormatoBanco = oracle Then
            SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.TIM_TED_COD_EDIFICIO}= " & cboCond.coluna(0).Valor
            SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.TIM_TED_COD_EDIFICIO}= " & cboCond.coluna(0).Valor
        End If
    End If
    
    If Trim(txtIM) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_tci_im}='" & txtIM & "'"
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.tim_tci_im}='" & txtIM & "'"
    End If
    'Tipo Logradouro
    If Trim(cboTipoLogr) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_TIPO_LOGR.TTL_NOME}='" & cboTipoLogr & "'"
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.TTL_NOME}='" & cboTipoLogr & "'"
    End If
    'Logradouro
    If Trim(cboLogr) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_LOGRADOURO.tlg_nome}= '" & cboLogr & "'"
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.tlg_nome}= '" & cboLogr & "'"
    End If
    'Bairro
    If Trim(cboBairro) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_BAIRRO.TBA_NOME}= '" & cboBairro & "'"
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.TBA_NOME}= '" & cboBairro & "'"
    End If
    If cboTipo.ListIndex <> -1 Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.TIM_TIPO_IMOVEL}= " & cboTipo.coluna(1).Valor
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.TIM_TIPO_IMOVEL}= " & cboTipo.coluna(1).Valor
    End If
    If Trim(txtBloco) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_bloco}  = '" & txtBloco & "'"
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.tim_bloco}  = '" & txtBloco & "'"
    End If
    If Trim(txtApto) <> "" Then
        SelecaoRpt = SelecaoRpt & "  and {TAB_IMOVEL.tim_apto}  = '" & txtApto & "'"
        SelecaoRptListagem = SelecaoRptListagem & "  and {VIS_IMOVEL.tim_apto}  = '" & txtApto & "'"
    End If
    If txtICAnt <> "" Then
        SelecaoRpt = SelecaoRpt & "  and {TAB_IMOVEL.tim_ic_anterior}  = '" & txtICAnt & "'"
        SelecaoRptListagem = SelecaoRptListagem & "  and {VIS_IMOVEL.tim_ic_anterior}  = '" & txtICAnt & "'"
    End If
    If Trim(txtSala) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.TIM_SALA_LOJA}  = '" & txtSala & "'"
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.TIM_SALA_LOJA}  = '" & txtSala & "'"
    End If
    'Razao Social
    If Trim(txtContrib) <> "" Then
        SelecaoRpt = SelecaoRpt & " and {Tab_Contribuinte.tci_nome} like '" & txtContrib & "*'"
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.tci_nome} like '" & txtContrib & "*'"
    End If
    'Quadra
    If Trim(txtSetor) <> "" Then
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            SelecaoRpt = SelecaoRpt & " and Mid({TAB_IMOVEL.tim_ic_AUXILIAR},3,2)='" & txtSetor & "'"
            SelecaoRptListagem = SelecaoRptListagem & " and Mid({VIS_IMOVEL.tim_ic_AUXILIAR},3,2)='" & txtSetor & "'"
        Else
            SelecaoRpt = SelecaoRpt & " and Mid({TAB_IMOVEL.tim_ic},3,2)='" & txtSetor & "'"
            SelecaoRptListagem = SelecaoRptListagem & " and Mid({VIS_IMOVEL.tim_ic},3,2)='" & txtSetor & "'"
        End If
    End If
    
    If txtQuadra <> "" Then
        SelecaoRpt = SelecaoRpt & " and {TAB_IMOVEL.tim_quadra}  = '" & txtQuadra & "'"
        SelecaoRptListagem = SelecaoRptListagem & " and {VIS_IMOVEL.tim_quadra}  = '" & txtQuadra & "'"
    End If
    'Quadra
    If Trim(txtQuadraInsc) <> "" Then
        If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            SelecaoRpt = SelecaoRpt & " and Mid({TAB_IMOVEL.tim_ic_AUXILIAR},5," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ")='" & txtQuadraInsc & "'"
            SelecaoRptListagem = SelecaoRptListagem & " and Mid({VIS_IMOVEL.tim_ic_AUXILIAR},5," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ")='" & txtQuadraInsc & "'"
        Else
            SelecaoRpt = SelecaoRpt & " and Mid({TAB_IMOVEL.tim_ic},5," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ")='" & txtQuadraInsc & "'"
            SelecaoRptListagem = SelecaoRptListagem & " and Mid({VIS_IMOVEL.tim_ic},5," & Nvl(Temp.PegaParametro(Bdados, "TAMANHO QUADRA"), 3) & ")='" & txtQuadraInsc & "'"
        End If
    End If
    
    '**********************************************************************************************
     Sql = Replace(Sql, "TTL_NOME", "TIPOLOGRADOURO")
     Sql = Replace(Sql, "tlg_nome", "LOGRADOURO")
     
     If Not grid.Preencher(Bdados, Sql) Then
        Util.Avisa "Consulta sem resultados."
     End If
End Sub
Private Sub CmdConsultaContribuinteOcupante_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtInscMunicipalOcupante
End Sub
Private Sub txtInscMunicipalOcupante_LostFocus()
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim cadastro As New VSImposto
    
    If Me.ActiveControl.ToolTipText = "Novo Contribuinte" Or _
        Me.ActiveControl.ToolTipText = "Pesquisa Contribuintes" Then Exit Sub
    If Trim(txtInscMunicipalOcupante) <> "" Then
        If Not Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
            txtInscMunicipalOcupante = cadastro.FormataInscricao(txtInscMunicipalOcupante, InscContrib)
        End If
        Sql = "Select  tci_Nome, tci_logradouro,tci_nome_logradouro, tci_numero, " & _
        " tci_complemento, tci_bairro, tci_cep, tci_cidade,tci_UF,TCI_CGC_CPF,TCI_COD_LOGRADOURO,tci_rg from Tab_Contribuinte where tci_im = '" & txtInscMunicipalOcupante & "'"
        If Bdados.AbreTabela(Sql, Rs) Then
            txtOcupante = "" & Rs(0)  'Rs!tci_Nome
            txtCPFOcupante.Formato = formDocumento
            txtCPFOcupante = "" & Rs!TCI_CGC_CPF
        Else
            Call Util.Informa("Contribuinte não cadastrado.")
            txtInscMunicipalOcupante.Enabled = True
            txtInscMunicipalOcupante.SetFocus
        End If
    End If
    Bdados.FechaTabela Rs
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{tab}"
End Sub

Private Sub cmdImprimir_Click()
    
    Screen.MousePointer = 11
    If grid.ListItems.Count > 0 Then
        With RPT
            If Not .DefinirArquivo(Bdados, App.Path & "\TCIU201.rpt") Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SMTU"), Temp.PegaParametro(Bdados, "SMTUSETOR")
            Else
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            End If
            .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
            .Selecao = SelecaoRpt
            .Titulo = "Ficha Cadastral"
            .Arvore = False
            .Visualizar
            DoEvents
        End With
    End If
    Set RPT = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdImprimirListagem_Click()
  Screen.MousePointer = 11
    If grid.ListItems.Count > 0 Then
        With RPT
            If Not .DefinirArquivo(Bdados, App.Path & "\TCIU201Listagem.rpt") Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SMTU"), Temp.PegaParametro(Bdados, "SMTUSETOR")
            Else
                .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
            End If
            .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
            .Selecao = SelecaoRptListagem
            .Titulo = "Listagem de Imóveis"
            .Arvore = False
            .Visualizar
            DoEvents
        End With
    End If
    Set RPT = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grid.ListItems.Clear
    txtIM.SetFocus
End Sub

Private Sub CmdResumo_Click()
    Screen.MousePointer = 11
    If grid.ListItems.Count > 0 Then
        With RPT
            If Not .DefinirArquivo(Bdados, App.Path & "\TCalculoIPTU.rpt") Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            .Selecao = SelecaoRptListagem
            .Titulo = "Ficha Cadastral"
            .Arvore = False
            .Visualizar
            DoEvents
        End With
    End If
    Set RPT = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    cboLogr.Preencher Bdados, "Select tlg_cod_logradouro,(tlg_nome)   From Tab_Logradouro order by tlg_nome", 1
    
    cboTipoLogr.Preencher Bdados, "Select TTL_COD_TIP_LOGR ,(ttl_nome) From Tab_Tipo_Logr", 1
    cboBairro.Preencher Bdados, "Select DISTINCT(tba_nome),tba_cod_bairro From Tab_Bairro "
    cboLoteamento.Preencher Bdados, "Select TLO_COD_LOTEAMENTO,TLO_DESCRICAO from TAB_LOTEAMENTO ORDER BY TLO_DESCRICAO", 1
    cboCond.Preencher Bdados, "Select TED_COD_EDIFICIO,TED_DESCRICAO from TAB_EDIFICIO ORDER BY TED_DESCRICAO", 1
    
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    
    cboTipo.PreencherGeral Bdados, "TIPO LOTE"
    cboUsuario.Preencher Bdados, "SELECT TUS_COD_USUARIO FROM TAB_USUARIO ORDER BY 1"
    If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
        cmdImprimirListagem.Enabled = True
    Else
        cmdImprimirListagem.Enabled = True
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
                TCIU203.Tag = "I" & grid.SelectedItem
                TCIU203.Show
        Case 0
               Dim ProjObrig As Object
               Set ProjObrig = CreateObject("VSTOBRI.Aplicacoes")
                    
                    Set ProjObrig.Banco = Bdados.Conexao
                    ProjObrig.Usuario = AplicacoesVTFuncoes.Usuario
                    ProjObrig.Codigo_Municipio = AplicacoesVTFuncoes.Codigo_Municipio
                    ProjObrig.municipio = AplicacoesVTFuncoes.municipio
                    
                ProjObrig.Abre_Aplicacao "TOBR401", 0, Cod_sis, Sistema, Desc_Form, "I" & Trim(grid.SelectedItem)
                
                TempContrib = "I" & Trim(grid.SelectedItem)
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
    If AplicacoesVTFuncoes.municipio = "PETROLINA" Then txtICAnt = Left(txtICAnt, 1) & "." & Mid(txtICAnt, 2, 4) & "." & Mid(txtICAnt, 6, 3) & "." & Mid(txtICAnt, 9, 2) & "." & Mid(txtICAnt, 11, 4) & "." & Right(txtICAnt, 4)
End Sub

Private Sub txtICAnterior_Change()
    Exit Sub
    Static aux As Boolean
    If aux Then
        aux = False
        Exit Sub
    End If
    txtICAnterior = Edita.TiraTudo(txtICAnterior)
    If AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        Select Case Len(txtICAnterior)
            Case Is < 5
                If Len(txtICAnterior) > 3 And Not aux Then
                    txtICAnterior = Left(txtICAnterior, 2) & "." & Right(txtICAnterior, 2)
                    aux = True
                Else
                    aux = False
                End If
            Case Is < 8
                If Not aux Then
                    txtICAnterior = Left(txtICAnterior, 2) & "." & Mid(txtICAnterior, 3, 2) & "." & Right(txtICAnterior, 3)
                    aux = True
                Else
                    aux = False
                End If
            Case Is < 12
                If Not aux Then
                    txtICAnterior = Left(txtICAnterior, 2) & "." & Mid(txtICAnterior, 3, 2) & "." & Mid(txtICAnterior, 5, 3) & "." & Right(txtICAnterior, 4)
                    aux = True
                Else
                    aux = False
                End If
            Case Is < 16
                If Not aux Then
                    txtICAnterior = Left(txtICAnterior, 2) & "." & Mid(txtICAnterior, 3, 2) & "." & Mid(txtICAnterior, 5, 3) & "." & Mid(txtICAnterior, 9, 4) & "." & Right(txtICAnterior, 3)
                    aux = True
                Else
                    aux = False
                End If
        End Select
    End If
End Sub

Private Sub txtIcAnterior_KeyPress(KeyAscii As Integer)
    'KeyAscii = Edita.AceitaDig(KeyAscii,Letra)
End Sub

Private Sub txtICAnterior_LostFocus()
    If Len(Trim(txtICAnterior)) <> 14 Then Exit Sub
    txtICAnterior = Edita.TiraTudo(txtICAnterior)
    txtICAnterior = Left(txtICAnterior, 2) & "." & Mid(txtICAnterior, 3, 2) & "." & Mid(txtICAnterior, 5, 3) & "." & Mid(txtICAnterior, 8, 4) & "." & Right(txtICAnterior, 3)
End Sub

Private Sub txtim_KeyPress(KeyAscii As Integer)
KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtIm_LostFocus()
    txtIM = Imposto.FormataInscricao(txtIM, InscContrib)
End Sub

Private Sub txtQuadra_KeyPress(KeyAscii As Integer)
'    KeyAscii = Edita.Maiuscula(KeyAscii)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
    
End Sub


Private Sub txtSetor_KeyPress(KeyAscii As Integer)
'KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub
