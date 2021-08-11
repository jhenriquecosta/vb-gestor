VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TNAV401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota Fiscal Avulsa"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11085
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      TabIndex        =   37
      Top             =   7965
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   1058
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   7440
         TabIndex        =   7
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   8655
         TabIndex        =   8
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9855
         TabIndex        =   9
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
   Begin VTOcx.fraFUTURO fraFUTURO1 
      Height          =   7215
      Left            =   60
      TabIndex        =   11
      Top             =   720
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   12726
      Caption         =   "Nota Fiscal Avulsa"
      Descricao       =   "Eliminação de Notas Fiscais"
      corFaixa        =   16711680
      Icone           =   "TNAV401.frx":0000
      Ocultavel       =   0   'False
      Altura          =   1905
      Begin VTOcx.fraVISUAL fraVISUAL1 
         Height          =   1500
         Left            =   120
         TabIndex        =   19
         Top             =   735
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   2646
         Altura          =   1905
         Caption         =   " Opções de Busca"
         CorTexto        =   16777215
         CorFaixa        =   16711680
         CorFundo        =   -2147483633
         Ocultavel       =   0   'False
         Begin VB.CheckBox chkRetido 
            Caption         =   "Retido"
            Height          =   255
            Left            =   3720
            TabIndex        =   44
            Top             =   1080
            Width           =   975
         End
         Begin VTOcx.cmdVISUAL cmdRelatorio 
            Height          =   375
            Left            =   4920
            TabIndex        =   43
            Top             =   960
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   661
            Caption         =   "&Relatorio"
            Acao            =   4
            CorBorda        =   16711680
            CorFrente       =   0
            CorFundo        =   16777088
         End
         Begin VTOcx.txtVISUAL txtFim 
            Height          =   480
            Left            =   1920
            TabIndex        =   4
            Top             =   840
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   847
            Caption         =   "Data Final"
            Text            =   ""
            Formato         =   0
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtInicio 
            Height          =   480
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   847
            Caption         =   "Data Inicial"
            Text            =   ""
            Formato         =   0
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.cmdVISUAL cmdBuscar 
            Height          =   345
            Left            =   6360
            TabIndex        =   2
            Top             =   960
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   609
            Caption         =   "&Buscar"
            Acao            =   5
            CorBorda        =   16711680
            CorFrente       =   0
            CorFundo        =   16777088
         End
         Begin VTOcx.txtVISUAL txtNomeBusca 
            Height          =   480
            Left            =   1860
            TabIndex        =   1
            Top             =   330
            Width           =   8790
            _ExtentX        =   15505
            _ExtentY        =   847
            Caption         =   "Nome"
            Text            =   ""
            AlinhamentoRotulo=   1
         End
         Begin VTOcx.txtVISUAL txtNumNota 
            Height          =   480
            Left            =   105
            TabIndex        =   0
            Top             =   330
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   847
            Caption         =   "Nº da Nota Fiscal"
            Text            =   ""
            AlinhamentoRotulo=   1
         End
      End
      Begin ActiveTabs.SSActiveTabs tabNota 
         Height          =   3180
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   5609
         _Version        =   131082
         TabCount        =   4
         TabOrientation  =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontSelectedTab {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tabs            =   "TNAV401.frx":08DA
         Images          =   "TNAV401.frx":09C7
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
            Height          =   2760
            Left            =   30
            TabIndex        =   45
            Top             =   30
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   4868
            _Version        =   131082
            TabGuid         =   "TNAV401.frx":2B1D
            Begin VTOcx.cmdVISUAL cmdAlterarServico 
               Height          =   405
               Left            =   8040
               TabIndex        =   46
               Top             =   360
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   714
               Caption         =   "Alterar Fiscal Assinatura"
               Acao            =   1
               CorBorda        =   16711680
               CorFrente       =   0
               CorFundo        =   12648447
            End
            Begin VTOcx.cboVISUAL cboFiscal 
               Height          =   315
               Left            =   240
               TabIndex        =   47
               Tag             =   "C"
               Top             =   360
               Width           =   7650
               _ExtentX        =   13494
               _ExtentY        =   556
               Caption         =   "Fiscal"
               Text            =   ""
               AutoFocaliza    =   0   'False
               Requerido       =   0   'False
            End
         End
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
            Height          =   2760
            Left            =   30
            TabIndex        =   28
            Top             =   30
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   4868
            _Version        =   131082
            TabGuid         =   "TNAV401.frx":2B45
            Begin VTOcx.fraVISUAL fra 
               Height          =   1815
               Index           =   0
               Left            =   105
               TabIndex        =   29
               Top             =   255
               Width           =   10425
               _ExtentX        =   18389
               _ExtentY        =   3201
               Altura          =   1905
               Caption         =   " Informações Gerais"
               CorTexto        =   16777215
               CorFaixa        =   16711680
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Enabled         =   0   'False
               Begin VTOcx.txtVISUAL txtCepDest 
                  Height          =   480
                  Left            =   8925
                  TabIndex        =   36
                  Top             =   1260
                  Width           =   1440
                  _ExtentX        =   2540
                  _ExtentY        =   847
                  Caption         =   "CEP"
                  Text            =   ""
                  Enabled         =   0   'False
                  Formato         =   4
                  AlinhamentoRotulo=   1
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtBairroDest 
                  Height          =   480
                  Left            =   90
                  TabIndex        =   35
                  Top             =   1260
                  Width           =   4620
                  _ExtentX        =   8149
                  _ExtentY        =   847
                  Caption         =   "Bairro"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtEnderecoDest 
                  Height          =   480
                  Left            =   90
                  TabIndex        =   34
                  Top             =   780
                  Width           =   9420
                  _ExtentX        =   16616
                  _ExtentY        =   847
                  Caption         =   "Endereço"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtCidadeDest 
                  Height          =   480
                  Left            =   4740
                  TabIndex        =   33
                  Top             =   1260
                  Width           =   4155
                  _ExtentX        =   7329
                  _ExtentY        =   847
                  Caption         =   "Cidade"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtImCpfCnpjDest 
                  Height          =   480
                  Left            =   120
                  TabIndex        =   32
                  Tag             =   "Ins. Municipal/CPF/CNPJ"
                  Top             =   300
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   847
                  Caption         =   "Ins. Municipal/CPF/CNPJ"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
                  MaxLen          =   20
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtNomeDest 
                  Height          =   480
                  Left            =   3195
                  TabIndex        =   31
                  Tag             =   "Nome"
                  Top             =   300
                  Width           =   7170
                  _ExtentX        =   12647
                  _ExtentY        =   847
                  Caption         =   "Nome"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.cboVISUAL cboUFDest 
                  Height          =   510
                  Left            =   9540
                  TabIndex        =   30
                  Top             =   780
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   900
                  Caption         =   "UF"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
                  Enabled         =   0   'False
               End
            End
         End
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
            Height          =   2760
            Left            =   30
            TabIndex        =   12
            Top             =   30
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   4868
            _Version        =   131082
            TabGuid         =   "TNAV401.frx":2B6D
            Begin VTOcx.txtVISUAL txtIRRF_VALOR 
               Height          =   480
               Left            =   2700
               TabIndex        =   41
               Tag             =   "IRRF"
               Top             =   1800
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   847
               Caption         =   "IRRF (R$)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               ValorPadrao     =   "0"
            End
            Begin VTOcx.txtVISUAL txtIRRF_INDICE 
               Height          =   480
               Left            =   1800
               TabIndex        =   42
               Tag             =   "IRRF"
               Top             =   1800
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   847
               Caption         =   "IRRF (%)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               CorRotulo       =   255
               ValorPadrao     =   "0"
            End
            Begin VTOcx.txtVISUAL txtINSS_Valor 
               Height          =   480
               Left            =   4470
               TabIndex        =   39
               Tag             =   "IRRF"
               Top             =   1800
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   847
               Caption         =   "INSS (R$)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               CorRotulo       =   0
               ValorPadrao     =   "0"
            End
            Begin VTOcx.txtVISUAL txtINSS_Indice 
               Height          =   480
               Left            =   3600
               TabIndex        =   40
               Tag             =   "IRRF"
               Top             =   1800
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   847
               Caption         =   "INSS (%)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               CorRotulo       =   255
               ValorPadrao     =   "0"
            End
            Begin VTOcx.txtVISUAL txtPeriodo 
               Height          =   480
               Left            =   5400
               TabIndex        =   13
               Top             =   1800
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   847
               Caption         =   "Período"
               Text            =   ""
               Enabled         =   0   'False
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtTotalNota 
               Height          =   480
               Left            =   6630
               TabIndex        =   14
               Tag             =   "Total da Nota"
               Top             =   1800
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   847
               Caption         =   "Total da Nota"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtBaseCalc 
               Height          =   480
               Left            =   8085
               TabIndex        =   15
               Tag             =   "Base de Cálculo"
               Top             =   1800
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   847
               Caption         =   "Base de Cálculo"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtISS 
               Height          =   480
               Left            =   9555
               TabIndex        =   16
               Tag             =   "ISS"
               Top             =   1800
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   847
               Caption         =   "ISS Devido"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.grdVISUAL grdItem 
               Height          =   1980
               Left            =   30
               TabIndex        =   17
               Top             =   75
               Width           =   10590
               _ExtentX        =   18680
               _ExtentY        =   3493
               CorBorda        =   16711680
               Caption         =   "Itens"
               CorTitulo       =   16711680
               CorCaption      =   16777215
               CorDica         =   16711680
               OcultarRodape   =   -1  'True
            End
         End
         Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
            Height          =   2760
            Left            =   30
            TabIndex        =   18
            Top             =   30
            Width           =   10665
            _ExtentX        =   18812
            _ExtentY        =   4868
            _Version        =   131082
            TabGuid         =   "TNAV401.frx":2B95
            Begin VTOcx.fraVISUAL fra 
               Height          =   1815
               Index           =   1
               Left            =   90
               TabIndex        =   20
               Top             =   255
               Width           =   10440
               _ExtentX        =   18415
               _ExtentY        =   3201
               Altura          =   1905
               Caption         =   " Informações Gerais"
               CorTexto        =   16777215
               CorFaixa        =   16711680
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Enabled         =   0   'False
               Begin VTOcx.txtVISUAL txtNomeContrib 
                  Height          =   480
                  Left            =   3195
                  TabIndex        =   27
                  Tag             =   "Nome "
                  Top             =   300
                  Width           =   7185
                  _ExtentX        =   12674
                  _ExtentY        =   847
                  Caption         =   "Nome"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtImCpfCnpj 
                  Height          =   480
                  Left            =   105
                  TabIndex        =   26
                  Tag             =   "Municipal/CPF/CNPJ"
                  Top             =   300
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   847
                  Caption         =   "Ins. Municipal/CPF/CNPJ"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
                  MaxLen          =   20
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtCidade 
                  Height          =   480
                  Left            =   4740
                  TabIndex        =   25
                  Top             =   1260
                  Width           =   4200
                  _ExtentX        =   7408
                  _ExtentY        =   847
                  Caption         =   "Cidade"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL TxtCepRem 
                  Height          =   480
                  Left            =   8970
                  TabIndex        =   24
                  Top             =   1260
                  Width           =   1395
                  _ExtentX        =   2461
                  _ExtentY        =   847
                  Caption         =   "CEP"
                  Text            =   ""
                  Enabled         =   0   'False
                  Formato         =   4
                  AlinhamentoRotulo=   1
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtEndereco 
                  Height          =   480
                  Left            =   90
                  TabIndex        =   23
                  Top             =   780
                  Width           =   9420
                  _ExtentX        =   16616
                  _ExtentY        =   847
                  Caption         =   "Endereço"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtBairro 
                  Height          =   480
                  Left            =   90
                  TabIndex        =   22
                  Top             =   1260
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   847
                  Caption         =   "Bairro"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.cboVISUAL cboUFEmi 
                  Height          =   510
                  Left            =   9555
                  TabIndex        =   21
                  Top             =   780
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   900
                  Caption         =   "UF"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
                  Enabled         =   0   'False
               End
            End
         End
      End
      Begin VTOcx.grdVISUAL grdNota 
         Height          =   1500
         Left            =   120
         TabIndex        =   5
         Top             =   2415
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   2646
         CorBorda        =   16711680
         Caption         =   "Notas Fiscais"
         CorTitulo       =   16711680
         CorCaption      =   16777215
         CorDica         =   16711680
         OcultarRodape   =   -1  'True
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   1138
      Icone           =   "TNAV401.frx":2BBD
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "TNAV401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NotaAvulsa As cNotaAvulsa
Dim ItemNota As cItemNotaAvulsa
Dim Contribuinte As cContribuinte
Dim ContribuinteAvulso As cContribuinteAvulso
Dim Aliquota As Single
Dim statusNota As Integer
Dim bcpSqlRelatorio As String
Dim Path As String
Dim notaSeleciona As String

Private Sub cmdAlterarServico_Click()
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim nota As String
    Dim Item As String
    nota = CDbl(grdNota.SelectedItem)
    'Item = grdNotas.SelectedItem.SubItems(5)
    
    'Sql = "update Tab_Item_Nota_Avulsa set tin_descricao_servico='" & txtDescricaoServicoAlteracao.Text & "' where tin_tna_numero_nota=" & nota & " and tin_codigo='" & Item & "'"
    'Bdados.Executa (Sql)
    Sql = "update Tab_Nota_Avulsa set tna_tus_cod_usuario = '" & cboFiscal.Text & "' where tna_numero_nota=" & nota
    Bdados.Executa (Sql)
    
    Informa "Fiscal Alterado com sucesso!"
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdBuscar_Click()
    NotaAvulsa.PreencherGrid grdNota, txtNumNota, txtNomeBusca
    grdNota.Caption = "Nota Referente ao cliente "
    If grdNota.ListItems.Count = 0 Then
        Util.Avisa "Nenhum registro encontrado"
    End If
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo TrataErro
    Dim NumNota As String
    Dim Sql As String
    Dim strNomeUsuario As String, matricula As String
    Dim Rs As VSRecordset
    Dim OBS As String
    If grdNota.SelectedItem Is Nothing Then
        Util.Avisa "Não existe registro selecionado para a impressão"
        Exit Sub
    End If
    
    NumNota = grdNota.SelectedItem.Text
    
'    If Temp.PegaParametro(Bdados, "MODELO NOTA AVULSA") = "2" Then
'        Sql = "SELECT * FROM VIS_NOTA_AVULSA WHERE TNA_NUMERO_NOTA = " & NumNota
'        VisualizarActiveReport AR_NotaAvulsa, Bdados, Sql
'        Exit Sub
'    End If
    
    Sql = "SELECT TUS_NOME,TUS_TSE_MATRICULA FROM TAB_USUARIO WHERE TUS_COD_USUARIO = '" & NotaAvulsa.CodUsuario & "'"
    If Bdados.AbreTabela(Sql, Rs) Then
        strNomeUsuario = "" & Rs(0).Value
        matricula = "" & Rs(1).Value
    End If
    Bdados.FechaTabela Rs
    With RPT
        If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
            Path = App.Path + "\TNotaAvulsaBarraMansa.rpt"
        Else
            Path = App.Path + "\TNotaAvulsa.rpt"
        End If
         
         If Dir(Path) <> "" Then
            OBS = Entrada("Observações...", "Mensagem")
            .DefinirArquivo Bdados, Path
            If UCase(AplicacoesVTFuncoes.municipio) = "BARRA MANSA" Then
                .Formulas "VT_GERENCIA", Temp.PegaParametro(Bdados, "SEMFAZ")
                .Formulas "VT_SETOR", Temp.PegaParametro(Bdados, "GFF")
                .Formulas "VT_IM_EMITENTE", txtImCpfCnpj
                .Formulas "VT_EMITENTECIDADE", txtCidade
                .Formulas "VT_EMITENTE_UF", cboUFEmi
                .Formulas "VT_IM_DESTINO", txtImCpfCnpjDest
            End If
            .Formulas "VT_EmitenteRazao ", txtNomeContrib
            .Formulas "VT_EmitenteEndereco", txtEndereco & " " & txtBairro & " " & TxtCepRem & " " & txtCidade & " " & " " & cboUFEmi
            .Formulas "VT_EmitenteCgcCpfIm ", Pega_Doc(txtImCpfCnpj)
            .Formulas "VTOBS", OBS
            .Formulas "VT_DestinoRazao ", txtNomeDest
            .Formulas "VT_DestinoEndereco ", txtEnderecoDest & " " & txtBairroDest & " " & txtCepDest
            .Formulas "VT_DestinoCgcCpfIm", Pega_Doc(txtImCpfCnpjDest)
            .Formulas "VT_DestinoUf ", cboUFDest
            .Formulas "VT_DestinoMunicipio ", txtCidadeDest
            .Formulas "VT_NumNota ", NumNota
            .Formulas "VT_Destinouf ", cboUFDest
            .Formulas "VT_ValorAliquota ", Format(NotaAvulsa.Aliquota, Const_Monetario)
            .Formulas "VT_ValorIss ", Format(txtISS, Const_Monetario)
            .Formulas "VT_ValorInss", Format(NotaAvulsa.INSS_Valor, Const_Monetario)
            .Formulas "VT_ValorNota ", Format(txtTotalNota, Const_Monetario)
            
            .Formulas "VT_ValorTotalDevido ", Format(CDbl(Nvl(Trim(txtTotalNota), 0)) - CDbl(Nvl(Trim(NotaAvulsa.Material), 0)), Const_Monetario)
            .Formulas "VT_ValorMulta", "0,00'"
            .Formulas "VT_Municipio", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
            .Formulas "VT_DATAEMISSAO", NotaAvulsa.DataEmissao
            .Formulas "VT_CGCMUNIC", CNPJCliente
            .Formulas "VT_ENDERECOMUNIC", UCase(Temp.PegaParametro(Bdados, "ENDERECO CLIENTE") & " - " & UCase(Aplicacoes.municipio))
            .Formulas "VT_MATERIAL", Format(Nvl(Trim(NotaAvulsa.Material), 0), Const_Monetario)
            
            'BCP
            .Formulas "VT_IRRF", Format(Nvl(NotaAvulsa.IRRF, 0), Const_Monetario)
            .Formulas "VT_IRRF_INDICE", Format(Nvl(NotaAvulsa.IRRF_INDICE, 0), Const_Monetario)
            .Formulas "VT_ValorINSS", Format(Nvl(NotaAvulsa.INSS_Valor, 0), Const_Monetario)
            .Formulas "VT_INSS_INDICE", Format(Nvl(NotaAvulsa.INSS_Indice, 0), Const_Monetario)
            .Formulas "VT_Funcionario", strNomeUsuario
            .Formulas "matriculaExpeditor", matricula
            .Formulas "valorNotaExtenso", "NOTA: " & Extenso(Format(txtTotalNota, Const_Monetario))
            .Formulas "valorISSExtenso", "ISS: " & Extenso(Format(txtISS, Const_Monetario))
            If statusNota = 3 Then 'CANCELADA
                .Formulas "CANCELADA", "NOTA FISCAL CANCELADA "
            End If
            
            '
            
            
            .Titulo = "Nota Fiscal Avulsa"
            .Arvore = False
            .Visualizar
         Else
            Util.Mensagem "Relatório não encontrado." & vbCrLf & Path
         End If
    End With
    
    Set RPT = Nothing
    
    Exit Sub
    
TrataErro:
    Util.Erro Err.Description
    Exit Sub
    Resume
End Sub
Public Function Extenso(ByVal nValor As Double)

        If nValor <= 0 Or nValor > 9999999.99 Then
            Extenso = "ZERO"
            Exit Function
        End If

        'Declara as variáveis da função
        Dim nContador, nTamanho As Integer
        Dim cValor, cParte, cFinal As String
        Dim aGrupo(4), aTexto(4) As String

        'Define matrizes com extensos parciais
        Dim aUnid(19) As String
        aUnid(1) = "UM ": aUnid(2) = "DOIS ": aUnid(3) = "TRES "
        aUnid(4) = "QUATRO ": aUnid(5) = "CINCO ": aUnid(6) = "SEIS "
        aUnid(7) = "SETE ": aUnid(8) = "OITO ": aUnid(9) = "NOVE "
        aUnid(10) = "DEZ ": aUnid(11) = "ONZE ": aUnid(12) = "DOZE "
        aUnid(13) = "TREZE ": aUnid(14) = "QUATORZE ": aUnid(15) = "QUINZE "
        aUnid(16) = "DEZESSEIS ": aUnid(17) = "DEZESSETE ": aUnid(18) = "DEZOITO "
        aUnid(19) = "DEZENOVE "

        Dim aDezena(9) As String
        aDezena(1) = "DEZ ": aDezena(2) = "VINTE ": aDezena(3) = "TRINTA "
        aDezena(4) = "QUARENTA ": aDezena(5) = "CINQUENTA "
        aDezena(6) = "SESSENTA ": aDezena(7) = "SETENTA ": aDezena(8) = "OITENTA "
        aDezena(9) = "NOVENTA "

        Dim aCentena(9) As String
        aCentena(1) = "CENTO ": aCentena(2) = "DUZENTOS "
        aCentena(3) = "TREZENTOS ": aCentena(4) = "QUATROCENTOS "
        aCentena(5) = "QUINHENTOS ": aCentena(6) = "SEISCENTOS "
        aCentena(7) = "SETECENTOS ": aCentena(8) = "OITOCENTOS "
        aCentena(9) = "NOVECENTOS "

        'Divide o valor em vários grupos
        cValor = Format(nValor, "0000000000.00")
        aGrupo(1) = Mid$(cValor, 2, 3)
        aGrupo(2) = Mid$(cValor, 5, 3)
        aGrupo(3) = Mid$(cValor, 8, 3)
        aGrupo(4) = "0" + Mid$(cValor, 12, 2)

        'Processa cada grupo
        For nContador = 1 To 4
            cParte = aGrupo(nContador)
            nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
            If nTamanho = 3 Then
                If Right$(cParte, 2) <> "00" Then
                    aTexto(nContador) = aTexto(nContador) + aCentena(Left(cParte, 1)) + "E "
                    nTamanho = 2
                Else
                    aTexto(nContador) = aTexto(nContador) + IIf(Left$(cParte, 1) = "1", "CEM ", aCentena(Left(cParte, 1)))
                End If
            End If
            If nTamanho = 2 Then
                If Val(Right(cParte, 2)) < 20 Then
                    aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 2))
                Else
                    aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
                    If Right$(cParte, 1) <> "0" Then
                        aTexto(nContador) = aTexto(nContador) + "E "
                        nTamanho = 1
                    End If
                End If
            End If
            If nTamanho = 1 Then
                aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 1))
            End If
        Next

        'Gera o formato final do texto
        If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
            cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "CENTAVO", "CENTAVOS")
        Else
            cFinal = ""
            cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, "MILHÕES ", "MILHÃO "), "")
            If Val(aGrupo(2) + aGrupo(3)) = 0 Then
                cFinal = cFinal + "DE "
            Else
                cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "MIL ", "")
            End If
            cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "REAL ", "REAIS ")
            cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "E " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, "CENTAVO", "CENTAVOS"), "")
        End If

        Extenso = cFinal & "****************"
        
    End Function

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    grdItem.ListItems.Clear
    grdNota.ListItems.Clear
    txtNumNota.SetFocus
    tabNota.Tabs(1).Selected = True
End Sub

Private Sub cmdRelatorio_Click()
    'bcpSqlRelatorio = NotaAvulsa.RetornaSQL(txtInicio, txtFim, "Tab_Nota_Avulsa.tna_numero_nota")
    Dim Rs As VSRecordset
    If Bdados.AbreTabela(bcpSqlRelatorio, Rs) Then
     With RPT
        Path = App.Path + "\TNotasAvulsas.rpt"
        .DefinirArquivo Bdados, Path
        If chkRetido.Value = 1 Then
            If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1179 Then 'CODO
                .Selecao = "{VIS_NOTAS_FISCAIS_AVULSAS.tna_tca_identidade_dest}='11015604-02' and {VIS_NOTAS_FISCAIS_AVULSAS.tna_data_emissao} >= #" & retornarData(txtInicio) & "# and {VIS_NOTAS_FISCAIS_AVULSAS.tna_data_emissao} <= #" & retornarData(txtFim) & "#"
                .Formulas "VT_PERIODO", "NOTAS EMITIDAS ENTRE:  " & txtInicio & " até: " & txtFim & " - COM IMPOSTO RETIDO NA FONTE"
            Else
                .Selecao = "{VIS_NOTAS_FISCAIS_AVULSAS.tna_data_emissao} >= #" & retornarData(txtInicio) & "# and {VIS_NOTAS_FISCAIS_AVULSAS.tna_data_emissao} <= #" & retornarData(txtFim) & "#"
                .Formulas "VT_PERIODO", "NOTAS EMITIDAS ENTRE:  " & txtInicio & " até: " & txtFim
            End If
        Else
            .Selecao = "{VIS_NOTAS_FISCAIS_AVULSAS.tna_data_emissao} >= #" & retornarData(txtInicio) & "# and {VIS_NOTAS_FISCAIS_AVULSAS.tna_data_emissao} <= #" & retornarData(txtFim) & "#"
            .Formulas "VT_PERIODO", "NOTAS EMITIDAS ENTRE:  " & txtInicio & " até: " & txtFim
        End If
        .Titulo = "Nota Fiscal Avulsa"
        .Arvore = False
        .Visualizar
        
    End With
    End If
   
    Set RPT = Nothing
End Sub
Private Function retornarData(data As String) As String
    Dim nd As String
    data = Replace(data, "/", "")
    nd = Right(data, 4) & "-" & Mid(data, 3, 2) & "-" & Left(data, 2)
    retornarData = nd
End Function


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set NotaAvulsa = New cNotaAvulsa
    Set ItemNota = New cItemNotaAvulsa
    Set Contribuinte = New cContribuinte
    Set ContribuinteAvulso = New cContribuinteAvulso
    
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
        
    cboUFDest.PreencherGeral Bdados, "UF"
    cboUFEmi.PreencherGeral Bdados, "UF"
    txtInicio = Format(Now, "DD/MM/YYYY")
    txtFim = Format(Now, "DD/MM/YYYY")
    Dim Rs As VSRecordset
    If Bdados.AbreTabela("SELECT TUS_COD_USUARIO FROM TAB_USUARIO ORDER BY TUS_COD_USUARIO", Rs) Then
    Do While Not Rs.EOF
        cboFiscal.AddItem Rs(0)
        Rs.MoveNext
    Loop
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set NotaAvulsa = Nothing
    Set ItemNota = Nothing
    Set Contribuinte = Nothing
    Set ContribuinteAvulso = Nothing
End Sub



'Private Sub grdNota_Click()
 '   Dim RetNome As String
  '  If grdNota.ListItems.Count >= 1 Then
   '         Contribuinte.BuscarContribuinte grdNota.SelectedItem.SubItems(1), RetNome
    '        grdNota.Caption = "Nota de número " & grdNota.SelectedItem & " Referente ao contribuinte - " & RetNome
    'End If

'End Sub

Private Sub grdNota_Click()
    'click
    Dim RetNome As String
    If grdNota.ListItems.Count >= 1 Then
            Contribuinte.BuscarContribuinte grdNota.SelectedItem.SubItems(1), RetNome
            grdNota.Caption = "Nota de número " & grdNota.SelectedItem & " Referente ao contribuinte - " & RetNome
    End If
    '
    'antes era dbl
    If grdNota.SelectedItem Is Nothing Then Exit Sub
    With NotaAvulsa
        If .Buscar(grdNota.SelectedItem) Then
            'BCP
            notaSeleciona = .NumNota
            cboFiscal = .CodUsuario
            '
            statusNota = .statusNota
            txtImCpfCnpj = .IdentidadeRemetente
            txtImCpfCnpj_LostFocus
            txtImCpfCnpjDest = .IdentidadeDestinatario
            txtImCpfCnpjDest_lostfocus
            txtPeriodo = .Periodo
            txtTotalNota = .ValorNota
            txtBaseCalc = txtTotalNota - .Material
            txtISS = .ValorImposto
            txtINSS_Indice = Format(.INSS_Indice, Const_Monetario)
            txtINSS_Valor = Format(.INSS_Valor, Const_Monetario)
            
            txtIRRF_INDICE = Format(.IRRF_INDICE, Const_Monetario)
            txtIRRF_VALOR = Format(.IRRF, Const_Monetario)
            'Aliquota = grdNota.SelectedItem.ListSubItems(7).Text
            If ItemNota.PreencherGrid(grdItem, grdNota.SelectedItem) = False Then
                Util.Avisa "Nota Fiscal sem Itens."
            End If
        Else
            Util.Avisa "Nota não encontrada."
        End If
    End With
    If notaSeleciona > 0 Or notaSeleciona <> "" Then
        'NotaAvulsa.PreencherGridComServico grdNotas, notaSeleciona, txtNomeContrib
        'grdNotas.Caption = "Nota Referente ao cliente "
        'If grdNotas.ListItems.Count = 0 Then
         '   Util.Avisa "Nenhum registro encontrado"
        'End If
    End If
    '
End Sub

Private Sub grdNota_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 40 Or KeyCode = 38 Then
            grdNota_Click
    End If

End Sub

'Private Sub grdNotas_Click()
 '   txtDescricaoServicoAlteracao.Text = grdNotas.SelectedItem.SubItems(4)
'End Sub

'Private Sub grdNota_ItemClick(ByVal Item As MSComctlLib.IListItem)

'End Sub

Private Sub txtImCpfCnpj_LostFocus()
  Dim NomeContrib As String, TipoLogrContr As String, LogrContr As String, NumeroContr As String, CompContri As String, _
          BairroContr As String, CepContr As String, MunicContr As String, UFContr As String, DocumentoContr As String
    LimpaRementente
    If Trim(txtImCpfCnpj) = "" Then Exit Sub
    If Len(txtImCpfCnpj) = 10 Then
        txtImCpfCnpj.Formato = formDoisDigitos
        txtImCpfCnpj.AgruparValores = False
    ElseIf Len(txtImCpfCnpj) = 11 And IsNumeric(txtImCpfCnpj) Then
        txtImCpfCnpj.Formato = formCPF
    ElseIf Len(txtImCpfCnpj) = 14 And IsNumeric(txtImCpfCnpj) Then
        txtImCpfCnpj.Formato = formCGC
    End If
    If Trim(txtImCpfCnpj) = "" Then
        txtImCpfCnpj.AgruparValores = True
        txtImCpfCnpj.Formato = formNenhum
        Exit Sub
    End If
    If Contribuinte.BuscarContribuinte(txtImCpfCnpj, NomeContrib, TipoLogrContr, LogrContr, NumeroContr, CompContri, _
        BairroContr, CepContr, MunicContr, UFContr, DocumentoContr) Then
        txtNomeContrib = NomeContrib
        txtEndereco = TipoLogrContr & "  " & LogrContr & "  " & NumeroContr & "  " & CompContri
        txtBairro = BairroContr
        TxtCepRem = CepContr
        txtCidade = MunicContr
        cboUFEmi.SetarLinha UFContr, 0
    Else
        With ContribuinteAvulso
            If .Buscar(txtImCpfCnpj) Then
                txtNomeContrib = .Nome
                txtEndereco = .Endereco
                txtBairro = .Bairro
                TxtCepRem = .Cep
                txtCidade = .Cidade
                cboUFEmi.SetarLinha .Uf, 0
            End If
        End With
    End If
    txtImCpfCnpj.Mascara = ""
    txtImCpfCnpj.Formato = formNenhum
End Sub

Private Sub txtImCpfCnpjDest_lostfocus()
    Dim NomeDest As String, TipoLogrDest As String, LogrDest As String, NumeroDest As String, CompDest As String, _
          BairroDest As String, CepDest As String, MunicDest As String, UFDest As String, DocumentoDest As String
    LimpaDestino
    If Trim(txtImCpfCnpjDest) = "" Then Exit Sub
    If Len(txtImCpfCnpjDest) = 10 Then
        txtImCpfCnpjDest.Formato = formDoisDigitos
        txtImCpfCnpjDest.AgruparValores = False
    ElseIf Len(txtImCpfCnpjDest) = 11 And IsNumeric(txtImCpfCnpjDest) Then
        txtImCpfCnpjDest.Formato = formCPF
    ElseIf Len(txtImCpfCnpjDest) = 14 And IsNumeric(txtImCpfCnpjDest) Then
        txtImCpfCnpjDest.Formato = formCGC
    End If
    If Trim(txtImCpfCnpjDest) = "" Then
        txtImCpfCnpjDest.AgruparValores = True
        txtImCpfCnpjDest.Formato = formNenhum
        Exit Sub
    End If
    If Contribuinte.BuscarContribuinte(txtImCpfCnpjDest, NomeDest, TipoLogrDest, LogrDest, NumeroDest, CompDest, _
            BairroDest, CepDest, MunicDest, UFDest, DocumentoDest) Then
            txtNomeDest = NomeDest
            txtEnderecoDest = TipoLogrDest & "  " & LogrDest & "  " & NumeroDest & "  " & CompDest
            txtBairroDest = BairroDest
            txtCidadeDest = MunicDest
            txtCepDest = CepDest
            cboUFDest.SetarLinha UFDest, 0
    Else
        With ContribuinteAvulso
            If .Buscar(txtImCpfCnpjDest) Then
                txtNomeDest = .Nome
                txtEnderecoDest = .Endereco
                txtBairroDest = .Bairro
                txtCidadeDest = .Cidade
                txtCepDest = .Cep
                cboUFDest.SetarLinha .Uf, 0
            End If
        End With
    End If
    txtImCpfCnpjDest.Mascara = ""
    txtImCpfCnpjDest.Formato = formNenhum
End Sub

Sub LimpaRementente()
    txtNomeContrib = ""
    txtEndereco = ""
    txtBairro = ""
    TxtCepRem = ""
    cboUFEmi = ""
    txtCidade = ""
End Sub

Sub LimpaDestino()
    txtNomeDest = ""
    txtEnderecoDest = ""
    txtBairroDest = ""
    txtCepDest = ""
    cboUFDest = ""
    txtCidadeDest = ""
End Sub
Private Function Pega_Doc(Im As String) As String
    On Error Resume Next
    Dim Sql As String
    Sql = "SELECT tci_cgc_cpf FROM tab_contribuinte  where tci_im = " & Bdados.Converte(Im, tctexto)
    If Bdados.AbreTabela(Sql) Then
        Pega_Doc = Bdados.Tabela(0)
    Else
        Pega_Doc = Im
    End If
End Function

