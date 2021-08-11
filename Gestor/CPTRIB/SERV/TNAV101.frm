VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TNAV101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   690
      Left            =   75
      TabIndex        =   52
      Top             =   690
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   1217
      Altura          =   1905
      Caption         =   " Data de Emissão"
      CorTexto        =   16777215
      CorFaixa        =   16711680
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtData 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Tag             =   "Data"
         Top             =   345
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         Caption         =   ""
         Text            =   ""
         Formato         =   0
      End
   End
   Begin ActiveTabs.SSActiveTabs tabNota 
      DragIcon        =   "TNAV101.frx":0000
      Height          =   4920
      Left            =   75
      TabIndex        =   10
      Top             =   1440
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   8678
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
      Enabled         =   0   'False
      Tabs            =   "TNAV101.frx":030A
      Images          =   "TNAV101.frx":03F5
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel4 
         Height          =   4500
         Left            =   30
         TabIndex        =   56
         Top             =   30
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7938
         _Version        =   131082
         TabGuid         =   "TNAV101.frx":1588
         Begin VTOcx.fraFUTURO fraFUTURO4 
            DragIcon        =   "TNAV101.frx":15B0
            Height          =   4455
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   7858
            Caption         =   "Consulta das Notas"
            Descricao       =   "Exibe as Notas Fiscais do Prestador"
            corFaixa        =   16711680
            Icone           =   "TNAV101.frx":18BA
            Ocultavel       =   0   'False
            Altura          =   1905
            Begin VB.TextBox txtDescricaoServicoAlteracao 
               Appearance      =   0  'Flat
               Height          =   765
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   61
               Top             =   3600
               Width           =   9075
            End
            Begin VTOcx.cmdVISUAL cmdAlterarServico 
               Height          =   330
               Left            =   9240
               TabIndex        =   60
               Top             =   4050
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   582
               Caption         =   "Alterar"
               Acao            =   1
               CorBorda        =   16711680
               CorFrente       =   0
               CorFundo        =   16777088
            End
            Begin VTOcx.cmdVISUAL cmdListarNotas 
               Height          =   330
               Left            =   9240
               TabIndex        =   59
               Top             =   3610
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   582
               Caption         =   "Buscar"
               Acao            =   5
               CorBorda        =   16711680
               CorFrente       =   0
               CorFundo        =   16777088
            End
            Begin VTOcx.grdVISUAL grdNotas 
               Height          =   3120
               Left            =   120
               TabIndex        =   58
               Top             =   720
               Width           =   10440
               _ExtentX        =   18415
               _ExtentY        =   5503
               CorBorda        =   16711680
               Caption         =   "Notas"
               CorTitulo       =   16711680
               CorCaption      =   16777215
               CorDica         =   16711680
               OcultarRodape   =   -1  'True
            End
         End
      End
      Begin VB.PictureBox PicBarra 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   765
         TabIndex        =   54
         Top             =   -615
         Visible         =   0   'False
         Width           =   795
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel3 
         Height          =   4500
         Left            =   30
         TabIndex        =   48
         Top             =   30
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7938
         _Version        =   131082
         TabGuid         =   "TNAV101.frx":18D6
         Begin VTOcx.fraFUTURO fraFUTURO3 
            Height          =   4365
            Left            =   0
            TabIndex        =   51
            Top             =   30
            Width           =   10605
            _ExtentX        =   18706
            _ExtentY        =   7699
            Caption         =   "Composição Nota Fiscal"
            Descricao       =   "Dados gerais e itens correspondente a nota"
            corFaixa        =   16711680
            Icone           =   "TNAV101.frx":18FE
            Ocultavel       =   0   'False
            Altura          =   1905
            Begin VTOcx.txtVISUAL txtIRRF_INDICE 
               Height          =   480
               Left            =   2120
               TabIndex        =   31
               Tag             =   "IRRF"
               Top             =   3810
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   847
               Caption         =   "IRRF (%)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
               CorRotulo       =   255
            End
            Begin VTOcx.txtVISUAL txtTotalNota 
               Height          =   480
               Left            =   5625
               TabIndex        =   35
               TabStop         =   0   'False
               Tag             =   "Total da Nota"
               Top             =   3810
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   847
               Caption         =   "Total da Nota"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtINSS_Valor 
               Height          =   480
               Left            =   4740
               TabIndex        =   34
               Tag             =   "IRRF"
               Top             =   3810
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
               Left            =   3870
               TabIndex        =   33
               Tag             =   "IRRF"
               Top             =   3810
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
            Begin VTOcx.txtVISUAL txtVence 
               Height          =   480
               Left            =   9150
               TabIndex        =   38
               TabStop         =   0   'False
               Tag             =   "ISS"
               Top             =   3810
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   847
               Caption         =   "Vencto Imposto"
               Text            =   ""
               Formato         =   0
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VB.TextBox txtDescServico 
               Appearance      =   0  'Flat
               Height          =   765
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   20
               Top             =   960
               Width           =   10395
            End
            Begin VTOcx.txtVISUAL txtAlqt 
               Height          =   495
               Left            =   150
               TabIndex        =   22
               Top             =   1800
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   873
               Caption         =   "Aliquota(%)"
               Text            =   ""
               Formato         =   5
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cmdVISUAL cmdRemover 
               Height          =   330
               Left            =   8730
               TabIndex        =   27
               Top             =   1950
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   582
               Caption         =   "Remover Item"
               Acao            =   2
               CorBorda        =   16711680
               CorFrente       =   0
               CorFundo        =   16777088
            End
            Begin VTOcx.txtVISUAL txtISS 
               Height          =   480
               Left            =   8100
               TabIndex        =   37
               TabStop         =   0   'False
               Tag             =   "ISS"
               Top             =   3810
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   847
               Caption         =   "ISS Devido"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtBaseCalc 
               Height          =   480
               Left            =   6930
               TabIndex        =   36
               TabStop         =   0   'False
               Tag             =   "Base de Cálculo"
               Top             =   3810
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   847
               Caption         =   "Base Cálculo"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtIRRF 
               Height          =   480
               Left            =   3000
               TabIndex        =   32
               Tag             =   "IRRF"
               Top             =   3810
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   847
               Caption         =   "IRRF (R$)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtValorMaterial 
               Height          =   480
               Left            =   900
               TabIndex        =   30
               Tag             =   "Valor Material"
               Top             =   3810
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   847
               Caption         =   "Valor Material"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtPeriodo 
               Height          =   480
               Left            =   120
               TabIndex        =   29
               TabStop         =   0   'False
               Tag             =   "Período"
               Top             =   3810
               Width           =   800
               _ExtentX        =   1402
               _ExtentY        =   847
               Caption         =   "Período"
               Text            =   ""
               Restricao       =   2
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.cmdVISUAL cmdInclui 
               Height          =   330
               Left            =   6900
               TabIndex        =   26
               Top             =   1950
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   582
               Caption         =   "Adicionar Item"
               Acao            =   1
               CorBorda        =   16711680
               CorFrente       =   0
               CorFundo        =   16777088
            End
            Begin VTOcx.txtVISUAL txtValorTotal 
               Height          =   480
               Left            =   5145
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   1800
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   847
               Caption         =   "Valor Total(R$)"
               Text            =   ""
               Enabled         =   0   'False
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtValorUnitario 
               Height          =   480
               Left            =   3315
               TabIndex        =   24
               Top             =   1800
               Width           =   1725
               _ExtentX        =   3043
               _ExtentY        =   847
               Caption         =   "Valor Unitário(R$)"
               Text            =   ""
               Formato         =   5
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.txtVISUAL txtQnt 
               Height          =   480
               Left            =   1770
               TabIndex        =   23
               Top             =   1800
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   847
               Caption         =   "Quantidade"
               Text            =   ""
               Restricao       =   3
               AlinhamentoRotulo=   1
            End
            Begin VTOcx.grdVISUAL grdNota 
               Height          =   1680
               Left            =   105
               TabIndex        =   28
               Top             =   2340
               Width           =   10440
               _ExtentX        =   18415
               _ExtentY        =   2963
               CorBorda        =   16711680
               Caption         =   "Itens"
               CorTitulo       =   16711680
               CorCaption      =   16777215
               CorDica         =   16711680
               OcultarRodape   =   -1  'True
            End
            Begin VB.Label Label1 
               Caption         =   "Descrição do Serviço"
               Height          =   285
               Left            =   180
               TabIndex        =   55
               Top             =   750
               Width           =   2535
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   4500
         Left            =   30
         TabIndex        =   47
         Top             =   30
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7938
         _Version        =   131082
         TabGuid         =   "TNAV101.frx":21D8
         Begin VTOcx.fraFUTURO fraFUTURO2 
            Height          =   5025
            Left            =   30
            TabIndex        =   49
            Top             =   15
            Width           =   10605
            _ExtentX        =   18706
            _ExtentY        =   8864
            Caption         =   "Definição de Tomador"
            Descricao       =   "Informações do tomador, cadastra novo caso não exista"
            corFaixa        =   16711680
            Icone           =   "TNAV101.frx":2200
            Ocultavel       =   0   'False
            Altura          =   1905
            Begin VTOcx.grdVISUAL grdDest 
               Height          =   1605
               Left            =   90
               TabIndex        =   21
               Top             =   3105
               Width           =   10455
               _ExtentX        =   18441
               _ExtentY        =   2831
               CorBorda        =   16711680
               Caption         =   "Destinatários"
               CorTitulo       =   16711680
               CorCaption      =   16777215
               CorDica         =   16711680
               OcultarRodape   =   -1  'True
            End
            Begin VTOcx.fraVISUAL fra 
               Height          =   2340
               Index           =   0
               Left            =   105
               TabIndex        =   50
               Top             =   720
               Width           =   10425
               _ExtentX        =   18389
               _ExtentY        =   4128
               Altura          =   1905
               Caption         =   " Informações Gerais"
               CorTexto        =   16777215
               CorFaixa        =   16711680
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.cboVISUAL cboUFDest 
                  Height          =   510
                  Left            =   9540
                  TabIndex        =   15
                  Top             =   780
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   900
                  Caption         =   "UF"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
               End
               Begin VTOcx.cboVISUAL cboAtiviEcon 
                  Height          =   510
                  Left            =   105
                  TabIndex        =   19
                  Top             =   1740
                  Width           =   4620
                  _ExtentX        =   8149
                  _ExtentY        =   900
                  Caption         =   "Atividade Econômica"
                  Text            =   ""
                  AutoFocaliza    =   0   'False
                  Alinhamento     =   1
                  Enabled         =   0   'False
               End
               Begin VTOcx.txtVISUAL txtNomeDest 
                  Height          =   480
                  Left            =   3615
                  TabIndex        =   13
                  Tag             =   "Nome"
                  Top             =   300
                  Width           =   6750
                  _ExtentX        =   11906
                  _ExtentY        =   847
                  Caption         =   "Nome"
                  Text            =   ""
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.cmdVISUAL cmBuscaDest 
                  Height          =   330
                  Left            =   3195
                  TabIndex        =   12
                  Top             =   450
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   582
                  Caption         =   ""
                  Acao            =   5
                  CorBorda        =   8421504
                  CorFrente       =   16384
               End
               Begin VTOcx.txtVISUAL txtImCpfCnpjDest 
                  Height          =   480
                  Left            =   120
                  TabIndex        =   11
                  Tag             =   "Ins. Municipal/CPF/CNPJ"
                  Top             =   300
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   847
                  Caption         =   "Ins. Municipal/CPF/CNPJ"
                  Text            =   ""
                  AlinhamentoRotulo=   1
                  MaxLen          =   20
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.txtVISUAL txtCidadeDest 
                  Height          =   480
                  Left            =   4740
                  TabIndex        =   17
                  Top             =   1260
                  Width           =   4155
                  _ExtentX        =   7329
                  _ExtentY        =   847
                  Caption         =   "Cidade"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtEnderecoDest 
                  Height          =   480
                  Left            =   90
                  TabIndex        =   14
                  Top             =   780
                  Width           =   9420
                  _ExtentX        =   16616
                  _ExtentY        =   847
                  Caption         =   "Endereço"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtBairroDest 
                  Height          =   480
                  Left            =   90
                  TabIndex        =   16
                  Top             =   1260
                  Width           =   4620
                  _ExtentX        =   8149
                  _ExtentY        =   847
                  Caption         =   "Bairro"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtCepDest 
                  Height          =   480
                  Left            =   8925
                  TabIndex        =   18
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
            End
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   4500
         Left            =   30
         TabIndex        =   44
         Top             =   30
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7938
         _Version        =   131082
         TabGuid         =   "TNAV101.frx":2ADA
         Begin VTOcx.fraFUTURO fraFUTURO1 
            Height          =   5040
            Left            =   30
            TabIndex        =   45
            Top             =   0
            Width           =   10620
            _ExtentX        =   18733
            _ExtentY        =   8890
            Caption         =   "Definição de Prestador"
            Descricao       =   "Traz informações do prestador, cadastra novo caso não exista"
            corFaixa        =   16711680
            Icone           =   "TNAV101.frx":2B02
            Ocultavel       =   0   'False
            Altura          =   1905
            Begin VTOcx.fraVISUAL fra 
               Height          =   1815
               Index           =   1
               Left            =   105
               TabIndex        =   46
               Top             =   705
               Width           =   10440
               _ExtentX        =   18415
               _ExtentY        =   3201
               Altura          =   1905
               Caption         =   " Informações Gerais"
               CorTexto        =   16777215
               CorFaixa        =   16711680
               CorFundo        =   -2147483633
               Ocultavel       =   0   'False
               Begin VTOcx.cboVISUAL cboUFEmi 
                  Height          =   510
                  Left            =   9555
                  TabIndex        =   5
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
               Begin VTOcx.txtVISUAL txtBairro 
                  Height          =   480
                  Left            =   90
                  TabIndex        =   6
                  Top             =   1260
                  Width           =   4635
                  _ExtentX        =   8176
                  _ExtentY        =   847
                  Caption         =   "Bairro"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtEndereco 
                  Height          =   480
                  Left            =   90
                  TabIndex        =   4
                  Top             =   780
                  Width           =   9420
                  _ExtentX        =   16616
                  _ExtentY        =   847
                  Caption         =   "Endereço"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL TxtCepRem 
                  Height          =   480
                  Left            =   8970
                  TabIndex        =   8
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
               Begin VTOcx.txtVISUAL txtCidade 
                  Height          =   480
                  Left            =   4740
                  TabIndex        =   7
                  Top             =   1260
                  Width           =   4200
                  _ExtentX        =   7408
                  _ExtentY        =   847
                  Caption         =   "Cidade"
                  Text            =   ""
                  Enabled         =   0   'False
                  AlinhamentoRotulo=   1
               End
               Begin VTOcx.txtVISUAL txtImCpfCnpj 
                  Height          =   480
                  Left            =   105
                  TabIndex        =   1
                  Tag             =   "Municipal/CPF/CNPJ"
                  Top             =   300
                  Width           =   3060
                  _ExtentX        =   5398
                  _ExtentY        =   847
                  Caption         =   "Ins. Municipal/CPF/CNPJ"
                  Text            =   ""
                  AlinhamentoRotulo=   1
                  MaxLen          =   20
                  RetirarMascara  =   0   'False
               End
               Begin VTOcx.cmdVISUAL cmdBuscarContr 
                  Height          =   330
                  Left            =   3180
                  TabIndex        =   2
                  Top             =   450
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   582
                  Caption         =   ""
                  Acao            =   5
                  CorBorda        =   8421504
                  CorFrente       =   16384
               End
               Begin VTOcx.txtVISUAL txtNomeContrib 
                  Height          =   480
                  Left            =   3615
                  TabIndex        =   3
                  Tag             =   "Nome "
                  Top             =   300
                  Width           =   6750
                  _ExtentX        =   11906
                  _ExtentY        =   847
                  Caption         =   "Nome"
                  Text            =   ""
                  AlinhamentoRotulo=   1
               End
            End
            Begin VTOcx.grdVISUAL grdEmit 
               Height          =   2010
               Left            =   90
               TabIndex        =   9
               Top             =   2565
               Width           =   10470
               _ExtentX        =   18468
               _ExtentY        =   3545
               CorBorda        =   16711680
               Caption         =   "Emitente"
               CorTitulo       =   16711680
               CorCaption      =   16777215
               CorDica         =   16711680
               OcultarRodape   =   -1  'True
            End
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   43
      Top             =   6390
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   979
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   8430
         TabIndex        =   40
         Top             =   105
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   7215
         TabIndex        =   39
         Top             =   105
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9660
         TabIndex        =   41
         Top             =   105
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
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   1138
      Icone           =   "TNAV101.frx":33DC
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "TNAV101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NovoRemetente As Boolean
Dim NovoDestino   As Boolean
Dim Aliquota      As Double
Dim Imposto       As New VSImposto
Dim Contribuinte As cContribuinte
Dim ContribuinteAvulso As cContribuinteAvulso
Dim NotaAvulsa As cNotaAvulsa
Dim ItemNota As cItemNotaAvulsa
Dim InscricaoRealContribuinte As String
Dim notaEmitida As Boolean
Private Sub CalculaTotais()
    Dim i As Integer
    Dim TotalItens As Double
    
    For i = 1 To grdNota.ListItems.Count
        TotalItens = TotalItens + CDbl(grdNota.ListItems(i).SubItems(4))
    Next
    txtTotalNota = Format(TotalItens, Const_Monetario)
    CalculaIss
End Sub

Sub CalculaIss()
    On Error Resume Next
    Dim i As Integer
    Dim Base As Double
    Dim Iss As Double
    If txtTotalNota = "" Then
        txtBaseCalc = ""
        txtISS = ""
        Exit Sub
    End If
    
    Base = Nvl(Trim(txtTotalNota), 0) - Nvl(Trim(txtValorMaterial), 0)
    Base = IIf(Base > 0, Base, 0)
    For i = 1 To grdNota.ListItems.Count
        Iss = Iss + (CDbl(grdNota.ListItems(i).SubItems(5)) / 100) * CDbl(grdNota.ListItems(i).SubItems(4))
    Next
    txtISS = Format(Iss, Const_Monetario)
    txtBaseCalc = Format(Base, Const_Monetario)
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
Sub HabilitaRemetente(Status As Boolean)
    txtNomeContrib.Enabled = Status
    txtEndereco.Enabled = Status
    txtBairro.Enabled = Status
    TxtCepRem.Enabled = Status
    cboUFEmi.Enabled = Status
    txtCidade.Enabled = Status
End Sub

Sub HabilitaDestino(Status As Boolean)
    txtNomeDest.Enabled = Status
    txtEnderecoDest.Enabled = Status
    txtBairroDest.Enabled = Status
    txtCepDest.Enabled = Status
    cboUFDest.Enabled = Status
    txtCidadeDest.Enabled = Status
End Sub

Private Sub cmdAlterarServico_Click()
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim nota As Variant
    Dim Item As Variant
    nota = grdNotas.SelectedItem
    Item = grdNotas.SelectedItem.SubItems(5)
    
    Sql = "update Tab_Item_Nota_Avulsa set tin_descricao_servico='" & txtDescricaoServicoAlteracao.Text & "' where tin_tna_numero_nota=" & nota & " and tin_codigo=" & Item
    Bdados.Executa (Sql)
    Informa "Descriçao serviço alterada"
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub
Private Sub MontargrdNota()
    grdNota.ColumnHeaders.Clear
    grdNota.ColumnHeaders.Add , , "Item", 5000
    grdNota.ColumnHeaders.Add , , "Qnt", 700
    grdNota.ColumnHeaders.Add , , "Unidade", 0
    grdNota.ColumnHeaders.Add , , "Valor Unit", 1100
    grdNota.ColumnHeaders.Add , , "Valor Total", 1200
    grdNota.ColumnHeaders.Add , , "Aliquota", 900
    grdNota.ColumnHeaders.Add , , "ISS", 900
    
End Sub



Private Sub cmBuscaDest_Click()
    
 AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtImCpfCnpjDest

End Sub

Private Sub cmdBuscarContr_Click()
       
     AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtImCpfCnpj
    
End Sub

Private Sub cmdInclui_Click()
    On Error Resume Next
    Dim ItmX As Object
    If Trim(txtQnt) = "" Then
        Avisa "Informe Quantidade."
        txtQnt.SetFocus
        Exit Sub
    End If
    If Trim(txtAlqt) = "" Then
        Avisa "Informe Aliquota."
        txtAlqt.SetFocus
        Exit Sub
    End If
    If Trim(txtDescServico) = "" Then
        Avisa "Informe Descricao do Item."
        txtDescServico.SetFocus
        Exit Sub
    End If
    If Trim(txtValorUnitario) = "" Then
        Avisa "Informe Valor Unitário."
        txtValorUnitario.SetFocus
        Exit Sub
    End If
    
    Set ItmX = grdNota.ListItems.Add(, , txtDescServico)
    ItmX.SubItems(1) = txtQnt
    'ItmX.SubItems(2) = txtUnd
    ItmX.SubItems(3) = txtValorUnitario
    ItmX.SubItems(4) = txtValorTotal
    ItmX.SubItems(5) = txtAlqt
    ItmX.SubItems(6) = (CDbl(txtAlqt) / 100) * CDbl(txtValorTotal)
    CalculaTotais
    txtDescServico = ""
    'txtQnt = ""
    'txtUnd = ""
    txtValorUnitario = ""
    txtValorTotal = ""
    'txtAlqt = ""
    txtDescServico.SetFocus
    txtVence = Format(Now, "DD/MM/YYYY")
    txtVence = Imposto.BuscaDataVencimento(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN)), Right(txtData, 4))
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    tabNota.Tabs(1).Selected = True
    grdNota.ListItems.Clear
    grdDest.ListItems.Clear
    grdEmit.ListItems.Clear
    MontargrdNota
    'If Screen.ActiveForm.ActiveControl.Name = Me.cmdLimpar.Name Then
    txtData.Text = Format(Date, "dd/MM/yyyy")
    txtData.SetFocus
    'End If
    tabNota.Enabled = False
    txtImCpfCnpj.Enabled = True
    txtImCpfCnpjDest.Enabled = True
    txtNomeContrib.Enabled = True
    txtNomeDest.Enabled = True
    txtValorMaterial = 0
    txtIRRF = 0
    txtIRRF_INDICE = 0
    InscricaoRealContribuinte = ""
    txtVence = Imposto.BuscaDataVencimento(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN)), Right(txtData, 4))
    notaEmitida = False
End Sub

Private Sub cmdListarNotas_Click()
    NotaAvulsa.PreencherGridComServico grdNotas, "", txtNomeContrib
    grdNotas.Caption = "Nota Referente ao cliente "
    If grdNotas.ListItems.Count = 0 Then
        Util.Avisa "Nenhum registro encontrado"
    End If
End Sub

Private Sub cmdRemover_Click()
    grdNota_dblclick
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo TRATA
Dim NumNota As String
Dim Item As Variant
Dim Path As String
Dim Sql As String
Dim strMensagem As String
Dim strNomeUsuario As String
Dim Rs As VSRecordset
Dim Observacao As String
Dim Cobranca As New VSCobranca
Dim Vencimento As String
Dim Obrig As New Obrigacao
Dim CodPagamento As String

'BCP
Dim OBS As String
Dim taxa As Double
taxa = 0
Dim imprimeDam As Boolean
imprimeDam = True
'


    If Not Edita.CriticaCampos(Me) Then Exit Sub
    
    If notaEmitida = True Then
        Util.Informa "Esta nota já foi gerada, click em Limpar para emitir uma nova."
        Exit Sub
    End If
    
    
    If Trim(txtImCpfCnpj) = Trim(txtImCpfCnpjDest) Then
        Util.Informa "O Prestador não pode ser igual ao Tomador do serviço."
        tabNota.Tabs(1).Selected = True
        txtImCpfCnpj.SetFocus
        Exit Sub
    End If
    
    If NovoRemetente Then
        With ContribuinteAvulso
            .Identidade = txtImCpfCnpj
            .Nome = txtNomeContrib
            .Endereco = txtEndereco
            .Uf = cboUFEmi
            .Bairro = txtBairro
            .Cidade = txtCidade
            .Cep = TxtCepRem
            .CodUsuario = AplicacoesVTFuncoes.Usuario
            .Salvar
        End With
    End If
    If NovoDestino Then
        With ContribuinteAvulso
            .Identidade = txtImCpfCnpjDest
            .Nome = txtNomeDest
            .Endereco = txtEnderecoDest
            .Uf = cboUFEmi
            .Bairro = txtBairroDest
            .Cidade = txtCidadeDest
            .Cep = TxtCepRem
            .CodUsuario = AplicacoesVTFuncoes.Usuario
            .Salvar
        End With
    End If
    
    NumNota = Imposto.GeraNumNota(1, 65)
    If Trim(txtVence) <> "" Then
        Vencimento = txtVence
    Else
        Vencimento = Imposto.BuscaDataVencimento(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN)), Right(txtData, 4))
    End If
    If Trim(InscricaoRealContribuinte) = "" Then InscricaoRealContribuinte = Const_ImAvulso
    'BCP
    If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1179 And txtImCpfCnpjDest = "11015604-02" Then 'CODO
        imprimeDam = False
        CodPagamento = 0
    Else
        imprimeDam = True
        CodPagamento = Obrig.CriaObrigacao(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN)), Edita.TiraPic(txtPeriodo, "/"), _
                Edita.TiraPic(txtPeriodo, "/"), InscricaoRealContribuinte, CDbl(txtISS), etsCreditoOriginalAberto, , , , , , , , , NumNota)
    End If
    'FIM BCP
    With NotaAvulsa
        .ObsCancelamento = ""
        .NumNota = NumNota
        .statusNota = 0
        .DataEmissao = Bdados.Converte(Format(txtData, "dd/mm/yyyy"), TCDataHora)
        .DataRecepcao = .DataEmissao
        .IdentidadeRemetente = txtImCpfCnpj
        .IdentidadeDestinatario = txtImCpfCnpjDest
        .ValorNota = txtTotalNota
        .ValorImposto = txtISS
        .CodUsuario = AplicacoesVTFuncoes.Usuario
        '.Aliquota = Aliquota
        'txtPeriodo = "10/2003"
        .CodPagamento = CodPagamento
        .Periodo = txtPeriodo
        .Material = txtValorMaterial
        .IRRF = txtIRRF
        .INSS_Indice = txtINSS_Indice
        .INSS_Valor = txtINSS_Valor
        .IRRF_INDICE = txtIRRF_INDICE
        .Salvar
    End With
    Dim codigoItem As Integer
    codigoItem = 1
    For Each Item In grdNota.ListItems
        With ItemNota
            .Codigo = codigoItem
            .NumNota = NumNota
            .DescricaoServico = Item.Text
            .Quantidade = Item.SubItems(1)
            .Unidade = Item.SubItems(2)
            .Valor = Item.SubItems(3)
            .Aliquota = Item.SubItems(5)
            .Salvar
            codigoItem = codigoItem + 1
        End With
    Next
    
    'Util.Informa "Transação Finalizada"
    strMensagem = "Transação Finalizada. Nº de Nota Fiscal Gerado: " & NumNota & "."
    Avisa strMensagem
    notaEmitida = True
    If CDbl(txtISS) = 0 Then
        Screen.MousePointer = 0
        Edita.LimpaCampos Me
        Exit Sub
    End If
    
    If Trim(CodPagamento) <> "" Then
        Observacao = "DOC. REF. AO TRIBUTO ISSQN DA NOTA FISCAL " & NumNota
        If Temp.PegaParametro(Bdados, "MUNICIPIO") = 1179 Then 'CODO
            taxa = Temp.PegaParametro(Bdados, "TXTDAM")
        End If
    End If
    'bcp
    If imprimeDam = True Then
        Cobranca.imprimeDam RPT, CodPagamento, InscricaoRealContribuinte, txtNomeContrib, Pega_Doc(InscricaoRealContribuinte), txtEndereco, "", "", _
        Imposto.BuscaCodImposto("ISSQN"), "ISSQN", "", txtPeriodo, 0, 1, Vencimento, txtBaseCalc, txtISS, _
        0, 0, taxa, "", "", Observacao, PicBarra, "", "", txtValorMaterial, , , , , , , , , etdNormal
    
    Else
        Avisa "O Valor do Tributo desta NOTA FISCAL será RETIDO NA FONTE"
    End If
    'fim bcp
    Set RPT = Nothing
    Screen.MousePointer = 0
    
    Exit Sub
    If Util.Confirma(strMensagem & " Deseja imprimir a nota?") Then
'        If Temp.PegaParametro(Bdados, "MODELO NOTA AVULSA") = "2" Then
'            Sql = "SELECT * FROM VIS_NOTA_AVULSA WHERE TNA_NUMERO_NOTA = " & NumNota
'            cmdLimpar_Click
'            Screen.MousePointer = 0
'            VisualizarActiveReport AR_NotaAvulsa, Bdados, Sql
'            Exit Sub
'        Else
        
            Sql = "SELECT TUS_NOME FROM TAB_USUARIO WHERE TUS_COD_USUARIO = '" & Aplicacoes.Usuario & "'"
            If Bdados.AbreTabela(Sql, Rs) Then
                strNomeUsuario = Rs(0).Value
            End If
            Bdados.FechaTabela Rs
            
            With RPT
                 Path = App.Path + "\TNotaAvulsa.rpt"
                 OBS = Entrada("Observações...", "Mensagem")
                 If Dir(Path) <> "" Then
                    .DefinirArquivo Bdados, Path
                    .Formulas "VT_EmitenteRazao ", txtNomeContrib
                    .Formulas "VT_EmitenteEndereco", txtEndereco & " " & txtBairro & " " & TxtCepRem & " " & txtCidade & " " & " " & cboUFEmi
                    .Formulas "VT_EmitenteCgcCpfIm ", txtImCpfCnpj
                    .Formulas "VT_DestinoRazao ", txtNomeDest
                    .Formulas "VT_DestinoEndereco ", txtEnderecoDest & " " & txtBairroDest & " " & txtCepDest
                    .Formulas "VT_DestinoCgcCpfIm", txtImCpfCnpjDest
                    .Formulas "VT_DestinoUf ", cboUFDest
                    .Formulas "VT_DestinoMunicipio ", txtCidadeDest
                    .Formulas "VT_NumNota ", NumNota
                    .Formulas "VTOBS", OBS
                    .Formulas "VT_Destinouf ", cboUFDest
                    .Formulas "VT_ValorAliquota ", Format(Aliquota, Const_Monetario)
                    .Formulas "VT_ValorIss ", Format(txtISS, Const_Monetario)
                    .Formulas "VT_ValorINSS", Format(txtINSS_Valor, Const_Monetario)
                    .Formulas "VT_ValorNota ", Format(txtTotalNota, Const_Monetario)
                    .Formulas "VT_ValorTotalDevido ", Format(CDbl(Nvl(Trim(txtTotalNota), 0)) - CDbl(Nvl(Trim(txtValorMaterial), 0)), Const_Monetario)
                    .Formulas "VT_ValorMulta", "0,00'"
                    .Formulas "VT_Municipio", UCase(Temp.PegaParametro(Bdados, "CLIENTE"))
                    .Formulas "VT_DATAEMISSAO", txtData
                    
                    Dim CgcPref As String
                    CgcPref = UCase(Temp.PegaParametro(Bdados, "CGC CLIENTE"))
                    CgcPref = Edita.TiraTudo(CgcPref)
                    .Formulas "VT_CGCMUNIC", Left(CgcPref, 2) & "." & Mid(CgcPref, 3, 3) & "." & Mid(CgcPref, 6, 3) & "/" & Mid(CgcPref, 9, 4) & "-" & Right(CgcPref, 2)
                    .Formulas "VT_ENDERECOMUNIC", UCase(Temp.PegaParametro(Bdados, "ENDERECO CLIENTE") & " - " & UCase(Aplicacoes.municipio))
                    .Formulas "VT_MATERIAL", Format(Nvl(Trim(txtValorMaterial), 0), Const_Monetario)
                    .Formulas "VT_IRRF", Format(Nvl(txtIRRF, 0), Const_Monetario)
                    .Formulas "VT_IRRF_INDICE", Format(Nvl(txtIRRF_INDICE, 0), Const_Monetario)
                    .Formulas "VT_ValorINSS", Format(Nvl(txtINSS_Valor, 0), Const_Monetario)
                    .Formulas "VT_INSS_INDICE", Format(Nvl(txtINSS_Indice, 0), Const_Monetario)
                
                    .Formulas "VT_Funcionario", strNomeUsuario
                    .Titulo = "Nota Fiscal Avulsa"
                    .Arvore = False
                    cmdLimpar_Click
                    .Visualizar
                 Else
                    Util.Mensagem "Relatório não encontrado." & vbCrLf & Path
                    cmdLimpar_Click
                 End If
            End With
        'End If
    Else
        cmdLimpar_Click
    End If
        
    Set RPT = Nothing
    
    'Fim
     
     Screen.MousePointer = 0
     
     Exit Sub
TRATA:
     Erro Err.Description
     Exit Sub
     Resume
End Sub



Private Sub Form_Load()
    notaEmitida = False
    cabVISUAL.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    Set Contribuinte = New cContribuinte
    Set ContribuinteAvulso = New cContribuinteAvulso
    Set NotaAvulsa = New cNotaAvulsa
    Set ItemNota = New cItemNotaAvulsa
    
    HabilitaDestino False
    HabilitaRemetente False
    
    cboUFDest.PreencherGeral Bdados, "UF"
    cboUFEmi.PreencherGeral Bdados, "UF"
    
    MontargrdNota
    
    Aliquota = NotaAvulsa.BuscaAliquota
     
    NotaAvulsa.PreencheCboAtividade cboAtiviEcon
    txtData.Text = Format(Date, "dd/MM/yyyy")
    txtValorMaterial = 0
    txtIRRF = 0
    txtIRRF_INDICE = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Contribuinte = Nothing
    Set ContribuinteAvulso = Nothing
    Set NotaAvulsa = Nothing
    Set ItemNota = Nothing
End Sub

Private Sub grdDest_DblClick()
    On Error Resume Next
    txtImCpfCnpjDest = grdDest.SelectedItem
    txtImCpfCnpjDest_lostfocus
End Sub

Private Sub grdEmit_DblClick()
    On Error Resume Next
    txtImCpfCnpj = grdEmit.SelectedItem
    txtImCpfCnpj_LostFocus
End Sub

Private Sub grdNota_dblclick()
    If grdNota.SelectedItem Is Nothing Then Exit Sub
    txtTotalNota = CDbl(txtTotalNota) - CDbl(grdNota.SelectedItem.SubItems(4))
    txtDescServico = grdNota.SelectedItem
    txtQnt = grdNota.SelectedItem.SubItems(1)
    'txtUnd = grdNota.SelectedItem.SubItems(2)
    txtValorUnitario = grdNota.SelectedItem.SubItems(3)
    txtValorTotal = grdNota.SelectedItem.SubItems(4)
    txtAlqt = grdNota.SelectedItem.SubItems(5)
     
    grdNota.ListItems.Remove (grdNota.SelectedItem.Index)
  
    CalculaTotais
End Sub

Private Sub grdNotas_dblClick()
    Dim Rs As VSRecordset
    Dim Sql As String
    Dim nota As Variant
    Dim Item As Variant
    
    nota = grdNotas.SelectedItem
    
    Item = grdNotas.SelectedItem.SubItems(5)
    Sql = "SELECT tin_descricao_servico from Tab_Item_Nota_Avulsa where tin_tna_numero_nota=" & nota & " and tin_codigo=" & Item
    If Bdados.AbreTabela(Sql, Rs) Then
        Rs.MoveFirst
        txtDescServico.Text = Rs("tin_descricao_servico")
        txtDescricaoServicoAlteracao.Text = Rs("tin_descricao_servico")
    Else
         txtDescServico.Text = "Nota " & nota & " não encontrada "
    End If
End Sub

Private Sub txtData_LostFocus()
    Dim Fator As Byte
    If Trim(txtData) <> "" Then
        tabNota.Enabled = True
        txtPeriodo = Month(txtData)
        Fator = IIf(txtPeriodo > 12, 1, 0)
        txtAlqt = 5
        txtQnt = 1
        
        If Fator = 1 Then txtPeriodo = Fator
        
        txtINSS_Indice = 0
        txtINSS_Valor = 0
      
        txtPeriodo = Format(txtPeriodo, "00") & "/" & Format(Year(txtData) + Fator, "0000")
        txtVence = Imposto.BuscaDataVencimento(Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ISSQN)), Right(txtData, 4))
        txtImCpfCnpj.Enabled = True
        tabNota.SelectedTab = 1
        If Screen.ActiveForm Is Me Then
            txtImCpfCnpj.SetFocus
        End If
    Else
        tabNota.Enabled = False
    End If
End Sub

Private Sub txtDescServico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtImCpfCnpj_LostFocus()
 Dim NomeContrib As String, TipoLogrContr As String, LogrContr As String, NumeroContr As String, CompContri As String, _
          BairroContr As String, CepContr As String, MunicContr As String, UFContr As String, DocumentoContr As String
    LimpaRementente
    
    If Trim(txtImCpfCnpj) = "" Then Exit Sub
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(txtImCpfCnpj) = 10 Then
            txtImCpfCnpj = Imposto.FormataInscricao(txtImCpfCnpj, InscContrib)
            txtImCpfCnpj.AgruparValores = False
        ElseIf Len(txtImCpfCnpj) = 11 And IsNumeric(txtImCpfCnpj) Then
            txtImCpfCnpj.Formato = formCPF
        ElseIf Len(txtImCpfCnpj) = 14 And IsNumeric(txtImCpfCnpj) Then
            txtImCpfCnpj.Formato = formCGC
        End If
    End If
    If Trim(txtImCpfCnpj) = "" Then
        txtImCpfCnpj.AgruparValores = True
        txtImCpfCnpj.Formato = formNenhum
        Exit Sub
    End If
    InscricaoRealContribuinte = ""
    If Contribuinte.BuscarContribuinte(txtImCpfCnpj, NomeContrib, TipoLogrContr, LogrContr, NumeroContr, CompContri, _
        BairroContr, CepContr, MunicContr, UFContr, DocumentoContr, , InscricaoRealContribuinte) Then
        txtNomeContrib = NomeContrib
        txtEndereco = TipoLogrContr & "  " & LogrContr & "  " & NumeroContr & "  " & CompContri
        txtBairro = BairroContr
        TxtCepRem = CepContr
        txtCidade = MunicContr
        cboUFEmi.SetarLinha UFContr, 0
        NovoRemetente = False
        HabilitaRemetente False
    Else
        With ContribuinteAvulso
            If .Buscar(txtImCpfCnpj) Then
                txtNomeContrib = .Nome
                txtEndereco = .Endereco
                txtBairro = .Bairro
                TxtCepRem = .Cep
                txtCidade = .Cidade
                cboUFEmi.SetarLinha .Uf, 0
                NovoRemetente = False
                HabilitaRemetente False
            Else
                Util.Avisa "Contribuinte nao encontrado."
                NovoRemetente = True
                HabilitaRemetente True
            End If
        End With
    End If
     If txtNomeContrib = "" And txtNomeContrib.Enabled = True Then
        Util.Avisa "Informe Nome/Razão Social Prestador."
        txtNomeContrib.SetFocus
        Exit Sub
    End If
    
    If Not Contribuinte.PreencherGrid(grdEmit, txtNomeContrib) Then
        txtNomeContrib.Enabled = True
        txtNomeContrib.SetFocus
    End If
    
    txtImCpfCnpj.Mascara = ""
    txtImCpfCnpj.Formato = formNenhum
End Sub

Private Sub txtImCpfCnpjDest_lostfocus()
    Dim NomeDest As String, TipoLogrDest As String, LogrDest As String, NumeroDest As String, CompDest As String, _
          BairroDest As String, CepDest As String, MunicDest As String, UFDest As String, DocumentoDest As String
    LimpaDestino
    If Trim(txtImCpfCnpjDest) = "" Then Exit Sub
    If Not AplicacoesVTFuncoes.municipio = "PETROLINA" Then
        If Len(txtImCpfCnpjDest) = 10 Then
            txtImCpfCnpjDest = Imposto.FormataInscricao(txtImCpfCnpjDest, InscContrib)
            txtImCpfCnpjDest.AgruparValores = False
        ElseIf Len(txtImCpfCnpjDest) = 11 And IsNumeric(txtImCpfCnpjDest) Then
            txtImCpfCnpjDest.Formato = formCPF
        ElseIf Len(txtImCpfCnpjDest) = 14 And IsNumeric(txtImCpfCnpjDest) Then
            txtImCpfCnpjDest.Formato = formCGC
        End If
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
            NovoDestino = False
            HabilitaDestino False
    Else
        With ContribuinteAvulso
            If .Buscar(txtImCpfCnpjDest) Then
                txtNomeDest = .Nome
                txtEnderecoDest = .Endereco
                txtBairroDest = .Bairro
                txtCidadeDest = .Cidade
                txtCepDest = .Cep
                cboUFDest.SetarLinha .Uf, 0
                NovoDestino = False
                HabilitaDestino False
            Else
                Util.Avisa "Contrivuinte nao encontrado."
                NovoDestino = True
                HabilitaDestino True
            End If
        End With
    End If
    
    If txtNomeDest = "" And txtNomeDest.Enabled = True Then
        Util.Avisa "Informe Nome/Razão Social do Tomador."
        txtNomeDest.SetFocus
        Exit Sub
    End If
    
    If Not Contribuinte.PreencherGrid(grdDest, txtNomeDest) Then
        txtNomeDest.Enabled = True
        txtNomeDest.SetFocus
    End If

    txtImCpfCnpjDest.Mascara = ""
    txtImCpfCnpjDest.Formato = formNenhum
End Sub
 
Private Sub txtINSS_Indice_LostFocus()
    
    If txtINSS_Indice = "" Then txtINSS_Indice = 0
    If txtTotalNota = "" Then txtTotalNota = 0
    
    Dim mInd As Currency
    mInd = CCur(txtINSS_Indice)
    If mInd > 11 Then
       Util.Informa "Valor de Indice Errado!"
       txtINSS_Indice = 0
       txtINSS_Valor = 0
       Exit Sub
    End If

    Dim mTot As Currency
    Dim mVlr As Currency
    
    mVlr = CCur(txtTotalNota)
    txtINSS_Valor = Format((mVlr * mInd) / 100, Const_Monetario)
     
End Sub
Private Sub txtIRRF_INDICE_LostFocus()
    
    If txtIRRF_INDICE = "" Then txtIRRF_INDICE = 0
    If txtTotalNota = "" Then txtTotalNota = 0
    
    Dim mInd As Currency
    mInd = CCur(txtIRRF_INDICE)
    If mInd > 11 Then
       Util.Informa "Valor de Indice Errado!"
       txtIRRF_INDICE = 0
       txtIRRF_INDICE = 0
       Exit Sub
    End If

    Dim mTot As Currency
    Dim mVlr As Currency
    
    mVlr = CCur(txtTotalNota)
    txtIRRF = Format((mVlr * mInd) / 100, Const_Monetario)
     
End Sub
Private Sub txtPeriodo_LostFocus()
    If Trim(txtPeriodo) <> "" Then
        If IsNumeric(txtPeriodo) Then
            txtPeriodo = Left(txtPeriodo, 2) & "/" & Right(txtPeriodo, 4)
        End If
    Else
        Informa "Informe o período de referência da nota."
    End If
End Sub

Private Sub txtValorMaterial_Change()
    Dim Carga As Double
    
End Sub

Private Sub txtValorMaterial_LostFocus()
    Dim ValorDeduzido As Double
    If CDbl(Nvl(Trim(txtValorMaterial), 0)) > 0 Then
        If grdNota.ListItems.Count > 1 Then
            Avisa "Faca uma nota exclusiva para o servico com fornecimento de material."
            txtValorMaterial.SetFocus
        Else
            ValorDeduzido = CDbl(txtTotalNota) - CDbl(Nvl(Trim(txtValorMaterial), 0))
            txtBaseCalc = Format(ValorDeduzido, Const_Monetario)
            txtISS = ValorDeduzido * (CDbl(grdNota.SelectedItem.SubItems(5)) / 100)
        End If
    End If
End Sub

Private Sub txtValorUnitario_change()
    If Trim(txtQnt) = "" Or Trim(txtValorUnitario) = "" Then Exit Sub
    If Not IsNumeric(txtValorUnitario) Then Exit Sub
    txtValorTotal = CDbl(txtQnt) * CDbl(txtValorUnitario)
End Sub

Private Function Pega_Doc(Im As String) As String
    Dim Sql As String
    Sql = "SELECT tci_cgc_cpf FROM tab_contribuinte  where tci_im = " & Bdados.Converte(Im, tctexto)
    If Bdados.AbreTabela(Sql) Then
        Pega_Doc = "" & Bdados.Tabela(0)
    End If
End Function
