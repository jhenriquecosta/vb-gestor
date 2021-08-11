VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#3.0#0"; "VTControles.ocx"
Begin VB.Form TCAF103 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "trib"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin ActiveTabs.SSActiveTabs tabAforamento 
      Height          =   6525
      Left            =   45
      TabIndex        =   36
      Top             =   675
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   11509
      _Version        =   131082
      TabCount        =   2
      TabOrientation  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
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
      TagVariant      =   ""
      Tabs            =   "TCAF103.frx":0000
      Images          =   "TCAF103.frx":0082
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   6105
         Left            =   30
         TabIndex        =   37
         Top             =   30
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   10769
         _Version        =   131082
         TabGuid         =   "TCAF103.frx":0F0E
         Begin VTOcx.cmdVISUAL cmdBuscar 
            Height          =   345
            Left            =   5205
            TabIndex        =   7
            Top             =   1575
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   609
            Caption         =   "&Buscar"
            Acao            =   5
            CorBorda        =   8421504
            CorFrente       =   16384
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtConsultaIC 
            Height          =   300
            Left            =   1065
            TabIndex        =   0
            Top             =   90
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   529
            Caption         =   "IC"
            Text            =   ""
            Restricao       =   2
            CorFundo        =   -2147483633
            MaxLen          =   15
         End
         Begin VTOcx.grdVISUAL grdAforamento 
            Height          =   4065
            Left            =   105
            TabIndex        =   38
            Top             =   1980
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   4339
            CorBorda        =   32768
            CorFundo        =   -2147483633
            Caption         =   "Aforamentos"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
         End
         Begin VTOcx.txtVISUAL txtConsultaIMCedente 
            Height          =   300
            Left            =   285
            TabIndex        =   1
            Top             =   465
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   529
            Caption         =   "IM Cedente"
            Text            =   ""
            Formato         =   8
            Restricao       =   2
            CorFundo        =   -2147483633
            MaxLen          =   12
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtConsultaLivro 
            Height          =   300
            Left            =   855
            TabIndex        =   3
            Top             =   1185
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   529
            Caption         =   "Livro"
            Text            =   ""
            Restricao       =   2
            CorFundo        =   -2147483633
            MaxLen          =   15
            Mascara         =   "0000"
         End
         Begin VTOcx.txtVISUAL txtConsultaIMAdquirente 
            Height          =   300
            Left            =   75
            TabIndex        =   2
            Top             =   825
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   529
            Caption         =   "IM Adquirente"
            Text            =   ""
            Formato         =   8
            Restricao       =   2
            CorFundo        =   -2147483633
            MaxLen          =   12
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtConsultaDataInicio 
            Height          =   300
            Left            =   885
            TabIndex        =   5
            Top             =   1560
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   529
            Caption         =   "Data"
            Text            =   ""
            Formato         =   0
            Restricao       =   2
            CorFundo        =   -2147483633
            MaxLen          =   15
         End
         Begin VTOcx.txtVISUAL txtConsultaDataFim 
            Height          =   300
            Left            =   2670
            TabIndex        =   6
            Tag             =   " "
            Top             =   1560
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            Caption         =   "a"
            Text            =   ""
            Formato         =   0
            Restricao       =   2
            CorFundo        =   -2147483633
            MaxLen          =   15
         End
         Begin VTOcx.txtVISUAL txtConsultaLogr 
            Height          =   300
            Left            =   3105
            TabIndex        =   50
            Top             =   90
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   529
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
            CorFundo        =   -2147483633
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtConsultaCedente 
            Height          =   300
            Left            =   2670
            TabIndex        =   51
            Top             =   465
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   529
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtConsultaAdquirente 
            Height          =   300
            Left            =   2670
            TabIndex        =   52
            Top             =   825
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   529
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cmdVISUAL cmdLimpar 
            Height          =   345
            Left            =   6240
            TabIndex        =   8
            Top             =   1575
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   609
            Caption         =   "&Limpar"
            Acao            =   6
            CorBorda        =   8421504
            CorFrente       =   16384
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtConsultaFicha 
            Height          =   300
            Left            =   2730
            TabIndex        =   4
            Top             =   1185
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            Caption         =   "Ficha"
            Text            =   ""
            Restricao       =   2
            CorFundo        =   -2147483633
            MaxLen          =   15
            Mascara         =   "0000"
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   6105
         Left            =   -99969
         TabIndex        =   39
         Top             =   30
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   10769
         _Version        =   131082
         TabGuid         =   "TCAF103.frx":0F36
         Begin VTOcx.txtVISUAL txtLimFrente 
            Height          =   300
            Left            =   2610
            TabIndex        =   12
            Tag             =   "Limite Frente"
            Top             =   660
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   529
            Caption         =   ""
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtData 
            Height          =   300
            Left            =   660
            TabIndex        =   20
            TabStop         =   0   'False
            Tag             =   "Data"
            Top             =   2655
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Caption         =   "Data"
            Text            =   ""
            Formato         =   0
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtFolha 
            Height          =   300
            Left            =   5760
            TabIndex        =   41
            TabStop         =   0   'False
            Tag             =   " "
            Top             =   2985
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            Caption         =   "Folha"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            Mascara         =   "0000"
         End
         Begin VTOcx.txtVISUAL txtLivro 
            Height          =   300
            Left            =   4185
            TabIndex        =   42
            TabStop         =   0   'False
            Tag             =   " "
            Top             =   2985
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            Caption         =   "Livro"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            Mascara         =   "0000"
         End
         Begin VTOcx.txtVISUAL txtFicha 
            Height          =   300
            Left            =   2610
            TabIndex        =   22
            Tag             =   "Ficha"
            Top             =   2985
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            Caption         =   "Ficha"
            Text            =   ""
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            Mascara         =   "0000"
         End
         Begin VTOcx.txtVISUAL txtOrdem 
            Height          =   300
            Left            =   480
            TabIndex        =   43
            TabStop         =   0   'False
            Tag             =   " "
            Top             =   2985
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   529
            Caption         =   "Ordem"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            Mascara         =   "00000"
         End
         Begin VTOcx.txtVISUAL txtIC 
            Height          =   300
            Left            =   675
            TabIndex        =   10
            Tag             =   "IC"
            Top             =   330
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   529
            Caption         =   "IC"
            Text            =   ""
            CorFundo        =   -2147483633
            MaxLen          =   15
         End
         Begin VTOcx.txtVISUAL txtCedente 
            Height          =   300
            Left            =   2700
            TabIndex        =   46
            Top             =   3675
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   529
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtIMCedente 
            Height          =   300
            Left            =   285
            TabIndex        =   21
            Tag             =   "IM Cedente"
            Top             =   3675
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   529
            Caption         =   "Contribuinte"
            Text            =   ""
            Formato         =   8
            Restricao       =   2
            CorFundo        =   -2147483633
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtAdquirente 
            Height          =   300
            Left            =   2700
            TabIndex        =   48
            Top             =   4260
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   529
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtIMAdquirente 
            Height          =   300
            Left            =   285
            TabIndex        =   23
            Tag             =   "IM Adquirente"
            Top             =   4260
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   529
            Caption         =   "Contribuinte"
            Text            =   ""
            Formato         =   8
            Restricao       =   2
            CorFundo        =   -2147483633
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtCPF1 
            Height          =   300
            Left            =   5190
            TabIndex        =   26
            Tag             =   "CPF 1"
            Top             =   5235
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   529
            Caption         =   "CPF"
            Text            =   ""
            Formato         =   1
            Restricao       =   2
            CorFundo        =   -2147483633
            MaxLen          =   15
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtTestemunha1 
            Height          =   300
            Left            =   285
            TabIndex        =   25
            Tag             =   "Testemunha 1"
            Top             =   5235
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   529
            Caption         =   "1"
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtCPF2 
            Height          =   300
            Left            =   5190
            TabIndex        =   28
            Tag             =   "CPF 2"
            Top             =   5565
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   529
            Caption         =   "CPF"
            Text            =   ""
            Formato         =   1
            Restricao       =   2
            CorFundo        =   -2147483633
            MaxLen          =   15
            AgruparValores  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtTestemunha2 
            Height          =   300
            Left            =   285
            TabIndex        =   27
            Tag             =   "Testemunha 2"
            Top             =   5565
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   529
            Caption         =   "2"
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtTamFrente 
            Height          =   300
            Left            =   345
            TabIndex        =   11
            Tag             =   "Tamanho Frente"
            Top             =   660
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   529
            Caption         =   "Frente"
            Text            =   ""
            Restricao       =   3
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtTamDireita 
            Height          =   300
            Left            =   315
            TabIndex        =   13
            Tag             =   "Tamanho Direita"
            Top             =   990
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            Caption         =   "Direita"
            Text            =   ""
            Restricao       =   3
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtTamEsquerda 
            Height          =   300
            Left            =   90
            TabIndex        =   15
            Tag             =   "Tamanho Esquerda"
            Top             =   1320
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   529
            Caption         =   "Esquerda"
            Text            =   ""
            Restricao       =   3
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtTamFundos 
            Height          =   300
            Left            =   285
            TabIndex        =   17
            Tag             =   "Tamanho Fundos"
            Top             =   1650
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            Caption         =   "Fundos"
            Text            =   ""
            Restricao       =   3
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtLimDireita 
            Height          =   300
            Left            =   2610
            TabIndex        =   14
            Tag             =   "Limite Direita"
            Top             =   990
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   529
            Caption         =   ""
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtLimEsquerda 
            Height          =   300
            Left            =   2610
            TabIndex        =   16
            Tag             =   "Limite Esquerda"
            Top             =   1320
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   529
            Caption         =   ""
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtLimFundos 
            Height          =   300
            Left            =   2610
            TabIndex        =   18
            Tag             =   "Limite Fundos"
            Top             =   1650
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   529
            Caption         =   ""
            Text            =   ""
            CorFundo        =   -2147483633
         End
         Begin VTOcx.txtVISUAL txtLogradouro 
            Height          =   300
            Left            =   2685
            TabIndex        =   53
            Top             =   330
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   529
            Caption         =   ""
            Text            =   ""
            Enabled         =   0   'False
            CorFundo        =   -2147483633
            RetirarMascara  =   0   'False
         End
         Begin VTOcx.txtVISUAL txtTotal 
            Height          =   300
            Left            =   5790
            TabIndex        =   54
            TabStop         =   0   'False
            Tag             =   " "
            Top             =   3300
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            Caption         =   "Total"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            Mascara         =   "0000"
         End
         Begin VTOcx.txtVISUAL txtQuadra 
            Height          =   300
            Left            =   2415
            TabIndex        =   55
            TabStop         =   0   'False
            Tag             =   " "
            Top             =   2655
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   529
            Caption         =   "Quadra"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            Mascara         =   "00000"
         End
         Begin VTOcx.txtVISUAL txtLote 
            Height          =   300
            Left            =   4245
            TabIndex        =   56
            TabStop         =   0   'False
            Tag             =   " "
            Top             =   2655
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            Caption         =   "Lote"
            Text            =   ""
            Enabled         =   0   'False
            Restricao       =   2
            AlinhamentoTexto=   1
            CorFundo        =   -2147483633
            Mascara         =   "00000"
         End
         Begin VTOcx.cboVISUAL cboDestinacao 
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   2280
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   556
            Caption         =   "Destinação"
            Text            =   ""
            AutoFocaliza    =   0   'False
            CorFundo        =   -2147483633
         End
         Begin VTOcx.cboVISUAL cboEstadoCivil 
            Height          =   315
            Left            =   315
            TabIndex        =   24
            Top             =   4620
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   556
            Caption         =   "Estado Civil"
            Text            =   ""
            AutoFocaliza    =   0   'False
            CorFundo        =   -2147483633
         End
         Begin VB.Label Label4 
            Caption         =   "m  Lim."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   1920
            TabIndex        =   60
            Top             =   1710
            Width           =   795
         End
         Begin VB.Label Label3 
            Caption         =   "m  Lim."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   1905
            TabIndex        =   59
            Top             =   1395
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "m  Lim."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   1905
            TabIndex        =   58
            Top             =   1050
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adquirente"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   47
            Top             =   4035
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Testemunhas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Index           =   4
            Left            =   120
            TabIndex        =   49
            Top             =   4980
            Width           =   1305
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cedente"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   45
            Top             =   3450
            Width           =   810
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Imóvel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   60
            Width           =   690
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aforamento"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   2010
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "m  Lim."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   1905
            TabIndex        =   57
            Top             =   735
            Width           =   795
         End
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   35
      Top             =   7275
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   979
      CorFundo        =   -2147483633
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   375
         Left            =   3780
         TabIndex        =   30
         Top             =   120
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdFicha 
         Height          =   375
         Left            =   5715
         TabIndex        =   32
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "&Ficha"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   375
         Left            =   2805
         TabIndex        =   29
         Top             =   120
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdNovo 
         Height          =   375
         Left            =   1905
         TabIndex        =   9
         Top             =   120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "&Novo"
         Acao            =   1
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdTitulo 
         Height          =   375
         Left            =   4755
         TabIndex        =   31
         Top             =   120
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         Caption         =   "&Título"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   6615
         TabIndex        =   33
         Top             =   120
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
         CorFundo        =   -2147483633
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   1138
      Icone           =   "TCAF103.frx":0F5E
   End
End
Attribute VB_Name = "TCAF103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Aforamento As New cAforamento
Private Const cteTestadaPrincipal As String = "TESTADA PRINCIPAL"

Private Sub cmdExcluir_Click()
    If Not grdAforamento.SelectedItem Is Nothing Then
        If Util.Confirma("Deseja Excluir " & grdAforamento.SelectedItem & "?") Then
            If Aforamento.ConfirmaUltimo(CStr(grdAforamento.SelectedItem), grdAforamento.SelectedItem.SubItems(2)) Then
                    If Aforamento.Excluir(grdAforamento.SelectedItem, grdAforamento.SelectedItem.SubItems(2)) Then
                        Util.Informa "Dados Excluídos."
                        Edita.LimpaCampos Me
                        cmdBuscar_Click
                        tabAforamento.Tabs(1).Selected = True
                    End If
            Else
                Util.Avisa "Impossivel Excluir esse Aforamento."
            End If
        End If
    End If
End Sub

Private Sub cmdFicha_Click()
    If Not grdAforamento.SelectedItem Is Nothing Then
        ImprimirFicha grdAforamento.SelectedItem
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    grdAforamento.Preencher Bdados, ""
    txtConsultaIC.SetFocus
End Sub

Private Sub ImprimirFicha(Ic As String)
    On Error GoTo trata
    Set Rpt = New VSRelatorio
        Rpt.DefinirArquivo Bdados, App.Path & "\TFichaAforamento.rpt"
        Rpt.Selecao = "{VIS_IMOVEL.tim_ic} = '" & Ic & "'"
        Rpt.Visualizar
    Set Rpt = Nothing
trata:
End Sub
Private Sub cmdNovo_Click()
    Edita.LimpaCampos Me
    txtData = Date
    txtOrdem = Aforamento.ProximoAforamento()
    tabAforamento.Tabs(2).Selected = True
    HabilitarCampos True
    txtIC.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    If CriticaCampos(Me) Then
        If Aforamento.Salvar(txtIC, txtOrdem, txtData, txtFicha, txtLivro, txtFolha, txtTotal, txtIMCedente, txtIMAdquirente, txtTamFrente, txtLimFrente, txtTamDireita, txtLimDireita, txtTamEsquerda, txtLimEsquerda, txtTamFundos, txtLimFundos, txtTestemunha1, txtCPF1, txtTestemunha2, txtCPF2, CInt(cboDestinacao.Coluna(1).Valor), CInt(cboEstadoCivil.Coluna(1).Valor)) Then
            Avisa "Aforamento " & txtOrdem & " gravado com sucesso."
            If Confirma("Imprimir o título de aforamento?") Then
                ImprimirTitulo txtIC, txtOrdem
            End If
            cmdNovo_Click
        End If
    End If
End Sub

Private Sub cmdTitulo_Click()
    If Not grdAforamento.SelectedItem Is Nothing Then
        ImprimirTitulo grdAforamento.SelectedItem, grdAforamento.SelectedItem.SubItems(2)
    End If
End Sub

Private Sub ImprimirTitulo(Ic As String, Ordem As String)
    On Error GoTo trata
    Dim Obs As String
    Set Rpt = New VSRelatorio
        Rpt.DefinirArquivo Bdados, App.Path & "\TTituloAforamento.rpt"
        Rpt.Selecao = "{TAB_AFORAMENTO.TAF_TIM_IC} = '" & Trim(Ic) & "' and " & _
                        "{TAB_AFORAMENTO.TAF_NUM_ORDEM} =" & Ordem & " and " & _
                        "{VIS_AREA.tco_descricao_componente} = 'ÁREA DO LOTE'" ' and " & _
                        " {VIS_DESTINACAO.tdi_tgc_cod_grupo} = 11"
        Rpt.Formulas "VTDestinacao", "" & cboDestinacao
        Rpt.Formulas "VTEstadoCivil", "" & cboEstadoCivil
        Obs = Util.Entrada("Titulo de Aforamento", "Obs:")
        If Trim(Obs) <> "" Then Rpt.Formulas "VTObs", Obs
        Rpt.Visualizar
    Set Rpt = Nothing
trata:
End Sub
Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
    
    Set Aforamento = New cAforamento
    Aforamento.PreencherCboDestinacao cboDestinacao
    cboEstadoCivil.PreencherGeral Bdados, "ESTADO CIVIL"
End Sub

Private Sub cmdBuscar_Click()
    Aforamento.PreencherGrid grdAforamento, txtConsultaIC, txtConsultaIMCedente, txtConsultaIMAdquirente, txtConsultaLivro, txtConsultaFicha, txtConsultaDataInicio, txtConsultaDataFim
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Aforamento = Nothing
End Sub

Private Sub grdAforamento_DblClick()
Dim Quadra As String, Lote As String, Destinacao As String
    If Not grdAforamento.SelectedItem Is Nothing Then
        tabAforamento.Tabs(2).Selected = True
        With Aforamento
            If .Buscar(grdAforamento.SelectedItem, grdAforamento.SelectedItem.SubItems(2)) Then
                '1. imovel
                txtIC = .Ic
                txtLogradouro = .BuscarLogradouro(txtIC)
                txtTamFrente = .TamFrente
                txtLimFrente = .LimFrente
                txtTamDireita = .TamDireita
                txtLimDireita = .LimDireita
                txtTamEsquerda = .TamEsquerda
                txtLimEsquerda = .LimEsquerda
                txtTamFundos = .TamFundos
                txtLimFundos = .LimFundos
                '2. Aforamento
                txtData = .DataAforamento
                txtFicha = .BuscaFicha(txtIC)
                txtOrdem = .NumOrdem
                txtLivro = .Livro
                txtFolha = .Folha
                '3. Cedente
                txtIMCedente = .IMCedente
                txtCedente = .BuscarContribuinte(txtIMCedente)
                '4. Adquirente
                txtIMAdquirente = .IMAdquirinte
                txtAdquirente = .BuscarContribuinte(txtIMAdquirente)
                 '5. Testemunhas
                txtTestemunha1 = .TestemunhaUm: txtCPF1 = .CPFUm
                txtTestemunha2 = .TestemunhaDois: txtCPF2 = .CPFDois
                '6. Informações de Destinação para o imóvel com Tipo Territorial
                If .BuscaDestinacao(txtIC, Destinacao) Then
                    cboDestinacao.SetarLinha Destinacao, 1
                    cboDestinacao.Enabled = False
                Else
                    cboDestinacao.SetarLinha .destinacaoTerritorial, 1
                    cboDestinacao.Enabled = True
                End If
                '7. Informações de Quadra e Lote do Imóvel
                If .BuscaQuadraLote(txtIC, Quadra, Lote) Then
                    txtQuadra = Quadra
                    txtLote = Lote
                End If
                '8- estado civil do adquirente
                cboEstadoCivil.SetarLinha .EstadoCivilAdquirinte, 1
            End If
        End With
    End If
End Sub



Private Sub txtConsultaIC_LostFocus()
    If txtConsultaIC <> "" Then txtConsultaLogr = Aforamento.BuscarLogradouro(txtConsultaIC)
End Sub

Private Sub txtConsultaIMAdquirente_LostFocus()
    txtConsultaAdquirente = Aforamento.BuscarContribuinte(txtConsultaIMAdquirente)
End Sub

Private Sub txtConsultaIMCedente_LostFocus()
    txtConsultaCedente = Aforamento.BuscarContribuinte(txtConsultaIMCedente)
End Sub

Private Sub txtCPF2_LostFocus()
    If txtCPF2 = txtCPF1 Then
        Util.Avisa "Testemunhas devem ser distintas."
        txtCPF2 = ""
        txtCPF2.SetFocus
    End If
End Sub


Private Sub txtic_LostFocus()
    Dim Numero As String, Ficha As String, Livro As String, Folha As String, Im As String, TamFrente As String, TamDireita As String, LimDireita As String, TamEsquerda As String, LimEsquerda As String, TamFundos As String, LimFundos As String, IMCedente As String, Total As String
    Dim PosVirgula As Integer
    Dim Quadra As String, Lote As String
    txtLogradouro = Aforamento.BuscarLogradouro(txtIC)
    'BuscarImovel txtIC
    Aforamento.BuscarImovel txtIC, Numero, Ficha, Livro, Folha, Im, TamFrente, TamDireita, LimDireita, TamEsquerda, LimEsquerda, TamFundos, LimFundos, IMCedente, Total
        If txtLogradouro <> "" Then
            PosVirgula = Edita.PosPic(txtLogradouro, ",")
            txtLimFrente = Left(txtLogradouro, PosVirgula - 1)
        End If
        txtFicha = Ficha
        txtLivro = Livro
        txtFolha = Folha
        txtTamFrente = TamFrente
        txtTamDireita = TamDireita
        txtLimDireita = LimDireita
        txtTamEsquerda = TamEsquerda
        txtLimEsquerda = LimEsquerda
        txtTamFundos = TamFundos
        txtLimFundos = LimFundos
        txtIMCedente = IMCedente
        txtCedente = Aforamento.BuscarContribuinte(txtIMCedente)
        txtTotal = Total
        
        Aforamento.BuscaQuadraLote txtIC, Quadra, Lote
        txtQuadra = Quadra
        txtLote = Lote
        txtAdquirente = ""
        txtIMAdquirente = ""
        txtCPF1 = ""
        txtCPF2 = ""
        txtTestemunha1 = ""
        txtTestemunha2 = ""
End Sub

Private Sub txtIMAdquirente_LostFocus()
    If txtIMAdquirente = "" Then Exit Sub
    If txtIMAdquirente = txtIMCedente Then
        Util.Avisa "Cedente e Adquirente não podem ser o mesmo contribuinte."
        txtIMAdquirente = ""
        txtIMAdquirente.SetFocus
    Else
        txtAdquirente = Aforamento.BuscarContribuinte(txtIMAdquirente)
    End If
End Sub

Private Sub HabilitarCampos(Valor As Boolean)
    txtIC.Enabled = Valor
    txtTamFrente.Enabled = Valor: txtLimFrente.Enabled = Valor
    txtTamDireita.Enabled = Valor: txtLimDireita.Enabled = Valor
    txtTamEsquerda.Enabled = Valor: txtLimEsquerda.Enabled = Valor
    txtTamFundos.Enabled = Valor: txtLimFundos.Enabled = Valor
    txtData.Enabled = Valor
    txtFicha.Enabled = Valor
    txtIMAdquirente.Enabled = Valor
    txtTestemunha1.Enabled = Valor: txtCPF1.Enabled = Valor
    txtTestemunha2.Enabled = Valor: txtCPF2.Enabled = Valor
    cmdSalvar.Enabled = Valor
End Sub

