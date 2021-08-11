VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Object = "{0A45DB48-BD0D-11D2-8D14-00104B9E072A}#2.0#0"; "sstabs2.ocx"
Begin VB.Form TRTT102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRTT102"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   1138
      Icone           =   "TRTT102.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   6345
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   6675
         TabIndex        =   5
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL CmdGravar 
         Height          =   375
         Left            =   7665
         TabIndex        =   4
         Top             =   90
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Gravar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9555
         TabIndex        =   3
         Top             =   90
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   8610
         TabIndex        =   2
         Top             =   90
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin ActiveTabs.SSActiveTabs TabDados 
      Height          =   5610
      Left            =   0
      TabIndex        =   6
      Top             =   660
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   9895
      _Version        =   131082
      TabCount        =   2
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
      Tabs            =   "TRTT102.frx":0A42
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel2 
         Height          =   5220
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   9208
         _Version        =   131082
         TabGuid         =   "TRTT102.frx":0AC4
         Begin VB.TextBox txtMotivo 
            Appearance      =   0  'Flat
            Height          =   4005
            Left            =   1500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   1020
            Width           =   8400
         End
         Begin VTOcx.txtVISUAL txtData 
            Height          =   315
            Left            =   1050
            TabIndex        =   10
            Top             =   645
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            Caption         =   "Data"
            Text            =   ""
            Formato         =   0
         End
         Begin VTOcx.txtVISUAL txtValorRestituicao 
            Height          =   300
            Left            =   6765
            TabIndex        =   11
            Top             =   660
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   529
            Caption         =   "Valor Restituição"
            Text            =   ""
            Formato         =   5
            Restricao       =   3
         End
         Begin VTOcx.cboVISUAL cboTipo 
            Height          =   315
            Left            =   3765
            TabIndex        =   12
            ToolTipText     =   "TIPO RESTITUICAO"
            Top             =   645
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            Caption         =   "Tipo"
            Text            =   ""
            AutoFocaliza    =   0   'False
         End
         Begin VTOcx.txtVISUAL txtProcesso 
            Height          =   315
            Left            =   690
            TabIndex        =   20
            Top             =   270
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   556
            Caption         =   "Processo"
            Text            =   ""
            TipoLetras      =   0
         End
         Begin VTOcx.txtVISUAL txtDamReal 
            Height          =   315
            Left            =   3690
            TabIndex        =   21
            Top             =   285
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            Caption         =   "Nº Obrigação"
            Text            =   ""
         End
         Begin VB.Label Label1 
            Caption         =   "Motivo"
            Height          =   285
            Left            =   975
            TabIndex        =   14
            Top             =   960
            Width           =   480
         End
      End
      Begin ActiveTabs.SSActiveTabPanel SSActiveTabPanel1 
         Height          =   5220
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   9208
         _Version        =   131082
         TabGuid         =   "TRTT102.frx":0AEC
         Begin VTOcx.fraVISUAL fraVISUAL1 
            Height          =   1185
            Left            =   120
            TabIndex        =   15
            Top             =   105
            Width           =   10200
            _ExtentX        =   17992
            _ExtentY        =   2090
            Altura          =   1905
            Caption         =   " Consultar Por :"
            CorTexto        =   16777215
            CorFaixa        =   32768
            CorFundo        =   -2147483644
            Ocultavel       =   0   'False
            Begin VTOcx.txtVISUAL txtDataConsultaFim 
               Height          =   315
               Left            =   6615
               TabIndex        =   19
               Top             =   720
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   556
               Caption         =   "Data Final"
               Text            =   ""
               Formato         =   0
            End
            Begin VTOcx.txtVISUAL txtDam 
               Height          =   315
               Left            =   3045
               TabIndex        =   18
               Top             =   375
               Width           =   2865
               _ExtentX        =   5054
               _ExtentY        =   556
               Caption         =   "Nº Obrigação"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtProcessoConsulta 
               Height          =   315
               Left            =   255
               TabIndex        =   17
               Top             =   375
               Width           =   2745
               _ExtentX        =   4842
               _ExtentY        =   556
               Caption         =   "Processo"
               Text            =   ""
            End
            Begin VTOcx.txtVISUAL txtDataInicioConsulta 
               Height          =   315
               Left            =   6495
               TabIndex        =   16
               Top             =   375
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               Caption         =   "Data Inicial"
               Text            =   ""
               Formato         =   0
            End
         End
         Begin VTOcx.grdVISUAL grdDados 
            Height          =   3840
            Left            =   105
            TabIndex        =   9
            Top             =   1335
            Width           =   10245
            _ExtentX        =   18071
            _ExtentY        =   6773
            CorBorda        =   32768
            Caption         =   "Restituições"
            CorTitulo       =   32768
            CorCaption      =   16777215
            CorDica         =   32768
            MarcaUnico      =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "TRTT102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Obrig As New Obrigacao



Private Sub cmdBuscar_Click()
    Dim Sql As String
    Sql = "SELECT Trt_NUMERO AS Número,"
    Sql = Sql & " trt_toc_cod_obrigacao as Obrigação,"
    Sql = Sql & " trt_tpr_protocolo as Processo,"
    Sql = Sql & " trt_data as Data,"
    Sql = Sql & " trt_valor_restituido as Valor,"
    Sql = Sql & " trt_motivo  as Motivo,"
    Sql = Sql & " trt_tipo as Tipo "
    Sql = Sql & " FROM TAB_RESTITUICAO where 1 = 1"
    
    If txtProcessoConsulta <> "" Then
        Sql = Sql & " and trt_tpr_protocolo = '" & txtProcessoConsulta & "'"
    End If
    
    If txtDam <> "" Then
        Sql = Sql & " and trt_toc_cod_obrigacao = '" & txtDam & "'"
    End If
    
    If txtDataInicioConsulta <> "" And txtDataConsultaFim <> "" Then
        Sql = Sql & " AND trt_data >= " & Bdados.Converte(txtDataInicioConsulta, TCDataHora) & " and trt_data <= " & Bdados.Converte(txtDataConsultaFim, TCDataHora)
    ElseIf txtDataInicioConsulta <> "" And txtDataConsultaFim = "" Then
        Sql = Sql & " AND trt_data >= " & Bdados.Converte(txtDataInicioConsulta, TCDataHora) & " and trt_data <= " & Bdados.Converte(txtDataInicioConsulta, TCDataHora)
    End If
    grdDados.Preencher Bdados, Sql
End Sub

Private Sub CmdGravar_Click()
    Dim Valores         As String
    Dim Campos          As String
    Dim Condicao        As String
    Dim NUMERO          As String
    Dim CONTA As New ContaCorrente
    NUMERO = grdDados.SelectedItem
        
    Campos = "TRT_NUMERO,TRT_TOC_COD_OBRIGACAO,TRT_TPR_PROTOCOLO,TRT_DATA,TRT_VALOR_RESTITUIDO,TRT_MOTIVO,TRT_TUS_COD_USUARIO,TRT_TIPO"
    Valores = Bdados.PreparaValor(Bdados.Converte(NUMERO, tctexto), Bdados.Converte(txtDamReal, tctexto), Bdados.Converte(txtProcesso, tctexto), Bdados.Converte(txtData, TCDataHora), Bdados.Converte(txtValorRestituicao, TCMonetario), txtMotivo, AplicacoesVTFuncoes.Usuario, cboTipo.Coluna(1).Valor)
    Condicao = "TRT_NUMERO = '" & grdDados.SelectedItem & "'"
    If Bdados.GravaDados("TAB_RESTITUICAO", Valores, Campos, Condicao) Then
        cmdLimpar_Click
        Avisa "Lançamento restituido com sucesso."
        cmdBuscar_Click
        TabDados.Tabs(1).Selected = True
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    cabVISUAL1.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
    cboTipo.PreencherGeral Bdados, cboTipo.ToolTipText
End Sub




Private Sub grdDados_DblClick()
    If grdDados.ListItems.Count >= 1 Then
        txtProcesso = grdDados.SelectedItem.SubItems(2)
        txtDamReal = grdDados.SelectedItem.SubItems(1)
        txtData = grdDados.SelectedItem.SubItems(3)
        cboTipo.SetarLinha grdDados.SelectedItem.SubItems(6), 1
        txtValorRestituicao = grdDados.SelectedItem.SubItems(4)
        txtMotivo = grdDados.SelectedItem.SubItems(5)
        TabDados.Tabs(2).Selected = True
    End If
End Sub

Private Sub txtValorRestituicao_LostFocus()
'    If cboTipo.Coluna(1).Valor = 1 Then   'INTEGRAL THEN
'        If Nvl(txtValorRestituicao, 0) <> Nvl(txtValor, 0) Then
'            Avisa "O Valor da restituição não pode ser diferente do valor lançado."
'            txtValorRestituicao.SetFocus
'        End If
'    ElseIf cboTipo.Coluna(1).Valor = 2 Then 'PARCIAL
'        If Nvl(txtValorRestituicao, 0) > Nvl(txtValor, 0) Then
'            Avisa "O Valor da restituição não pode ser maior que o  valor lançado."
'            txtValorRestituicao.SetFocus
'        End If
'    End If
    
End Sub
