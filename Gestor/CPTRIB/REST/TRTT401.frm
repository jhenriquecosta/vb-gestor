VERSION 5.00
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TRTT401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRTT401"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   1138
      Icone           =   "TRTT401.frx":0000
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   5805
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   873
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   7380
         TabIndex        =   4
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   661
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9300
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
         Left            =   8355
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
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1185
      Left            =   60
      TabIndex        =   5
      Top             =   690
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   2090
      Altura          =   1905
      Caption         =   " Consultar Por :"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtDataInicioConsulta 
         Height          =   315
         Left            =   6495
         TabIndex        =   9
         Top             =   375
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         Caption         =   "Data Inicial"
         Text            =   ""
         Formato         =   0
      End
      Begin VTOcx.txtVISUAL txtProcessoConsulta 
         Height          =   315
         Left            =   255
         TabIndex        =   8
         Top             =   375
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   556
         Caption         =   "Processo"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtDam 
         Height          =   315
         Left            =   3045
         TabIndex        =   7
         Top             =   375
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   556
         Caption         =   "Nº Obrigação"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtDataConsultaFim 
         Height          =   315
         Left            =   6615
         TabIndex        =   6
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         Caption         =   "Data Final"
         Text            =   ""
         Formato         =   0
      End
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   3840
      Left            =   45
      TabIndex        =   10
      Top             =   1920
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
Attribute VB_Name = "TRTT401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Obrig As New Obrigacao



Private Sub cmdBuscar_Click()
    Dim sql As String
    sql = "SELECT Trt_NUMERO AS Número,"
    sql = sql & " trt_toc_cod_obrigacao as Obrigação,"
    sql = sql & " trt_tpr_protocolo as Processo,"
    sql = sql & " trt_data as Data,"
    sql = sql & " trt_valor_restituido as Valor,"
    sql = sql & " trt_motivo  as Motivo,"
    sql = sql & " trt_tipo as Tipo "
    sql = sql & " FROM TAB_RESTITUICAO where 1 = 1"
    
    If txtProcessoConsulta <> "" Then
        sql = sql & " and trt_tpr_protocolo = '" & txtProcessoConsulta & "'"
    End If
    
    If txtDam <> "" Then
        sql = sql & " and trt_toc_cod_obrigacao = '" & txtDam & "'"
    End If
    
    If txtDataInicioConsulta <> "" And txtDataConsultaFim <> "" Then
        sql = sql & " AND trt_data >= " & Bdados.Converte(txtDataInicioConsulta, TCDataHora) & " and trt_data <= " & Bdados.Converte(txtDataConsultaFim, TCDataHora)
    ElseIf txtDataInicioConsulta <> "" And txtDataConsultaFim = "" Then
        sql = sql & " AND trt_data >= " & Bdados.Converte(txtDataInicioConsulta, TCDataHora) & " and trt_data <= " & Bdados.Converte(txtDataInicioConsulta, TCDataHora)
    End If
    grdDados.Preencher Bdados, sql
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
