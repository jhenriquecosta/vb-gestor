VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TBAL401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TBAL401"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   1138
      Icone           =   "TBAL401.frx":0000
   End
   Begin VTOcx.grdVISUAL grdInfra 
      Height          =   4065
      Left            =   30
      TabIndex        =   6
      Top             =   1770
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7170
      CorBorda        =   32768
      Caption         =   "Infrações"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
   Begin Cabecalho.rodVISUAL rod 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   7
      Top             =   5850
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   820
      CorFrente       =   0
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   375
         Left            =   8265
         TabIndex        =   2
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   0
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   375
         Left            =   9240
         TabIndex        =   3
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   0
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   10215
         TabIndex        =   4
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   0
         CorFoco         =   14737632
      End
   End
   Begin VTOcx.fraVISUAL fraInfra 
      Height          =   1065
      Left            =   45
      TabIndex        =   8
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   690
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   1879
      Altura          =   1905
      Caption         =   " Dados da Infração"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtAgravanteUFM 
         Height          =   480
         Left            =   3615
         TabIndex        =   5
         Tag             =   "Agravate"
         Top             =   -570
         Visible         =   0   'False
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   847
         Caption         =   "Agravante (UFM%)"
         Text            =   ""
         Formato         =   5
         Restricao       =   2
         AlinhamentoRotulo=   1
         AlinhamentoTexto=   1
      End
      Begin VTOcx.txtVISUAL txtCodigo 
         Height          =   285
         Left            =   765
         TabIndex        =   9
         Top             =   405
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   503
         Caption         =   "Nº Infração"
         Text            =   ""
         Restricao       =   2
         MaxLen          =   8
      End
      Begin VTOcx.txtVISUAL txtDescricao 
         Height          =   285
         Left            =   915
         TabIndex        =   0
         Tag             =   "Descricao"
         Top             =   720
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   503
         Caption         =   "Descricao"
         Text            =   ""
         TipoLetras      =   0
      End
      Begin VTOcx.cboVISUAL cboGravidade 
         Height          =   510
         Left            =   285
         TabIndex        =   1
         Top             =   -570
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   900
         Caption         =   "Gravidade"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Alinhamento     =   1
      End
   End
End
Attribute VB_Name = "TBAL401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Codigo As New ContaCorrente
Dim CodigoInfracao As String
Private Sub cmdExcluir_Click()
    If grdInfra.ListItems.Count >= 1 Then
        If txtCodigo <> "" Then
            If Util.Confirma("Deseja excluir a infração?", "Excluir Infração?") = True Then
                If Bdados.DeletaDados("TAB_INFRACAO", "TIN_COD_INFRACAO=" & grdInfra.SelectedItem) Then
                    Avisa "Infração excluida com sucesso."
                    LimpaCampos Me
                    PreencherGrid
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
    PreencherGrid
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub



Private Sub Form_Load()
      
      cabVISUAL1.Exibir Bdados, Me.Name, App.Path
      rod.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
      cboGravidade.Preencher Bdados, "select * from TAB_GRAVIDADE_INFRACAO"
End Sub

Private Sub PreencherGrid()
    Dim sql As String
    
    sql = " SELECT TIN_COD_INFRACAO as Código ,tin_referencia as Infração,"
    sql = sql & " tin_descricao_infracao as Descrição,"
    sql = sql & " tin_valor_ufm As VALOR, tin_artigo As Artigo,TIN_AGRAVANTE_UFM AS Agravante"
    sql = sql & " From TAB_INFRACAO where 1 = 1"
    
    If txtCodigo <> "" Then
        sql = sql & " and tin_referencia = '" & txtCodigo & "'"
    End If
    
    If txtDescricao <> "" Then
        sql = sql & " and  tin_descricao_infracao  like '%" & txtDescricao & "%'"
    End If
    
    grdInfra.Preencher Bdados, sql
End Sub

