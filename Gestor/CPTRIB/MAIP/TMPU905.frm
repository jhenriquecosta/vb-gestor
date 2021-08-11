VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{E0872E25-0E50-421F-B72C-CC6D0210DC30}#1.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TMPU905 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   60
      ScaleHeight     =   570
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TMPU905.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.grdVISUAL lstZona 
      Height          =   4020
      Left            =   75
      TabIndex        =   8
      Top             =   1995
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   4339
   End
   Begin Threed.SSFrame fra 
      Height          =   1230
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   705
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   2170
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      ShadowStyle     =   1
      Begin VTOcx.txtVISUAL txtValor 
         Height          =   270
         Left            =   960
         TabIndex        =   7
         Top             =   810
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   476
         Text            =   ""
      End
      Begin VTOcx.cboVISUAL cboBairro 
         Height          =   315
         Left            =   900
         TabIndex        =   5
         Top             =   75
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
      End
      Begin VTOcx.cboVISUAL cboInfra 
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Top             =   450
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   556
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   1470
      TabIndex        =   3
      Top             =   -570
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   1138
      Icone           =   "TMPU905.frx":2123
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   7170
      TabIndex        =   1
      Top             =   6270
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
   End
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   375
      Left            =   5970
      TabIndex        =   0
      Top             =   6270
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
   End
End
Attribute VB_Name = "TMPU905"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cadastro As VSImposto

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
    Dim Tipologia As Integer
    Dim Estrutura As Integer
    Dim Padrao As Integer
    Dim Valores As String
    Dim Campos As String
    
    If Not Edita.CriticaCampos(Me) Then Exit Sub
    Campos = "TZH_TBA_COD_BAIRRO,TZH_TGC_COD_GRUPO,TZH_VALOR,TZH_TMU_COD_MUNICIPIO"
    Valores = Bdados.PreparaValor(cboBairro.Coluna(0).Valor, cboInfra.Coluna(0).Valor, Bdados.Converte(txtValor, tctexto), Temp.PegaParametro(Bdados, "MUNICIPIO"))
    If Bdados.GravaDados("TAB_ZONA_HOMOGENEA", Valores, Campos, "TZH_TBA_COD_BAIRRO=" & cboBairro.Coluna(0).Valor & " and TZH_TGC_COD_GRUPO =" & cboInfra.Coluna(0).Valor) Then
        Informa "Transação completada."
        Edita.LimpaCampos Me
    End If
    cboBairro.SetFocus
End Sub

Private Sub Form_Activate()
    cboBairro.Preencher Bdados, "select tba_cod_bairro,tba_nome from tab_bairro where tba_tmu_cod_municipio = " & Temp.PegaParametro(Bdados, "MUNICIPIO"), 1
    cboInfra.Preencher Bdados, "select TGL_COD_GRUPO,TGL_NOME_GRUPO from TAB_GRUPO_DETALHE_LOGRADOURO", 1
    lstZona.Preencher Bdados, "Select TBA_COD_BAIRRO as Codigo,TBA_NOME as Bairro,TGL_NOME_GRUPO as Infra,TZH_VALOR as Valor FROM VIS_ZONA_HOMOGENEA order by TBA_COD_BAIRRO asc,TGL_NOME_GRUPO  asc"
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 0
    cabVisual.Exibir Bdados, Me.Name, App.Path
End Sub


Private Sub lstzona_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Confirma("Deseja realmente excluir o item " & lstZona.SelectedItem & "?") Then
            If Bdados.DeletaDados("TAB_CUB", "TCU_COD_ITEM ='" & lstZona.SelectedItem & "'") Then
                Avisa "Dados eliminados com sucesso."
                LimpaCampos Me
                cboBairro.SetFocus
            End If
        End If
    End If
End Sub
