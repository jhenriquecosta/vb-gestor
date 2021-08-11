VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form TCOB501 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCOB501"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   45
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   30
      Width           =   555
      Begin VB.Image Image1 
         Height          =   600
         Left            =   0
         Picture         =   "TCOB501.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   585
      End
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   810
      Left            =   30
      TabIndex        =   4
      Top             =   630
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   1429
      Altura          =   1905
      Caption         =   " Consultar Por:"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483644
      Begin VTOcx.txtVISUAL txtVISUAL1 
         Height          =   315
         Left            =   195
         TabIndex        =   5
         Top             =   375
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   556
         Caption         =   "Contribuinte"
         Text            =   ""
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   1138
      Descricao       =   "Consulta"
      Icone           =   "TCOB501.frx":2123
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   6060
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   375
         Left            =   7260
         TabIndex        =   6
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Buscar"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   8445
         TabIndex        =   2
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin VTOcx.grdVISUAL Grid 
      Height          =   4590
      Left            =   15
      TabIndex        =   3
      Top             =   1470
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   8096
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
End
Attribute VB_Name = "TCOB501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    Dim Sql As String
    If txtVISUAL1 = "" Then
        Util.Avisa "Informe o nome do contribuinte."
        txtVISUAL1.SetFocus
        Exit Sub
    End If
    
    If Me.Tag = "IC" Then
        Sql = "SELECT distinct  tim_ic  as IC,tim_tci_im as IM,tci_nome as Contribuinte,"
        Sql = Sql & " tim_tlg_cod_logradouro as CodLogr,TTL_NOME as Logr,"
        Sql = Sql & " tlg_nome as Nome,tim_numero as [Nº],TBA_NOME as Bairro,"
        Sql = Sql & " tim_valor as [Valor(R$)] FROM VIS_IMOVEL,    TAB_DETALHE_IMOVEL"
        Sql = Sql & " Where tim_ic = tdi_tim_ic and tci_nome like '%" & txtVISUAL1.Text & "%'"
    Else
       Sql = " SELECT VCI_IM AS Im,"
       Sql = Sql & " VCI_CGC_CPF as DOC,"
       Sql = Sql & " VCI_RAZAO     as [Razão Social],"
       Sql = Sql & " VCI_ENDERECO   as Endereço,"
       Sql = Sql & " VCI_COD_ATIVIDADE as Cod_Atividade,"
       Sql = Sql & " VCI_NOME_ATIVIDADE   as Atividade,"
       Sql = Sql & " VCI_INICIO_ATIVIDADE   As Inicio_Atividade"
       Sql = Sql & " From VIS_CONTRIBUINTE"
    End If
    Grid.Preencher Bdados, Sql
End Sub

Private Sub Grid_DblClick()
 If Grid.ListItems.Count > 1 Then
        Ic = Grid.SelectedItem
        Unload Me
    End If
End Sub
