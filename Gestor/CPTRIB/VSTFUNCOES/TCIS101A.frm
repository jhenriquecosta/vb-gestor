VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TCIS101A 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BCP-Consultoria e Tecnologia em Administração Pública"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   6510
      TabIndex        =   3
      Top             =   7200
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdBuscar 
      Height          =   375
      Left            =   5340
      TabIndex        =   2
      Top             =   7200
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "&Buscar"
      Acao            =   5
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin MSComctlLib.ListView Grid 
      Height          =   5415
      Left            =   75
      TabIndex        =   4
      Top             =   1710
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   9551
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   2540
      EndProperty
   End
   Begin Threed.SSFrame fra 
      Height          =   945
      Left            =   60
      TabIndex        =   5
      Top             =   690
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1667
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
      Begin VTOcx.txtVISUAL txtInsc 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Top             =   120
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         Caption         =   "Insc. Municipal"
         Text            =   ""
      End
      Begin VTOcx.txtVISUAL txtContrib 
         Height          =   315
         Left            =   300
         TabIndex        =   1
         Top             =   510
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   556
         Caption         =   "Contribuinte"
         Text            =   ""
      End
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   390
      TabIndex        =   6
      Top             =   2280
      Width           =   375
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   1138
      Icone           =   "TCIS101A.frx":0000
   End
End
Attribute VB_Name = "TCIS101A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cadastro As VSImposto

Private Sub cmdBuscar_Click()
    Dim Sql As String
    Dim Rs As VSRecordset
    
    Sql = "SELECT v.tim_ic as IC, t.tci_nome as Contribuinte,v.ttl_nome as Logr," & _
        "v.tlg_nome as Nome,v.tim_numero as [Nº]," & _
        " v.tim_complemento as Complemento,v.tba_nome as Bairro FROM vis_imovel v, " & _
        " Tab_Contribuinte t " & _
        " where v.tim_tci_im = t.tci_im and v.tim_tsc_cod_sit_cad =1 AND "
    Sql = Sql & "V.TBA_TMU_COD_MUNICIPIO = " & AplicacoesVTFuncoes.Codigo_Municipio & " AND V.tlg_tmu_cod_municipio = " & AplicacoesVTFuncoes.Codigo_Municipio
    
    If Trim(txtInsc) <> "" Then
        Sql = Sql & " and v.tim_tci_im = '" & txtInsc & "'"
    ElseIf Trim(txtContrib) <> "" Then
        Sql = Sql & " and t.tci_nome like '%" & txtContrib & "%'"
    End If
       
    If Not Bdados.AbreTabela(Sql, Rs) Then
        Util.Informa "Nenhum registro encontrado."
    End If
    Bdados.FechaTabela Rs
    MontaGrid Bdados, Grid, Sql, 1400
End Sub

Private Sub cmdEnter_Click()
    SendKeys "{TAB}"
    If Me.ActiveControl.Name = "txtNome" Then
        cmdBuscar_Click
    End If
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    cabVisual.Exibir Bdados, Me.Name, App.Path
    Set cadastro = New VSImposto
End Sub

Private Sub Grid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Util.OrdenaGrid Grid, ColumnHeader
End Sub

Private Sub grid_DblClick()
    Me.Tag = Grid.SelectedItem.Text
    Unload Me
    SendKeys "{TAB}"
End Sub

Private Sub Timer_Timer()
    On Error Resume Next
End Sub


Private Sub txtContrib_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        cmdBuscar_Click
    End If
End Sub

Private Sub txtInsc_KeyPress(KeyAscii As Integer)
    KeyAscii = Edita.AceitaDig(KeyAscii, Numero)
End Sub

Private Sub txtInsc_LostFocus()
    txtInsc = cadastro.FormataInscricao(txtInsc, InscContrib)
End Sub
