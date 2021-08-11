VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "CABECALHO.OCX"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form TMPU702 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TMPU702"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   ControlBox      =   0   'False
   Icon            =   "TMPU702.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   6
      Top             =   4305
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   926
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   7440
         TabIndex        =   4
         ToolTipText     =   "Sair"
         Top             =   105
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   330
         Left            =   5580
         TabIndex        =   2
         ToolTipText     =   "Salvar"
         Top             =   105
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   330
         Left            =   6510
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   105
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   16711680
         CorFrente       =   0
         CorFundo        =   16777088
      End
   End
   Begin VTOcx.grdVISUAL grdLogradouro 
      Height          =   3075
      Left            =   60
      TabIndex        =   5
      Top             =   1185
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   5424
      Caption         =   "Logradouros"
      CorTitulo       =   16711680
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VTOcx.txtVISUAL txtPar 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Tag             =   "Código"
      Top             =   795
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      Caption         =   "Código"
      Text            =   ""
      Restricao       =   2
      MaxLen          =   4
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtDes 
      Height          =   315
      Left            =   2340
      TabIndex        =   1
      Tag             =   "Descricao"
      Top             =   780
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   556
      Caption         =   "Nome"
      Text            =   ""
      MaxLen          =   50
      RetirarMascara  =   0   'False
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   1138
      Icone           =   "TMPU702.frx":08CA
   End
End
Attribute VB_Name = "TMPU702"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private municipio As Integer

Private Sub cmdExcluir_Click()
    If grdLogradouro.ListItems.Count <= 1 Then Exit Sub
    If txtPar = "" Then
        Avisa "Selecione o Bairro"
        Exit Sub
    End If
    If Util.Confirma("Deseja mesmo apagar " & txtPar & "?") Then
            If Bdados.DeletaDados("TAB_BAIRRO", " TBA_TMU_COD_MUNICIPIO = " & municipio & "and TBA_COD_BAIRRO = " & txtPar) Then
                Util.Informa "Bairro apagado."
                carregaBairro
                LimpaCampos Me
            End If
    End If
End Sub

Private Sub cmdSalvar_Click()
    Dim Valores As String
    Dim Campos As String
    Dim Condicao As String
    
     If CriticaCampos(Me) = False Then Exit Sub
    
      
    Campos = " TBA_TMU_COD_MUNICIPIO,TBA_COD_BAIRRO,TBA_NOME"
    Valores = Bdados.PreparaValor(municipio, txtPar, txtDes)
    Condicao = " TBA_COD_BAIRRO = " & Trim(txtPar)
    If Bdados.GravaDados("TAB_BAIRRO", Valores, Campos, Condicao) Then
        Informa "Dados gravados com sucesso!"
        carregaBairro
        cmdLimpar_Click
    End If
End Sub
Private Sub cmdLimpar_Click()
    LimpaCampos Me
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
      municipio = 2886
      carregaBairro
      cabVISUAL1.Exibir Bdados, Me.Name, App.Path
      rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
End Sub


Private Sub carregaBairro()
    Dim Sql As String
    
    Sql = "SELECT   TBA_COD_BAIRRO as Código,TBA_NOME AS Bairro From TAB_BAIRRO     ORDER BY TBA_COD_BAIRRO"
    grdLogradouro.Preencher Bdados, Sql
End Sub

Private Sub grdLogradouro_dblClick()
    If grdLogradouro.ListItems.Count < 1 Then Exit Sub
    txtPar = grdLogradouro.SelectedItem
    txtDes = grdLogradouro.SelectedItem.SubItems(1)
End Sub
