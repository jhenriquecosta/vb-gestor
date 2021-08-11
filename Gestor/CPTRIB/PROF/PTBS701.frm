VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Begin VB.Form PTBS701 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PTBS701"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   ControlBox      =   0   'False
   Icon            =   "PTBS701.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   6
      Top             =   5070
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   820
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   4305
         TabIndex        =   4
         ToolTipText     =   "Sair"
         Top             =   105
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   16384
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSalvar 
         Height          =   330
         Left            =   2430
         TabIndex        =   2
         ToolTipText     =   "Salvar"
         Top             =   105
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   16384
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdExcluir 
         Height          =   330
         Left            =   3375
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   105
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         Caption         =   "&Excluir"
         Acao            =   2
         CorBorda        =   16384
         CorFrente       =   16384
      End
   End
   Begin VTOcx.grdVISUAL grdLogradouro 
      Height          =   3765
      Left            =   105
      TabIndex        =   5
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6641
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   192
   End
   Begin VTOcx.txtVISUAL txtPar 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Tag             =   "Código"
      Top             =   810
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      Caption         =   "Modulo"
      Text            =   ""
      Restricao       =   1
      MaxLen          =   4
      RetirarMascara  =   0   'False
   End
   Begin VTOcx.txtVISUAL txtDes 
      Height          =   315
      Left            =   1755
      TabIndex        =   1
      Tag             =   "Descricao"
      Top             =   780
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   556
      Caption         =   "Classe"
      Text            =   ""
      Restricao       =   1
      MaxLen          =   50
      RetirarMascara  =   0   'False
   End
   Begin Cabecalho.cabVISUAL cabVisual1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   1138
      Icone           =   "PTBS701.frx":08CA
   End
End
Attribute VB_Name = "PTBS701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcluir_Click()
    If grdLogradouro.ListItems.Count <= 1 Then Exit Sub
    If txtPar = "" Then
        Avisa "Selecione o Bairro"
        Exit Sub
    End If
    If Util.Confirma("Deseja mesmo apagar " & txtPar & "?") Then
            If BDados.DeletaDados("TAB_MODULO_TRIBUTARIO", " TMT_MODULO = '" & txtPar & "'") Then
                Util.Informa "Dados Excluidos com Sucesso."
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
    
      
    Campos = "   TMT_MODULO, TMT_CLASSE"
    Valores = BDados.PreparaValor(txtPar, txtDes)
    Condicao = " TMT_MODULO = '" & txtPar & "'"
    If BDados.GravaDados("TAB_MODULO_TRIBUTARIO", Valores, Campos, Condicao) Then
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
      
      cabVisual1.Exibir BDados, Me.Name, App.Path
      rodVISUAL1.Exibir BDados, Me.Name, App.Path, App.Minor, App.Revision
      carregaBairro
End Sub


Private Sub carregaBairro()
    Dim sql As String
    
    sql = "SELECT TMT_MODULO AS Modulo, TMT_CLASSE AS Classe From TAB_MODULO_TRIBUTARIO order by TMT_MODULO"
    grdLogradouro.Preencher BDados, sql
End Sub

Private Sub grdLogradouro_dblClick()
    If grdLogradouro.ListItems.Count < 1 Then Exit Sub
    txtPar = grdLogradouro.SelectedItem
    txtDes = grdLogradouro.SelectedItem.SubItems(1)
End Sub
