VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form CDTR103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDTR103"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   1138
      Icone           =   "CDTR103.frx":0000
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1125
      Left            =   45
      TabIndex        =   8
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   690
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1984
      Altura          =   1905
      Caption         =   " Dados do Proprietário"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Borda           =   0
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   2760
         TabIndex        =   1
         Top             =   375
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         Caption         =   ""
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.txtVISUAL txtRazao 
         Height          =   285
         Left            =   3165
         TabIndex        =   2
         Top             =   375
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   503
         Caption         =   "Nome/Razăo Social"
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   4210752
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Tag             =   "Ins. Municipal"
         Top             =   375
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   4210752
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   300
         Left            =   465
         TabIndex        =   3
         Top             =   750
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   529
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   4210752
         CorTexto        =   4194304
      End
   End
   Begin VTOcx.grdVISUAL grdVeiculo 
      Height          =   3390
      Left            =   30
      TabIndex        =   9
      Top             =   1905
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   5980
      CorBorda        =   32768
      Caption         =   "Veículos"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
      OcultarRodape   =   -1  'True
      CheckBox        =   -1  'True
      MarcaUnico      =   -1  'True
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   10
      Top             =   5325
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   820
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   345
         Left            =   6135
         TabIndex        =   11
         Top             =   105
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   609
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdCancelar 
         Height          =   345
         Left            =   7140
         TabIndex        =   4
         Top             =   105
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   609
         Caption         =   "&Salvar"
         Acao            =   3
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   345
         Left            =   9360
         TabIndex        =   6
         Top             =   105
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   609
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   345
         Left            =   8250
         TabIndex        =   5
         Top             =   105
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   609
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
End
Attribute VB_Name = "CDTR103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TransportePassageiro As New etransportePassageiro


Private Sub cmdCancelar_Click()
    Dim i As Integer
    Dim cancelou As Boolean
    
    For i = 1 To grdVeiculo.ListItems.Count
        If (grdVeiculo.ListItems(i).Checked) Then
            With TransportePassageiro.cadastro
                .icad = grdVeiculo.ListItems(i)
                .Im = txtIm
                .Status = ecCancelado
                .Data_Cancelamento = Date
                If .Baixa Then
                    cancelou = True
                Else
                    Exit Sub
                End If
            End With
         End If
    Next
    If (cancelou) Then
        Avisa "Cadastros cancelados com sucesso."
        TransportePassageiro.PreencherGrd grdVeiculo, txtIm
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdVeiculo.ListItems.Clear
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub grdVeiculo_ItemCheck(ByVal Item As MSComctlLib.IListItem)
     txtIm = grdVeiculo.SelectedItem.SubItems(15)
        txtIm_LostFocus
End Sub

Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
    
End Sub
Private Sub cmdBuscar_Click()
    TransportePassageiro.PreencherGrd grdVeiculo, txtIm
End Sub
