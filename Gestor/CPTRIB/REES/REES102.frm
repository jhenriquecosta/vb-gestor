VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#4.0#0"; "VTControles.ocx"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Begin VB.Form REES102 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   1
      Top             =   5385
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   794
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   330
         Left            =   5925
         TabIndex        =   11
         Top             =   75
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         Caption         =   "&Buscar"
         Acao            =   4
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdLimpar 
         Height          =   330
         Left            =   7020
         TabIndex        =   4
         Top             =   90
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   330
         Left            =   4815
         TabIndex        =   3
         Top             =   90
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   582
         Caption         =   "&Imprimir"
         Acao            =   4
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   330
         Left            =   8115
         TabIndex        =   2
         Top             =   105
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   32768
         CorFrente       =   16384
         CorFoco         =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   1138
      Icone           =   "REES102.frx":0000
   End
   Begin VTOcx.fraVISUAL fraProPrietario 
      Height          =   1020
      Left            =   30
      TabIndex        =   6
      ToolTipText     =   "Pesquisa Contribuintes"
      Top             =   660
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   1799
      Altura          =   1905
      Caption         =   " Dados do Contribuinte"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   -2147483633
      Ocultavel       =   0   'False
      Begin VTOcx.cmdVISUAL cmdOpcao 
         Height          =   285
         Left            =   2760
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   375
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   503
         Caption         =   ""
         Text            =   ""
         Enabled         =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
      Begin VTOcx.txtVISUAL txtIm 
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Tag             =   "Insc. Municipal"
         Top             =   375
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         Caption         =   "Ins. Municipal"
         Text            =   ""
         Restricao       =   2
         CorRotulo       =   16384
         AgruparValores  =   0   'False
      End
      Begin VTOcx.txtVISUAL txtEndereco 
         Height          =   285
         Left            =   450
         TabIndex        =   7
         Top             =   690
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   503
         Caption         =   "Endereço"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
         CorRotulo       =   16384
         CorTexto        =   4194304
      End
   End
   Begin VTOcx.grdVISUAL grdDados 
      Height          =   3585
      Left            =   15
      TabIndex        =   10
      Top             =   1710
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   6324
      CorBorda        =   32768
      Caption         =   "Processos em Andamento"
      CorTitulo       =   32768
      CorCaption      =   16777215
      CorDica         =   32768
   End
End
Attribute VB_Name = "REES102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    If txtIm <> "" Then
        carregaProcesso txtIm
    Else
        carregaProcesso
    End If
End Sub

Private Sub cmdLimpar_Click()
    LimpaCampos Me
    grdDados.ListItems.Clear
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdOpcao_Click()
    AplicacoesVTFuncoes.BuscaInscricao InscContrib, txtIm, txtRazao
End Sub



Private Sub txtIm_LostFocus()
    If txtIm = "" Then Exit Sub
    txtIm = BuscaContribuinte(txtIm, txtRazao, txtEndereco, , etiContribuinte)
End Sub

Private Sub Form_Load()
     cabVisual.Exibir Bdados, Me.Name, App.Path
     rodVISUAL1.Exibir Bdados, Me.Name, App.Path, App.Minor, App.Revision
     If Temp.PegaParametro(Bdados, "TIPO INSCRICAO") = "REDUZIDA" Then
        txtIm.Formato = formNenhum
     End If
    
End Sub

Private Sub grdDados_dblClick()
    If grdDados.ListItems.Count >= 1 Then
        If Confirma("Deseja visualizar o detalhe do processo?", "visualizar?") Then
            Load REES101
            REES103.Tag = grdDados.SelectedItem
            REES103.Show
        End If
    End If
End Sub

Private Sub carregaProcesso(Optional Im As String)
    Dim sql As String
    
    sql = "select TPR_NUMERO_PROCESSO as Processo, "
    sql = sql & " TPR_INSCRICAO as Inscrição, "
    sql = sql & " TPR_DESCRICAO_PEDIDO as Descrição, "
    sql = sql & " TGE_NOME   As Status "
    sql = sql & " From tab_processo, vis_status_Processo "
    sql = sql & " Where TPR_TIPO_PROCESSO = 3 And TPR_STATUS = TGE_CODIGO "
    
    If Im <> "" Then sql = sql & " and TPR_INSCRICAO = '" & Im & "'"
    
    If Not grdDados.Preencher(Bdados, sql, 1200, 1200, 5500, 1500) Then
        Avisa "Busca sem resultados"
    End If
    

End Sub
