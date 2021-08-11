VERSION 5.00
Object = "{EFE1998D-9A09-441A-815B-3FE6DC0A7FB5}#1.0#0"; "Cabecalho.ocx"
Object = "{A20BD75B-ABC8-4CBF-B2AF-137983075B4A}#1.0#0"; "VTControles.ocx"
Begin VB.Form TOBR411 
   Caption         =   "TOBR411"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1605
      Index           =   3
      Left            =   15
      TabIndex        =   15
      Top             =   630
      Width           =   10860
      Begin VTOcx.txtVISUAL txtCedente 
         Height          =   300
         Left            =   750
         TabIndex        =   1
         Top             =   180
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   529
         Caption         =   "Cedente"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboRestricao 
         Height          =   315
         Left            =   690
         TabIndex        =   16
         Tag             =   "Tributo"
         Top             =   -360
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   556
         Caption         =   "Restrição"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.cboVISUAL cboStatus 
         Height          =   315
         Left            =   5775
         TabIndex        =   7
         Tag             =   "Tributo"
         Top             =   1185
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   556
         Caption         =   "Status"
         Text            =   ""
         AutoFocaliza    =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtAdiquirente 
         Height          =   300
         Left            =   495
         TabIndex        =   2
         Top             =   510
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   529
         Caption         =   "Adiquirente"
         Text            =   ""
         Enabled         =   0   'False
         Requerido       =   0   'False
      End
      Begin VTOcx.txtVISUAL txtExercicioInicial 
         Height          =   300
         Left            =   255
         TabIndex        =   3
         Tag             =   "Periodo Inicial"
         Top             =   840
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   529
         Caption         =   "Periodo Inicial"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtExercicioFinal 
         Height          =   300
         Left            =   2670
         TabIndex        =   4
         Tag             =   "Periodo Final"
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         Caption         =   "Periodo Final"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtgeracaoInicio 
         Height          =   300
         Left            =   4920
         TabIndex        =   5
         Tag             =   "Periodo Inicial"
         Top             =   840
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   529
         Caption         =   "Geração (Início)"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtDAM 
         Height          =   300
         Left            =   8025
         TabIndex        =   0
         Top             =   180
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   529
         Caption         =   "Número DAM"
         Text            =   ""
         Restricao       =   2
         Requerido       =   0   'False
         RetirarMascara  =   0   'False
         AutoTAB         =   -1  'True
      End
      Begin VTOcx.txtVISUAL txtGeracaoFim 
         Height          =   300
         Left            =   7935
         TabIndex        =   6
         Tag             =   "Periodo Inicial"
         Top             =   825
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   529
         Caption         =   "Geração (Fim)"
         Text            =   ""
         Formato         =   0
         Restricao       =   2
         Requerido       =   0   'False
         MinLen          =   4
         AutoTAB         =   -1  'True
      End
      Begin VB.Label LblPercento 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   17
         Top             =   1590
         Width           =   45
      End
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   14
      Top             =   6030
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1005
      Begin VTOcx.cmdVISUAL cmdBuscar 
         Height          =   390
         Left            =   4530
         TabIndex        =   9
         Top             =   120
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   688
         Caption         =   "&Buscar"
         Acao            =   5
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdCancela 
         Height          =   375
         Left            =   8550
         TabIndex        =   12
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "&Limpar"
         Acao            =   6
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   9690
         TabIndex        =   13
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdImprimir 
         Height          =   375
         Left            =   6870
         TabIndex        =   11
         Top             =   120
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   661
         Caption         =   "&Imprimir DAM"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
      Begin VTOcx.cmdVISUAL cmdRelatorio 
         Height          =   390
         Left            =   5685
         TabIndex        =   10
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   688
         Caption         =   "&Relatorio"
         Acao            =   4
         CorBorda        =   8421504
         CorFrente       =   16384
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1138
      Icone           =   "TOBR411.frx":0000
   End
   Begin VTOcx.grdVISUAL grdVISUAL1 
      Height          =   4020
      Left            =   0
      TabIndex        =   8
      Top             =   2250
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   7091
      OcultarRodape   =   -1  'True
   End
End
Attribute VB_Name = "TOBR411"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cobranca As New VSCobranca
Dim Selecao As String
Private Sub cmdBuscar_Click()
    Dim Sql As String
        
    Sql = "Select * from Vis_ITBI where 1 = 1"
    Selecao = "1 = 1"
    If txtCedente <> "" Then
        Sql = Sql & " and Nome_Cedente  like '" & txtCedente & "%'"
        Selecao = Selecao & " and {VIS_ITBI.Nome_Cedente} like '" & txtCedente & "*'"
    End If
    If txtAdiquirente <> "" Then
        Sql = Sql & " and Nome_Adquirente like '" & txtAdiquirente & "%'"
        Selecao = Selecao & " and {VIS_ITBI.Nome_Adquirente} like '" & txtAdiquirente & "*'"
    End If
    If txtExercicioInicial <> "" And txtExercicioFinal <> "" Then
        Sql = Sql & " and Período >= '" & txtExercicioInicial & "' and Período <= '" & txtExercicioFinal & "'"
        Selecao = Selecao & " and {VIS_ITBI.Período} >= '" & txtExercicioInicial & "' and {VIS_ITBI.Período} <= '" & txtExercicioFinal & "'"
    ElseIf txtExercicioInicial <> "" And txtExercicioFinal = "" Then
        Sql = Sql & " and Período >= '" & txtExercicioInicial & "' and Período <= '" & txtExercicioInicial & "'"
        Selecao = Selecao & " and {VIS_ITBI.Período} >= '" & txtExercicioInicial & "' and {VIS_ITBI.Período} <= '" & txtExercicioInicial & "'"
    End If
    If txtgeracaoInicio <> "" And txtGeracaoFim <> "" Then
        Sql = Sql & " and Geração >= " & Bdados.Converte(txtgeracaoInicio, TCDataHora) & " and Geração <= " & Bdados.Converte(txtGeracaoFim, TCDataHora)
        Selecao = Selecao & " and {VIS_ITBI.Geração} >= #" & txtgeracaoInicio & "# and {VIS_ITBI.Geração} <= #" & txtGeracaoFim & "#"
    ElseIf txtgeracaoInicio <> "" And txtGeracaoFim = "" Then
        Selecao = Selecao & " and {VIS_ITBI.Geração} >= #" & txtgeracaoInicio & "# and {VIS_ITBI.Geração} <= #" & txtgeracaoInicio & "#"
    End If
    If cboStatus.Text <> "" Then
        Sql = Sql & " and Status = '" & cboStatus.Text & "'"
        Selecao = Selecao & " and {VIS_ITBI.Status} = '" & cboStatus.Text & "'"
    End If
    If txtDAM <> "" Then
        Sql = Sql & " and Obrigação = '" & txtDAM & "'"
    End If
    
    If grdVISUAL1.Preencher(Bdados, Sql) = False Then
        Avisa "Consulta sem resultados."
    End If
End Sub

Private Sub cmdPesquisaInscricao_Click()

End Sub

Private Sub cmdCancela_Click()
    LimpaCampos Me
End Sub

Private Sub cmdImprimir_Click()
        Dim PicBarra As Object
        With grdVISUAL1
         Cobranca.ImprimeDamITBI .SelectedItem, .SelectedItem.SubItems(7), .SelectedItem.SubItems(9), _
                  .SelectedItem.SubItems(3), .SelectedItem.SubItems(4), Date, .SelectedItem.SubItems(8), _
                      .SelectedItem.SubItems(10), .SelectedItem.SubItems(5), .SelectedItem.SubItems(6), .SelectedItem.SubItems(2), .SelectedItem.SubItems(16), _
                     CDbl(Nvl(0, 0)), CDbl(Nvl(0, 0)), .SelectedItem.SubItems(11), .SelectedItem.SubItems(13), _
                      .SelectedItem.SubItems(15), Temp.PegaParametro(Bdados, "UFM"), .SelectedItem.SubItems(19), .SelectedItem.SubItems(18), PicBarra, .SelectedItem.SubItems(14), .SelectedItem.SubItems(12), .SelectedItem.SubItems(11), .SelectedItem.SubItems(1), Imposto.NomeTributo(ttr_ITBI), Imposto.BuscaCodImposto(Imposto.NomeTributo(ttr_ITBI))
        End With
End Sub

Private Sub cmdRelatorio_Click()
    Dim Rpt As New VSRelatorio
    
    With Rpt
        If Not .DefinirArquivo(Bdados, App.Path & "\TListagemITBI.rpt") Then Exit Sub
        .Cabecalho Temp.PegaParametro(Bdados, "ESTADO"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "SEMFAZ"), Temp.PegaParametro(Bdados, "SETOR")
        .Rodape Temp.PegaParametro(Bdados, "RESPONSAVEL"), Temp.PegaParametro(Bdados, "CLIENTE"), Temp.PegaParametro(Bdados, "ENDERECO CLIENTE"), Aplicacoes.Usuario, Me.Name
        .Selecao = Selecao
        .Visualizar
    End With
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   cboStatus.Preencher Bdados, "select Tge_codigo,tge_nome from vis_status_obrigacao where tge_codigo in (3,8,2)", 1
   txtAdiquirente.Enabled = True
   txtCedente.Enabled = True
   cabVisual.Exibir Bdados, Me.Name, App.Path
    rodVISUAL1.Exibir Bdados, Me.Name, App.Major, App.Minor, App.Revision
End Sub

