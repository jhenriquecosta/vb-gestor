VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#1.1#0"; "VTControles.ocx"
Begin VB.Form TCIM401 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grdGradesCadastradas 
      Height          =   2475
      Left            =   60
      TabIndex        =   10
      Top             =   2940
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4339
      CorBorda        =   32768
      Caption         =   "Grades Cadastradas"
      CorTitulo       =   32768
      CorCaption      =   16777215
      OcultarRodape   =   -1  'True
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   1785
      Left            =   60
      TabIndex        =   8
      Top             =   690
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   3149
      Altura          =   1905
      Caption         =   " Grade"
      CorTexto        =   16777215
      CorFaixa        =   32768
      CorFundo        =   16777215
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtBoletimCondominio 
         Height          =   285
         Left            =   3090
         TabIndex        =   5
         Tag             =   "BC"
         Top             =   1350
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         Caption         =   "BC's"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   16777215
         MaxLen          =   5
      End
      Begin VTOcx.txtVISUAL txtBoletimPredial 
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Tag             =   "BP"
         Top             =   1020
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         Caption         =   "BP's"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   16777215
         MaxLen          =   5
      End
      Begin VTOcx.txtVISUAL txtBoletimTerritorial 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Tag             =   "BT"
         Top             =   690
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   503
         Caption         =   "BT's"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   16777215
         MaxLen          =   5
      End
      Begin VTOcx.txtVISUAL txQuadra 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Tag             =   "Quadra"
         Top             =   1350
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         Caption         =   "Quadra"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   16777215
         MaxLen          =   3
      End
      Begin VTOcx.txtVISUAL txtSetor 
         Height          =   285
         Left            =   300
         TabIndex        =   1
         Tag             =   "Setor"
         Top             =   1020
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         Caption         =   "Setor"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   16777215
         MaxLen          =   2
      End
      Begin VTOcx.txtVISUAL txtDistrito 
         Height          =   285
         Left            =   150
         TabIndex        =   0
         Tag             =   "Distrito"
         Top             =   690
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
         Caption         =   "Distrito"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   16777215
         MaxLen          =   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3510
         TabIndex        =   9
         Top             =   390
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   3540
         X2              =   4540
         Y1              =   630
         Y2              =   630
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   1138
      Icone           =   "Tcim401.frx":0000
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   4830
      TabIndex        =   6
      Top             =   2520
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCIM401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSair_Click()
    Unload Me
End Sub

'Private Sub cmdSalvar_Click()
'Dim sCampos As String
'Dim sValores As String
'Dim Sql As String
'
'' GRAVA OS DADOS DE DISTRITO, SETOR E QUADRA PARA A ABERTURA DE UM NOVO LOTE DE DIGITAÇAO
'' UTILIZA-SE PARA O CONTROLE DE LOTES CADASTRADOS E RELATÓRIOS DE PRODUTIVIDADE E CONSISTËNCIA DE LOTES
'' ÉDERSON 29/01/2003 -IMPERATRIZ
'
'    If Not Edita.CriticaCampos(Me) Then
'        Screen.MousePointer = 11
'        sCampos = "TCL_DISTRITO,TCL_SETOR,TCL_QUADRA,TCL_QTD_BT,TCL_QTD_BP,TCL_QTD_BC,TCL_DATA_ABERTURA,TCL_TUS_COD_USUARIO"
'        sValores = Bdados.PreparaValor(txtDistrito, txtSetor, txtQuadra, txtBoletimTerritorial, txtBoletimPredial, txtBoletimCondominio, Format(Date, "dd/mm/yyyy"), Aplicacoes.Usuario)
'        Sql = "SELECT TCL_TUS_COD_USUARIO FROM TAB_CONTROLE_LOTE WHERE " _
'            & "TCL_DISTRITO = " & txtDistrito _
'            & " TCL_SETOR=" & txtSetor _
'            & "TCL_QUADRA = " & txtQuadra
'        If Not Bdados.AbreTabela(Sql) Then
'            If Bdados.InsereDados("", sValores, sCampos) Then
'                Util.Informa "Grade Informada cadastrada."
'            Else
'                Util.Avisa "Grade näo pode ser cadastrada."
'            End If
'        Else
'            Util.Informa "Grade já cadastrada anteriormente. Verifique o quadro abaixo. Digitador: " & Bdados.Tabela(0)
'            Bdados.FechaTabela
'        End If
'        Screen.MousePointer = 0
'    End If
'End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    Call MostraGrades
End Sub

Private Sub MostraGrades()
Dim Sql As String
    Sql = "SELECT TCL_DISTRITO AS Distrito,TCL_SETOR AS Setor,TCL_QUADRA AS Quadra, " _
        & "TCL_DATA_ABERTURA AS Abertura,TCL_TUS_COD_USUARIO AS Digitador,TCL_STATUS AS Status" _
        & "FROM TAB_CONTROLE_LOTE"
    grdGradesCadastradas.Preencher Bdados, Sql
End Sub
