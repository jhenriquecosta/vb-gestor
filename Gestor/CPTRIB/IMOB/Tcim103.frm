VERSION 5.00
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#1.0#0"; "VTControles.ocx"
Begin VB.Form TCIM103 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAT - Sistema de Administração Tributária"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grdGradesCadastradas 
      Height          =   2475
      Left            =   0
      TabIndex        =   12
      Top             =   1350
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   4339
      CorBorda        =   32768
      Caption         =   "Capas Abertas"
      CorTitulo       =   32768
      CorCaption      =   16777215
   End
   Begin VTOcx.fraVISUAL fraVISUAL1 
      Height          =   675
      Left            =   0
      TabIndex        =   11
      Top             =   660
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   1191
      Altura          =   1905
      Caption         =   " Lote"
      CorTexto        =   16777215
      CorFaixa        =   32768
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtQuadra 
         Height          =   285
         Left            =   1950
         TabIndex        =   2
         Tag             =   "Quadra"
         Top             =   330
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         Caption         =   "Quadra"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   14737632
         MaxLen          =   3
         MinLen          =   3
      End
      Begin VTOcx.txtVISUAL txtSetor 
         Height          =   285
         Left            =   1110
         TabIndex        =   1
         Tag             =   "Setor"
         Top             =   330
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Caption         =   "Setor"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   14737632
         MaxLen          =   2
         MinLen          =   2
      End
      Begin VTOcx.txtVISUAL txtDistrito 
         Height          =   285
         Left            =   90
         TabIndex        =   0
         Tag             =   "Distrito"
         Top             =   330
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Caption         =   "Distrito"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   14737632
         MaxLen          =   2
         MinLen          =   2
      End
      Begin VB.Label lblUsuario 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   810
         TabIndex        =   13
         Top             =   405
         Width           =   60
      End
   End
   Begin Cabecalho.cabVISUAL cabVisual 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   1138
      Icone           =   "Tcim103.frx":0000
   End
   Begin VTOcx.cmdVISUAL cmdSair 
      Height          =   375
      Left            =   5460
      TabIndex        =   9
      Top             =   3900
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Sai&r"
      Acao            =   7
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdSalvar 
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   3900
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Salvar"
      Acao            =   3
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.fraVISUAL fraVISUAL2 
      Height          =   675
      Left            =   3090
      TabIndex        =   14
      Top             =   660
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   1191
      Altura          =   1905
      Caption         =   " Quantidade"
      CorTexto        =   16777215
      CorFaixa        =   32768
      Ocultavel       =   0   'False
      Begin VTOcx.txtVISUAL txtBoletimTerritorial 
         Height          =   285
         Left            =   60
         TabIndex        =   3
         Tag             =   "BT"
         Top             =   330
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Caption         =   "BT"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   14737632
         MaxLen          =   5
         MinLen          =   1
      End
      Begin VTOcx.txtVISUAL txtBoletimPredial 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Tag             =   "BP"
         Top             =   330
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Caption         =   "BP"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   14737632
         MaxLen          =   5
         MinLen          =   1
      End
      Begin VTOcx.txtVISUAL txtBoletimCondominio 
         Height          =   285
         Left            =   2100
         TabIndex        =   5
         Tag             =   "BC"
         Top             =   330
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         Caption         =   "BC"
         Text            =   ""
         Restricao       =   2
         CorFundo        =   14737632
         MaxLen          =   5
         MinLen          =   1
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   810
         TabIndex        =   15
         Top             =   405
         Width           =   60
      End
   End
   Begin VTOcx.cmdVISUAL cmdExcluir 
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   3900
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Excluir"
      Acao            =   2
      CorBorda        =   8421504
      CorFrente       =   16384
   End
   Begin VTOcx.cmdVISUAL cmdLimpar 
      Height          =   375
      Left            =   3420
      TabIndex        =   7
      Top             =   3900
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Limpar"
      Acao            =   6
      CorBorda        =   8421504
      CorFrente       =   16384
   End
End
Attribute VB_Name = "TCIM103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Distrito As cDistritoSetor
Dim Capa As cCapa

Private Sub cmdExcluir_Click()
    Dim MsgErro As String
    
    If txtDistrito <> "" Then
        If txtSetor <> "" Then
            If txtQuadra <> "" Then
                If Confirma("Excluir capa " & txtDistrito & "." & txtSetor & "." & txtQuadra & " ?") Then
                    Set Capa = New cCapa
                    If Capa.Buscar(txtDistrito, txtSetor, txtQuadra) Then
                        If Capa.Excluir(MsgErro) Then
                            Avisa "Capa excluída com sucesso."
                            cmdLimpar_Click
                            MostraGrades
                        Else
                            Avisa MsgErro
                        End If
                    Else
                        Erro "Capa não encontrada."
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdLimpar_Click()
    Edita.LimpaCampos Me
    txtDistrito.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSalvar_Click()
' GRAVA OS DADOS DE DISTRITO, SETOR E QUADRA PARA A ABERTURA DE UM NOVO LOTE DE DIGITAÇAO
' UTILIZA-SE PARA O CONTROLE DE LOTES CADASTRADOS E RELATÓRIOS DE PRODUTIVIDADE E CONSISTËNCIA DE LOTES
' ÉDERSON 29/01/2003 -IMPERATRIZ

    If Edita.CriticaCampos(Me) Then
        Screen.MousePointer = 11
        Set Capa = New cCapa
        With Capa
            .Distrito = txtDistrito
            .Setor = txtSetor
            .Quadra = txtQuadra
            .QtdBT = txtBoletimTerritorial
            .QtdBP = txtBoletimPredial
            .QtdBC = txtBoletimCondominio
            .DataAbertura = Format(Date, "dd/mm/yyyy")
            .Status = sglAberto
            .Usuario = Aplicacoes.Usuario
            If .Gravar Then
                Util.Informa "Capa cadastrada."
                If .FechaLote(txtDistrito, txtSetor, txtQuadra) Then Util.Avisa "Lote fechado."
                Set Capa = Nothing
                MostraGrades
                cmdLimpar_Click
            Else
                Util.Avisa "Capa não pôde ser cadastrada."
            End If
        End With
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    cabVisual.Exibir Bdados, Me.Name, App.Path
    MostraGrades
End Sub

Private Sub MostraGrades()
    Set Capa = New cCapa
    Capa.PreencherGrid grdGradesCadastradas, sglAberto
    Set Capa = Nothing
End Sub

Private Sub grdGradesCadastradas_Click()
    If Not (grdGradesCadastradas.SelectedItem Is Nothing) Then
        '0. Distrito, 1. Setor, 2. Quadra, 3. Dt Abertura, 4. Digitador, 5. Status, 6. Alteracao
        Set Capa = New cCapa
        With Capa
            .Buscar grdGradesCadastradas.SelectedItem.Text, grdGradesCadastradas.SelectedItem.SubItems(1), grdGradesCadastradas.SelectedItem.SubItems(2)
            txtDistrito = .Distrito
            txtSetor = .Setor
            txtQuadra = .Quadra
            txtBoletimCondominio = .QtdBC
            txtBoletimPredial = .QtdBP
            txtBoletimTerritorial = .QtdBT
        End With
    End If
End Sub

Private Sub txtDistrito_Change()
    If Len(txtDistrito.Text) = txtDistrito.MaxLen Then SendKeys "{ENTER}"
End Sub

Private Sub txtQuadra_Change()
    If Len(txtQuadra.Text) = txtQuadra.MaxLen Then SendKeys "{ENTER}"
End Sub

Private Sub txtSetor_Change()
    If Len(txtSetor.Text) = txtSetor.MaxLen Then SendKeys "{ENTER}"
End Sub

Private Sub txtSetor_LostFocus()
    If txtDistrito <> "" And txtSetor <> "" Then
        Set Distrito = New cDistritoSetor
        If Not Distrito.Contem(txtDistrito, txtSetor) Then
            Util.Informa "Setor não existe no distrito selecionado."
            txtSetor = "": txtSetor.SetFocus
        End If
        Set Distrito = Nothing
    End If
End Sub
