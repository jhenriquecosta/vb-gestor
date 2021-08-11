VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5012403C-6EE0-11D5-ADEC-00D0096D1D42}#9.2#0"; "Cabecalho.ocx"
Object = "{EA761AE1-8FDE-4340-8E6D-420E99B0C363}#1.0#0"; "VTCONTROLES.OCX"
Begin VB.Form CTRN102 
   BackColor       =   &H00FFF5EC&
   Caption         =   "CTRN102"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grdConsulta 
      Height          =   1305
      Left            =   1530
      TabIndex        =   9
      Top             =   2220
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4339
      CorBorda        =   8421504
      CorFundo        =   -2147483633
      Caption         =   "Consulta"
      CorTitulo       =   12632256
      CorCaption      =   8388608
      CorDica         =   8388608
      CheckBox        =   -1  'True
   End
   Begin MSComDlg.CommonDialog dlgLocalizar 
      Left            =   2880
      Top             =   3660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Exportar para..."
      Filter          =   "*.vta"
   End
   Begin VTOcx.txtVISUAL txtDestino 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   3990
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   556
      Caption         =   "Destino"
      Text            =   ""
      RetirarMascara  =   0   'False
   End
   Begin VB.CheckBox chkMarcar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFF5EC&
      Caption         =   "Marcar todos"
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
      Height          =   225
      Left            =   3990
      TabIndex        =   5
      Top             =   3690
      Width           =   1635
   End
   Begin VTOcx.grdVISUAL grdSistema 
      Height          =   3195
      Left            =   60
      TabIndex        =   3
      Top             =   690
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   4339
      CorBorda        =   12632064
      Caption         =   "Sistemas"
      CorTitulo       =   12632064
      CorCaption      =   16777215
      CorDica         =   8388608
   End
   Begin Cabecalho.rodVISUAL rodVISUAL1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   4380
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   820
      CorFundo        =   14737632
      CorFrente       =   8421504
      Begin VTOcx.cmdVISUAL cmdVISUAL1 
         Height          =   375
         Left            =   5250
         TabIndex        =   10
         Top             =   75
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Importar"
         Acao            =   8
         CorBorda        =   4210752
         CorFrente       =   4210752
         CorFundo        =   14737632
         Icone           =   "CTRN102.frx":0000
      End
      Begin VTOcx.cmdVISUAL cmdExportar 
         Height          =   375
         Left            =   6510
         TabIndex        =   8
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "E&xportar"
         Acao            =   8
         Enabled         =   0   'False
         CorBorda        =   4210752
         CorFrente       =   4210752
         CorFundo        =   14737632
         Icone           =   "CTRN102.frx":039A
      End
      Begin VTOcx.cmdVISUAL cmdSair 
         Height          =   375
         Left            =   7680
         TabIndex        =   1
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         Caption         =   "Sai&r"
         Acao            =   7
         CorBorda        =   4210752
         CorFrente       =   4210752
         CorFundo        =   14737632
      End
   End
   Begin Cabecalho.cabVISUAL cabVISUAL1 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1138
      Icone           =   "CTRN102.frx":0734
   End
   Begin VTOcx.grdVISUAL grdGrupo 
      Height          =   3195
      Left            =   3990
      TabIndex        =   4
      Top             =   690
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   4339
      CorBorda        =   12632064
      Caption         =   "Exportar"
      CorTitulo       =   12632064
      CorCaption      =   16777215
      CorDica         =   8388608
      CheckBox        =   -1  'True
   End
   Begin VTOcx.cmdVISUAL cmdLocalizar 
      Height          =   315
      Left            =   8100
      TabIndex        =   7
      Top             =   3990
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      Caption         =   "..."
      Enabled         =   0   'False
      CorBorda        =   4210752
      CorFrente       =   4210752
   End
End
Attribute VB_Name = "CTRN102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkMarcar_Click()
    grdGrupo.MarcarTodos chkMarcar.Value
    HabilitarExportar
End Sub

Private Sub cmdExportar_Click()
    Exportar txtDestino
End Sub

Private Sub cmdLocalizar_Click()
    With dlgLocalizar
        .DialogTitle = "Exportar para..."
        .Filter = "Arquivos de Transferência (*.ftt) | *.ftt"
        .ShowSave
        txtDestino = .FileName
    End With
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdVISUAL1_Click()
    Importar txtDestino
End Sub

Private Sub Form_Load()
    grdSistema.Preencher Bdados, "SELECT DISTINCT TSI_COD_SISTEMA, TSI_NOME AS Sistema FROM TAB_SISTEMA,TAB_GRUPO_TRANSFERENCIA WHERE TSI_COD_SISTEMA=TGT_TSI_COD_SISTEMA ORDER BY 1", 0, 3000
End Sub

Private Sub PreencherGrupos(CodSistema As String)
    grdGrupo.Preencher Bdados, "SELECT TGT_GRUPO AS Grupo, TGT_COD_GRUPO  FROM TAB_GRUPO_TRANSFERENCIA WHERE TGT_TSI_COD_SISTEMA='" & CodSistema & "'", 4000, 0
    chkMarcar_Click
End Sub

Private Sub grdGrupo_ItemCheck(ByVal Item As MSComctlLib.IListItem)
    HabilitarExportar
End Sub

Private Sub grdSistema_Click()
    If Not grdSistema.SelectedItem Is Nothing Then
        PreencherGrupos grdSistema.SelectedItem
    End If
End Sub

Private Sub HabilitarExportar()
    Dim Item As ListItem, TemMarcado As Boolean
    
    For Each Item In grdGrupo.ListItems
        TemMarcado = Item.Checked
        If TemMarcado Then Exit For
    Next
    cmdLocalizar.Enabled = TemMarcado
    cmdExportar.Enabled = TemMarcado
End Sub

Public Sub Exportar(Destino As String)
'ROTEIRO
'-----------
'   1. Travar o envio de comandos ao banco
'   2. Abrir o grupo
'   3. Varrer as consultas
'       3.1 Apagar os dados da consulta
'       3.2 Enviar o comando para o arquivo-destino
'       3.3 Coletar os campos da tabela
'       3.1.4 Varrer os dados da tabela
'               3.1.4.1 Coletar os valores do registro
'               3.1.4.2 Enviar o comando para o arquivo-destino
'   4. Destravar o envio de comandos ao banco

    Dim Campos As String, Valores As String
    Dim sql As String, SqlFrom As String, SqlWhere As String
    Dim Consulta As ListItem, Coluna As ColumnHeader, Valor As ListItem
    Dim Conexao As Object
    Dim rsConsultas As Object, QtdColunas As Integer
    Dim Arq As Integer, I As Integer
    On Error GoTo Trata
    Set Conexao = ConectaBanco(grdSistema.SelectedItem)
    If Not Conexao Is Nothing Then
        Arq = FreeFile(0)
        Open Destino For Output As Arq
        Print #Arq, "@" & grdSistema.SelectedItem

        '1.
        Conexao.ModoTexto = True
        
        For Each Consulta In grdGrupo.ListItems
            If Consulta.Checked Then
                '2.
                sql = "SELECT TTT_TABELA,TTT_LIMPAR_DESTINO FROM TAB_TABELA_TRANSFERENCIA WHERE TTT_TGT_COD_GRUPO=" & Consulta.SubItems(1)
                If Conexao.AbreTabela(sql, rsConsultas) Then
                    '3.
                    Do While Not rsConsultas.EOF
                        '3.1
                        SqlFrom = Util.ParseString(rsConsultas(0), " FROM ", 2)
                        SqlFrom = Util.ParseString(SqlFrom, " WHERE ", 1)
                        SqlWhere = Util.ParseString(rsConsultas(0), " WHERE ", 2)
                        SqlWhere = Util.ParseString(SqlWhere, " ORDER BY ", 1)
                        If rsConsultas(1) Then
                            sql = "DELETE " & SqlFrom & IIf(SqlWhere <> "", " WHERE " & SqlWhere, "")
                            Conexao.Executa sql
                            '3.2
                            Print #Arq, Conexao.UltimoComando
                        End If
                        If InStr(1, rsConsultas(0), "SELECT") = 1 Then
                            '3.3
                            grdConsulta.Preencher Conexao, rsConsultas(0)
                            Campos = ""
                            For Each Coluna In grdConsulta.ColumnHeaders
                                Campos = Campos & IIf(Campos <> "", ",", "") & Coluna
                            Next
                            '3.1.4
                            QtdColunas = grdConsulta.ColumnHeaders.Count - 1
                            For Each Valor In grdConsulta.ListItems
                                '3.1.4.1
                                Valores = Conexao.PreparaValor(Valor)
                                For I = 1 To QtdColunas
                                    Valores = Valores & Conexao.PreparaValor(Valor.SubItems(I))
                                Next
                                '3.1.4.2
                                Conexao.InsereDados SqlFrom, Valores, Campos
                                Print #Arq, Conexao.UltimoComando
                            Next
                        Else
                            Print #Arq, rsConsultas(0)
                        End If
                        rsConsultas.MoveNext
                    Loop
                End If
                Conexao.FechaTabela rsConsultas
            End If
        Next
        
        '4.
        Conexao.ModoTexto = False
        
        Close #Arq
        Util.avisa "Exportação concluída com sucesso."
        Set Conexao = Nothing
    End If
    Exit Sub
Trata:
    MsgBox Err.Description, vbOKOnly, "Erro"
End Sub

Private Function ConectaBanco(CodSis As String) As Object
    Dim ArqBanco As String, BdSis As Object
    Dim Tipo As Long
    Dim Dsn As String
    Dim User As String
    Dim Pass As String
    Dim Cat As String

    ArqBanco = App.Path & "\Banco.cfg"
    If Instala.PegaConfig(ArqBanco, CodSis, 1) = "" Then
        Util.erro "Não foi localizada a configuração de conexao ao banco " & CodSis
    Else
        Tipo = Instala.PegaConfig(ArqBanco, CodSis, 0)
        Dsn = Instala.PegaConfig(ArqBanco, CodSis, 1)
        User = Instala.PegaConfig(ArqBanco, CodSis, 2)
        Cat = Instala.PegaConfig(ArqBanco, CodSis, 3)
        Pass = Instala.PegaConfig(ArqBanco, CodSis, 4)
        Set BdSis = CreateObject("VSClass.VSDados")
        If Not BdSis.AbreBanco(Tipo, Dsn, User, Pass, Cat) Then
            Util.erro "Erro ao conectar com o banco " & CodSis
        Else
            Set ConectaBanco = BdSis
        End If
    End If
    

End Function

Public Sub Importar(Origem As String)
'ROTEIRO
'----------
'   1. Varrer o arquivo
'       1.1 Ler o comando do arquivo-origem
'       1.2 Enviar o comando para o banco
    Dim Arq As Integer, Linha As String
    Dim Conexao As Object
    
    If Dir$(Origem) <> "" Then
        Arq = FreeFile(0)
        Open Origem For Input As Arq
        Do While Not EOF(Arq)
            Line Input #Arq, Linha
            If Mid$(Linha, 1, 1) = "@" Then
                Set Conexao = ConectaBanco(Mid$(Linha, 2))
            Else
                Conexao.Executa Linha
            End If
        Loop
        Util.avisa "Importação concluída com sucesso."
    End If
End Sub

