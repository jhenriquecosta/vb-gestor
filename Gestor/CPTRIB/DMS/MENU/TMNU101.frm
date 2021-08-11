VERSION 5.00
Begin VB.MDIForm TMNU101 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8235
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9780
   Icon            =   "TMNU101.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuCAD 
      Caption         =   "&Cadastro"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCADContr 
         Caption         =   "&Cadastro de Contribuintes"
      End
      Begin VB.Menu mnuCADPrest 
         Caption         =   "Cadastro de &Prestadores/Tomadores de Servico"
      End
   End
   Begin VB.Menu mnuAPU 
      Caption         =   "&Apuracão de Imposto"
      Begin VB.Menu mnuAPUDecl 
         Caption         =   "&Preenchimento de Declaracão"
      End
      Begin VB.Menu mnuAPUUnid 
         Caption         =   "&Definicão da Unidade Fiscal"
      End
      Begin VB.Menu mnuAPULinha 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAPUConsulta 
         Caption         =   "&Consulta de Documentos Emitidos"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sai&r"
   End
End
Attribute VB_Name = "TMNU101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Me.Caption = Temp.PegaParametro(BdSis, "RAZAO") & " - " & Temp.PegaParametro(BdSis, "SISTEMA")
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = 1 Then Exit Sub
    Me.WindowState = 2
End Sub

Private Sub mnuAPUConsulta_Click()
    ChamaAplicacao "TOBR401", "VsTObri", "", ""
End Sub

Private Sub mnuAPUDecl_Click()
    ChamaAplicacao "TDEC101", "VsTDecl", "", ""
End Sub

Private Sub mnuAPUGuia_Click()
    ChamaAplicacao "TDEC103", "VsTDecl", "", ""
End Sub

Private Sub mnuAPUUnid_Click()
    ChamaAplicacao "TDEC102", "VsTDecl", "", ""
End Sub

Private Sub mnuCADContr_Click()
    ChamaAplicacao "TMCO101", "VsTEcon", "", ""
End Sub

Private Sub mnuCADPrest_Click()
    ChamaAplicacao "TMCO102", "VsTEcon", "", ""
End Sub

Private Sub mnuSair_Click()
    End
End Sub
