VERSION 5.00
Object = "{81CC7CD2-6894-4EEB-9FC6-A881BF8E4133}#1.1#0"; "VTControles.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VTOcx.grdVISUAL grdVISUAL1 
      Height          =   3135
      Left            =   90
      TabIndex        =   1
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
      Ordenavel       =   0   'False
   End
   Begin VTOcx.cmdVISUAL cmdVISUAL1 
      Height          =   390
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   688
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As VSClass.VSDados

Private Sub cmdVISUAL1_Click()
  '  grdVISUAL1.OcultarRodape = Not grdVISUAL1.OcultarRodape
    
    
    grdVISUAL1.Preencher b, "SELECT TAB_LOGRADOURO.tlg_cod_logradouro AS Codigo, TTL_NOME + ' ' + tlg_nome as Logradouro, TBA_NOME as Bairro,  TAB_LOGRADOURO.tlg_ttl_cod_tip_logr,  TAB_LOGRADOURO.tlg_nome,  TAB_LOGRADOURO.tlg_tba_cod_bairro, TAB_LOGRADOURO.tlg_cod_logradouro_inicial,  TAB_LOGRADOURO.tlg_cod_bairro_inicial,  TAB_LOGRADOURO.tlg_cod_logradouro_final,  TAB_LOGRADOURO.tlg_cod_bairro_final  FROM TAB_BAIRRO, TAB_LOGRADOURO, TAB_TIPO_LOGR WHERE tlg_tba_cod_bairro = tba_cod_bairro AND tlg_ttl_cod_tip_logr=ttl_cod_tip_logr"
    
End Sub

Private Sub Form_Load()

Set b = New VSClass.VSDados

b.AbreBanco SQLServer, "visual", "sgm", "vtsgm", "VTTRIB"

End Sub

