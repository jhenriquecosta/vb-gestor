VERSION 4.00
Begin VB.Form Objfatu 
   Caption         =   "Objfatu"
   ClientHeight    =   6930
   ClientLeft      =   -255
   ClientTop       =   1725
   ClientWidth     =   11880
   Height          =   7335
   Left            =   -315
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   360
   ScaleWidth      =   2400
   Top             =   1380
   Width           =   12000
   Begin VB.Frame frm_grid_fatu 
      Height          =   2715
      Left            =   2700
      TabIndex        =   0
      Top             =   1620
      Width           =   6675
      Begin VB.TextBox txt229 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1590
         Locked          =   -1  'True
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2130
         Width           =   2085
      End
      Begin VB.TextBox txt242 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   4650
         Locked          =   -1  'True
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1845
      End
      Begin TrueDBGrid.TDBGrid grid_faturamento 
         Height          =   2460
         Left            =   150
         OleObjectBlob   =   "OBTFATU.frx":0000
         TabIndex        =   1
         Top             =   150
         Width           =   6435
      End
   End
End
Attribute VB_Name = "Objfatu"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

