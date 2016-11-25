VERSION 5.00
Begin VB.Form frmEdit 
   Caption         =   "Edit"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   420
      Left            =   3015
      TabIndex        =   2
      Top             =   4455
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   420
      Left            =   1710
      TabIndex        =   1
      Top             =   4455
      Width           =   1050
   End
   Begin VB.TextBox txtFile 
      Height          =   4020
      Left            =   225
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   225
      Width           =   7440
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Open dlg1.filename For Input As #1
   txtFile = Input(LOF(1), (1))
End Sub
