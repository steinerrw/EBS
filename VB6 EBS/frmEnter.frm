VERSION 5.00
Begin VB.Form frmEnter 
   Caption         =   "Initial Data"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2790
      TabIndex        =   12
      Top             =   2655
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2655
      Width           =   1230
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   270
      TabIndex        =   10
      Top             =   2655
      Width           =   1095
   End
   Begin VB.TextBox txtShrink 
      Height          =   315
      Left            =   1395
      TabIndex        =   8
      Top             =   2205
      Width           =   465
   End
   Begin VB.TextBox txtComments 
      Height          =   870
      Left            =   900
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1215
      Width           =   2850
   End
   Begin VB.TextBox txtRouteId 
      Height          =   315
      Left            =   900
      TabIndex        =   4
      Top             =   855
      Width           =   1455
   End
   Begin VB.ComboBox comboCounty 
      Height          =   315
      Index           =   0
      ItemData        =   "frmEnter.frx":0000
      Left            =   900
      List            =   "frmEnter.frx":008F
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   495
      Width           =   1500
   End
   Begin VB.ComboBox comboUnits 
      Height          =   315
      ItemData        =   "frmEnter.frx":0259
      Left            =   900
      List            =   "frmEnter.frx":0263
      TabIndex        =   1
      Top             =   135
      Width           =   1500
   End
   Begin VB.Label lblShrink 
      Caption         =   "Shrinkage Factor"
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   2250
      Width           =   1320
   End
   Begin VB.Label lblComments 
      Caption         =   "Comments"
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   1260
      Width           =   825
   End
   Begin VB.Label lblRouteId 
      Caption         =   "Route ID"
      Height          =   240
      Left            =   135
      TabIndex        =   5
      Top             =   900
      Width           =   780
   End
   Begin VB.Label LBLcOUNTY 
      Caption         =   "County"
      Height          =   240
      Left            =   270
      TabIndex        =   3
      Top             =   540
      Width           =   645
   End
   Begin VB.Label lblUnits 
      Caption         =   "Units"
      Height          =   255
      Left            =   405
      TabIndex        =   0
      Top             =   180
      Width           =   495
   End
End
Attribute VB_Name = "frmEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
   comboUnits.Clear
   comboCounty(0).Clear
   txtRouteId = ""
   txtComments = ""
   txtShrink = ""
End Sub

Private Sub cmdContinue_Click()
   frmStatcnf.Show vbModal, Me
End Sub

Private Sub cmdQuit_Click()
Unload frmEnter
End Sub

