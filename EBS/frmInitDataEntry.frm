VERSION 5.00
Begin VB.Form frmInitDataEntry 
   Caption         =   "Initial Data Entry"
   ClientHeight    =   3195
   ClientLeft      =   8865
   ClientTop       =   7365
   ClientWidth     =   4680
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.ComboBox comboUnits 
      Height          =   315
      ItemData        =   "frmInitDataEntry.frx":0000
      Left            =   810
      List            =   "frmInitDataEntry.frx":000A
      TabIndex        =   0
      Top             =   0
      Width           =   1500
   End
   Begin VB.ComboBox comboCounty 
      Height          =   315
      Index           =   0
      ItemData        =   "frmInitDataEntry.frx":001F
      Left            =   810
      List            =   "frmInitDataEntry.frx":00AE
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1500
   End
   Begin VB.TextBox txtRouteId 
      Height          =   315
      Left            =   810
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtComments 
      Height          =   870
      Left            =   810
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1080
      Width           =   3480
   End
   Begin VB.TextBox txtShrink 
      Height          =   315
      Left            =   1305
      TabIndex        =   4
      Top             =   2070
      Width           =   465
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   180
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1350
      TabIndex        =   6
      Top             =   2520
      Width           =   1230
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2700
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblUnits 
      Caption         =   "Units"
      Height          =   255
      Left            =   315
      TabIndex        =   12
      Top             =   45
      Width           =   495
   End
   Begin VB.Label lblCounty 
      Caption         =   "County"
      Height          =   240
      Left            =   180
      TabIndex        =   11
      Top             =   405
      Width           =   645
   End
   Begin VB.Label lblRouteId 
      Caption         =   "Route ID"
      Height          =   240
      Left            =   45
      TabIndex        =   10
      Top             =   765
      Width           =   780
   End
   Begin VB.Label lblComments 
      Caption         =   "Comments"
      Height          =   240
      Left            =   0
      TabIndex        =   9
      Top             =   1125
      Width           =   825
   End
   Begin VB.Label lblShrink 
      Caption         =   "Shrinkage Factor"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   2115
      Width           =   1320
   End
End
Attribute VB_Name = "frmInitDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
   comboUnits = ""
   comboCounty(0) = ""
   txtRouteId = ""
   txtComments = ""
   txtShrink = ""
End Sub

Private Sub cmdContinue_Click()
   If comboUnits <> Units Or _
      comboCounty(0) <> County Or _
      txtRouteId <> Route Or _
      txtComments <> Comments Or _
      txtShrink <> Shrink Then
      cmdContinue.Enabled = True
   End If
   cmdContinue.Enabled = False
   cmdCancel.Enabled = False
   Write #2, comboUnits, comboCounty(0), txtRouteId, txtComments, txtShrink
   frmCNFData.Show vbModal, Me
End Sub

Private Sub cmdQuit_Click()
   Unload frmInitDataEntry
   frmEbs.mnuDataProc.Enabled = False
   frmEbs.mnuEdit.Enabled = False
'   frmEbs.mnuPrint.Enabled = False
   Close #1
   Close #2
   Close #7
End Sub

Private Sub Form_Load()
Dim MySize
   Open frmEbs.dlg1.filename For Append As #1
   MySize = LOF(1)
   Close #1
   Open "tmp.dat" For Output As #2
   If MySize = 0 Then
      Open frmEbs.dlg1.filename For Append As #1
      comboUnits = ""
      comboCounty(0) = ""
      txtRouteId = ""
      txtComments = ""
      txtShrink = ""
   Else
      Open frmEbs.dlg1.filename For Input As #1
      Input #1, Units, County, route_id, Comments, Shrink
      comboUnits = Units
      comboCounty(0) = County
      txtRouteId = route_id
      txtComments = Comments
      txtShrink = Shrink
   End If
   IPos = InStr(1, frmEbs.dlg1.filename, ".", 0)
   Rpt_filename = Mid(frmEbs.dlg1.filename, 1, IPos) & "Ewk"
   Open Rpt_filename For Output As #7
End Sub
