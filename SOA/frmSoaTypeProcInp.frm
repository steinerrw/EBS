VERSION 5.00
Begin VB.Form frmSoaTypeProcInp 
   Caption         =   "Type Proc Data Entry Screen"
   ClientHeight    =   3195
   ClientLeft      =   9315
   ClientTop       =   5400
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox txtROff 
      Height          =   285
      Left            =   3375
      TabIndex        =   8
      Top             =   1620
      Width           =   645
   End
   Begin VB.TextBox txtLOff 
      Height          =   285
      Left            =   1575
      TabIndex        =   7
      Top             =   1620
      Width           =   645
   End
   Begin VB.TextBox txtField1 
      Height          =   285
      Left            =   2655
      TabIndex        =   5
      Top             =   540
      Width           =   645
   End
   Begin VB.TextBox txtField3 
      Height          =   285
      Left            =   2655
      TabIndex        =   3
      Top             =   1170
      Width           =   645
   End
   Begin VB.TextBox txtField2 
      Height          =   285
      Left            =   2655
      TabIndex        =   1
      Top             =   855
      Width           =   645
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   465
      Left            =   3465
      TabIndex        =   0
      Top             =   2205
      Width           =   1005
   End
   Begin VB.Label lblROff 
      Caption         =   "Right Offset"
      Height          =   240
      Left            =   2385
      TabIndex        =   10
      Top             =   1665
      Width           =   960
   End
   Begin VB.Label lblLOff 
      Caption         =   "Left Offset"
      Height          =   240
      Left            =   630
      TabIndex        =   9
      Top             =   1665
      Width           =   915
   End
   Begin VB.Label lblField1 
      Height          =   195
      Left            =   630
      TabIndex        =   6
      Top             =   585
      Width           =   2040
   End
   Begin VB.Label lblField3 
      Height          =   240
      Left            =   630
      TabIndex        =   4
      Top             =   1215
      Width           =   2040
   End
   Begin VB.Label lblField2 
      Height          =   195
      Left            =   630
      TabIndex        =   2
      Top             =   900
      Width           =   2040
   End
End
Attribute VB_Name = "frmSoaTypeProcInp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Private Sub cmdContinue_Click()
   If frmSoaInp.optTypeProc(0) = True Then
   ElseIf frmSoaInp.optTypeProc(1) = True Then
      BegPnt = Val(txtField2)
      EndPnt = Val(txtField3)
   ElseIf frmSoaInp.optTypeProc(2) = True Then
      BegSta = Val(txtField2)
      EndSta = Val(txtField3)
   ElseIf frmSoaInp.optTypeProc(3) = True Then
      AlignNum = Val(txtField1)
   ElseIf frmSoaInp.optTypeProc(4) = True Then
      AlignNum = Val(txtField1)
      BegPnt = Val(txtField2)
      EndPnt = Val(txtField3)
   ElseIf frmSoaInp.optTypeProc(5) = True Then
      AlignNum = Val(txtField1)
      BegSta = Val(txtField2)
      EndSta = Val(txtField3)
   ElseIf frmSoaInp.optTypeProc(6) = True Then
   End If
   LOff = Val(txtLOff)
   ROff = Val(txtROff)
   Do While Not EOF(4)
      Input #4, Align, ElemType, AlignDesc, x1, y1, X2, Y2, X3, Y3, bsta, ESta, CEsta, A, B, C, dir
      If Align <> 0 And frmSoaInp.optTypeProc(6) = False Then
         If frmSoaInp.optTypeProc(3) = True Or frmSoaInp.optTypeProc(4) = True Or frmSoaInp.optTypeProc(5) = True Then
            Do While AlignNum = Align And Not EOF(4)
               Input #4, Align, ElemTyp, AlignDesc, x1, y1, X2, Y2, X3, Y3, bsta, ESta, CEsta, A, B, C, dir
'
               If ElemType = "L" Then
                  Call SoaLine(x1, y1, X2, Y2, A, B, C, bsta, LOff, ROff, BegPnt, EndPnt, BegSta, EndSta, AlignNum, AlignDesc)
               Else
                  C1 = A
                  C2 = B
                  Radius = C
                  Call SoaCurve(x1, y1, X3, Y3, bsta, C1, C2, Radius, dir, LOff, ROff, BegPnt, EndPnt, BegSta, EndSta, AlignNum, AlignDesc)
               End If
'
            Loop
         Else
'
            If ElemType = "L" Then
               Call SoaLine(x1, y1, X2, Y2, A, B, C, bsta, LOff, ROff, BegPnt, EndPnt, BegSta, EndSta, Align, AlignDesc)
            Else
               C1 = A
               C2 = B
               Radius = C
               Call SoaCurve(x1, y1, X3, Y3, bsta, C1, C2, Radius, dir, LOff, ROff, BegPnt, EndPnt, BegSta, EndSta, Align, AlignDesc)
            End If
'
         End If
      End If
   Loop
   Unload frmSoaTypeProcInp
End Sub


Private Sub Form_Initialize()
   txtLOff = -100
   txtROff = 100
   If frmSoaInp.optTypeProc(6) = True Then
      If Align_Num <> Align_nr Then
         If wUnits = "English" Then
            txtLOff = -250
            txtROff = 250
         Else
            txtLOff = -100
            txtROff = 100
         End If
      End If
      Align_Num = Align_nr
   Else
      If AlnCntr = 0 Or Align_Num <> Align_nr Then
         If wUnits = "English" Then
            txtLOff = -50
            txtROff = 50
         Else
            txtLOff = -15
            txtROff = 15
         End If
      End If
      Align_Num = Align_nr
   End If

End Sub

Private Sub Form_Load()
   txtField1.Visible = False
   txtField2.Visible = False
   txtField3.Visible = False
   lblField1.Visible = False
   lblField2.Visible = False
   lblField3.Visible = False
   cmdContinue.Visible = False
End Sub

Private Sub Form_Activate()
Dim txtLOff As Integer
Dim txtROff As Integer

   If frmSoaInp.optTypeProc(0) = True Then         ' All / All          ~popt 1
   ElseIf frmSoaInp.optTypeProc(1) = True Then     ' All / Points       ~popt 2
      txtField2.Enabled = True
      txtField3.Enabled = True
      txtField2.Visible = True
      txtField3.Visible = True
      lblField2.Visible = True
      lblField3.Visible = True
      lblField2.Caption = "Enter Begining Point"
      lblField3.Caption = "Enter Ending Point"
      txtField2.SetFocus
   ElseIf frmSoaInp.optTypeProc(2) = True Then     ' All / Station      ~popt 3
      txtField2.Enabled = True
      txtField3.Enabled = True
      txtField2.Visible = True
      txtField3.Visible = True
      lblField2.Visible = True
      lblField3.Visible = True
      lblField2.Caption = "Enter Begining Station"
      lblField3.Caption = "Enter Ending Station"
      txtField2.SetFocus
   ElseIf frmSoaInp.optTypeProc(3) = True Then     ' Single / All       ~popt 4
      txtField1.Enabled = True
      txtField1.Visible = True
      lblField1.Visible = True
      lblField1.Caption = "Enter Alignment to Process"
      txtField1.SetFocus
   ElseIf frmSoaInp.optTypeProc(4) = True Then     ' Single / Points    ~popt 5
      txtField1.Enabled = True
      txtField2.Enabled = True
      txtField3.Enabled = True
      txtField1.Visible = True
      txtField2.Visible = True
      txtField3.Visible = True
      lblField1.Visible = True
      lblField2.Visible = True
      lblField3.Visible = True
      lblField1.Caption = "Enter Alignment to Process"
      lblField2.Caption = "Enter Begining Point"
      lblField3.Caption = "Enter Ending Point"
      txtField1.SetFocus
   ElseIf frmSoaInp.optTypeProc(5) = True Then     ' Single / Station   ~popt 6
      txtField1.Enabled = True
      txtField2.Enabled = True
      txtField3.Enabled = True
      txtField1.Visible = True
      txtField2.Visible = True
      txtField3.Visible = True
      lblField1.Visible = True
      lblField2.Visible = True
      lblField3.Visible = True
      lblField1.Caption = "Enter Alignment to Process"
      lblField2.Caption = "Enter Begining Station"
      lblField3.Caption = "Enter Ending Station"
      txtField1.SetFocus
   ElseIf frmSoaInp.optTypeProc(6) = True Then     ' Drainage           ~popt 7
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Close #22
End Sub

Private Sub txtField3_Change()
      cmdContinue.Visible = True
End Sub

