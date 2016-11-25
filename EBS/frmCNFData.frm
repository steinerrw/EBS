VERSION 5.00
Begin VB.Form frmCNFData 
   Caption         =   "Station, Cut & Fill"
   ClientHeight    =   4155
   ClientLeft      =   8430
   ClientTop       =   6930
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   7170
   Begin VB.CheckBox chkEquality 
      Height          =   285
      Left            =   3915
      TabIndex        =   21
      Top             =   2880
      Width           =   285
   End
   Begin VB.CheckBox chkBalance 
      Height          =   285
      Left            =   3915
      TabIndex        =   20
      Top             =   2475
      Width           =   285
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stationing"
      Height          =   1095
      Left            =   4995
      TabIndex        =   19
      Top             =   1215
      Width           =   1770
      Begin VB.OptionButton optStation 
         Caption         =   "Stationing"
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   28
         Top             =   765
         Width           =   1275
      End
      Begin VB.OptionButton optStation 
         Caption         =   "Top to Bottom"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   27
         Top             =   495
         Width           =   1455
      End
      Begin VB.OptionButton optStation 
         Caption         =   "Bottom To Top"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   26
         Top             =   225
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Direction"
      Height          =   870
      Left            =   5040
      TabIndex        =   18
      Top             =   315
      Width           =   1545
      Begin VB.OptionButton optDir 
         Caption         =   "Backwards"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   25
         Top             =   540
         Width           =   1320
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Forwards"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   24
         Top             =   270
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   420
      Left            =   1890
      TabIndex        =   17
      Top             =   3465
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   400
      Left            =   3645
      TabIndex        =   16
      Top             =   3465
      Width           =   750
   End
   Begin VB.TextBox txtEndStation 
      Height          =   285
      Left            =   5850
      TabIndex        =   15
      Top             =   2835
      Width           =   1095
   End
   Begin VB.TextBox txtBegStation 
      Height          =   285
      Left            =   5850
      TabIndex        =   12
      Top             =   2475
      Width           =   1095
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   400
      Left            =   4500
      TabIndex        =   11
      Top             =   3465
      Width           =   750
   End
   Begin VB.ListBox lstStatCNF 
      Height          =   2010
      ItemData        =   "frmCNFData.frx":0000
      Left            =   720
      List            =   "frmCNFData.frx":0002
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   315
      Width           =   2805
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   400
      Left            =   5355
      TabIndex        =   8
      Top             =   3465
      Width           =   750
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   400
      Left            =   2790
      TabIndex        =   7
      Top             =   3465
      Width           =   750
   End
   Begin VB.TextBox txtFill 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   2790
      Width           =   650
   End
   Begin VB.TextBox txtCut 
      Height          =   285
      Left            =   990
      TabIndex        =   2
      Top             =   2790
      Width           =   650
   End
   Begin VB.TextBox txtStation 
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   2790
      Width           =   650
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   400
      Left            =   945
      TabIndex        =   6
      Top             =   3465
      Width           =   750
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   400
      Left            =   945
      TabIndex        =   10
      Top             =   3465
      Width           =   750
   End
   Begin VB.Label lblEquality 
      Caption         =   "Equality"
      Height          =   285
      Left            =   3195
      TabIndex        =   23
      Top             =   2880
      Width           =   690
   End
   Begin VB.Label lblBalance 
      Caption         =   "Forced Balance"
      Height          =   285
      Left            =   2565
      TabIndex        =   22
      Top             =   2475
      Width           =   1320
   End
   Begin VB.Label lblEndStation 
      Caption         =   "Ending Station"
      Height          =   240
      Left            =   4680
      TabIndex        =   14
      Top             =   2835
      Width           =   1140
   End
   Begin VB.Label lblBegStation 
      Caption         =   "Beginning Station"
      Height          =   240
      Left            =   4455
      TabIndex        =   13
      Top             =   2520
      Width           =   1365
   End
   Begin VB.Label lblFill 
      Caption         =   "Fill"
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   2475
      Width           =   645
   End
   Begin VB.Label lblCut 
      Caption         =   "Cut"
      Height          =   285
      Left            =   1035
      TabIndex        =   3
      Top             =   2475
      Width           =   645
   End
   Begin VB.Label lblStation 
      Caption         =   "Station"
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   2475
      Width           =   645
   End
End
Attribute VB_Name = "frmCNFData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Units As String
Dim County As String
Dim RouteID As String
Dim Comments As String
Dim Shrink As String

Dim Station As Variant
Dim Cut As Variant
Dim Fill As Variant
Dim Entry As String
Dim Sign As String

Dim Pos1 As Integer
Dim Pos2 As Integer
Dim Pos3 As Integer
Dim Pos4 As Integer

Dim I As Integer
Dim sList As String
Dim InRec As String
Dim IP As Integer
Dim IBP As Integer
Dim NP As Integer
Dim SF As Integer
Dim Bal_Dir As String

Dim Save_Flag As Boolean
Dim Cut_Flag As String
Dim Cut_Sign As String
Dim Fill_Flag As String
Dim Fill_Sign As String

Dim IStation(1000) As Double
Dim ICut(1000) As Double
Dim IFill(1000) As Double
Dim XSign(1000) As String
Dim DCut As Double
Dim DFill As Double

Dim Cut_Dist As Double
Dim Fill_Dist As Double
Dim Sta_Dist As Double
Dim SFF As Double

Dim Tot_Cut As Double
Dim Tot_Fill As Double
Dim GTot_Cut As Double
Dim GTot_Fill As Double
Dim Tot_Dist As Double
Dim Tot_Dirt As Double
Dim Pg_Tot_Cut As Double
Dim Pg_Tot_Fill As Double

Dim Bal_Cut As Integer
Dim Bal_Fill As Integer
Dim Bal_SW As Integer

Dim Cut_Vol(1000) As Double
Dim Fill_Vol(1000) As Double

Dim JBP As Integer
Dim JEP As Integer
Dim STA As Double
Dim INCR As Integer

Dim PTC As Double
Dim PTF As Double
Dim DSTA As Double
Dim Bal_Sta As Double
Dim DBC As Double
Dim DBF2 As Double
Dim DBF As Integer
Dim F_Bal_Fill As Double

Dim NEQU As Integer

Dim Msg As String
Dim style As String
Dim title As String
Dim response As String

Private Sub cmdAdd_Click()
   Station = txtStation
   Cut = txtCut
   Fill = txtFill
   Balance = " "
   If chkBalance.Value = 1 Then
      Balance = "B"
   End If
   Equality = " "
   If chkEquality.Value = 1 Then
      Equality = "E"
   End If
   Sign = "00" & Balance & Equality
   Entry = Station & "* " & Cut & "* " & Fill & "* " & Sign & "*"
   lstStatCNF.AddItem Entry
   txtStation = ""
   txtCut = ""
   txtFill = ""
   chkBalance.Value = 0
   chkEquality.Value = 0
   lstStatCNF.Refresh
   txtStation.SetFocus
   Save_Flag = True
   cmdSave.Visible = True
End Sub

Private Sub cmdDelete_Click()
   lstStatCNF.RemoveItem lstStatCNF.ListIndex
   lstStatCNF.Refresh
   txtStation.SetFocus
   Save_Flag = True
End Sub

Private Sub cmdReplace_Click()
   Station = txtStation
   Cut = txtCut
   Fill = txtFill
   Balance = " "
   If chkBalance.Value = 1 Then
      Balance = "B"
   End If
   Equality = " "
   If chkEquality.Value = 1 Then
      Equality = "E"
   End If
   Sign = "00" & Balance & Equality
   Entry = Station & "* " & Cut & "* " & Fill & "* " & Sign & "*"
   lstStatCNF.RemoveItem lstStatCNF.ListIndex
   lstStatCNF.AddItem Entry
   txtStation = ""
   txtCut = ""
   txtFill = ""
   chkBalance.Value = 0
   chkEquality.Value = 0
   lstStatCNF.Refresh
   txtStation.SetFocus
   Save_Flag = True
   cmdSave.Visible = True
End Sub

Private Sub cmdCancel_Click()
   txtStation = ""
   txtCut = ""
   txtFill = ""
   chkBalance.Value = 0
   chkEquality.Value = 0
   frmCNFData.Refresh
   txtStation.SetFocus
End Sub

Private Sub cmdQuit_Click()
   If Save_Flag = True Then
      cmdSave_Click
      Save_Flag = False
   End If
   Close #1
   Close #2
   Kill "tmp.dat"
   Unload frmCNFData
End Sub
Private Sub CreStr(sList, Station, Cut, Fill, Sign)
   
   Pos1 = InStr(1, sList, "* ", 0)
   Pos2 = InStr(Pos1 + 2, sList, "* ", 0)
   Pos3 = InStr(Pos2 + 2, sList, "* ", 0)
   Pos4 = InStr(Pos3 + 2, sList, "*", 0)
    
   Station = Mid(sList, 1, (Pos1 - 1))
   Cut = Mid(sList, Pos1 + 2, (Pos2 - Pos1 - 2))
   Fill = Mid(sList, Pos2 + 2, (Pos3 - Pos2 - 2))
   Sign = Mid(sList, Pos3 + 2, (Pos4 - Pos3 - 2))
End Sub

Private Sub cmdRun_Click()
Units = frmInitDataEntry.comboUnits
County = frmInitDataEntry.comboCounty(0)
RouteID = frmInitDataEntry.txtRouteId
Comments = frmInitDataEntry.txtComments
Shrink = frmInitDataEntry.txtShrink

   For I = 0 To lstStatCNF.ListCount - 1
      sList = lstStatCNF.List(I)
      Call CreStr(sList, Station, Cut, Fill, Sign)
      
      NP = I + 1
      IStation(NP) = Str(Station)
      ICut(NP) = Str(Cut)
      IFill(NP) = Str(Fill)
      XSign(NP) = Sign
   Next I
   
   Call EBS3A(Units, Shrink, Bal_Dir)
   
   Msg = "FINISHED PROCESSING"
   style = vbOKOnly
   title = "RUN PROC"
   repsponse = MsgBox(Msg, style, title)
   cmdRun.Visible = False
End Sub

Private Sub EBS3A(Units, SF, Bal_Dir)
Dim IFACT As Integer
Dim DA As Double
Dim DB As Double
Dim tolerance As Double
Dim SDBC As String
Dim SDBF As String
Dim STot_Cut As String
Dim STot_Fill As String
Dim SBal_Sta As String
Dim CUnits As String

   IFACT = 54
   DSkipDistance = 200
   tolerance = 5
   CUnits = "C.Y."
   If Units = "Metric" Then
      IFACT = 2
      DSkipDistance = 60.69
      tolerance = 1.524
      CUnits = "C.M."
   End If
   Tot_Cut = 0
   Tot_Fill = 0
   GTot_Cut = 0
   GTot_Fill = 0
   Tot_Dist = 0
   Tot_Dirt = 0
   Pg_Tot_Cut = 0
   Pg_Tot_Fill = 0
   NEQU = 0
   If Bal_Dir = "F" Then
      Call FWD_BAL(STA, NP, JBP, JEP, INCR)
   Else
      Call REV_BAL(STA, NP, JBP, JEP, INCR)
   End If
   Call Trunc(IStation(JBP), IStation(JBP))
   Call Trunc(ICut(JBP), ICut(JBP))
   Call Trunc(IFill(JBP), IFill(JBP))
   Write #7, IStation(JBP), ICut(JBP), "", "", "", "", IFill(JBP), "", "", "", "", ""
   
   IBP = JBP + INCR
   
   If SF > 0 And SF < 1 Then
      X = 1
   Else
      X = 100
   End If
   SFF = 1 + SF / X
   
   Bal_SW = 1
'
   Sta_Dist = Abs(IStation(IBP) - IStation(IBP - INCR))
   Cut_Dist = Sta_Dist
   Fill_Dist = Sta_Dist
   
   Cut_Vol(IBP) = Int((ICut(IBP) + ICut(IBP - INCR)) * Cut_Dist / IFACT + 0.5)
   Fill_Vol(IBP) = Int((IFill(IBP) + IFill(IBP - INCR)) * Fill_Dist / IFACT + 0.5)
'
   If Cut_Vol(IBP) < Fill_Vol(IBP) * SFF Then
      Bal_SW = -1
   End If
   For IP = IBP To JEP Step INCR
      DCut = ICut(IP) + ICut(IP - INCR)
      DFill = IFill(IP) + IFill(IP - INCR)
' equality processing.
      If Len(XSign(IP)) = 4 And Mid(XSign(IP), 4, 1) = "E" Then
         NEQU = NEQU + 1
         IRem = NEQU Mod 2
         If IRem = 0 Then
            Sta_Dist = 0
            Write #7, "** Equality **   ", "(Sta.", IStation(IP - INCR), " = ", "Sta.", IStation(IP), ")", "", "", "", "", ""
         End If
      End If
' equality processing finished.
      Sta_Dist = Abs(IStation(IP) - IStation(IP - INCR))
      Cut_Dist = Sta_Dist
      Fill_Dist = Sta_Dist
      Call Trunc(Cut_Dist, Cut_Dist)
      Call Trunc(Fill_Dist, Fill_Dist)
      Tot_Dist = Tot_Dist + Sta_Dist

      If Mid(XSign(IP), 1, 1) = "1" Or Mid(XSign(IP - INCR), 1, 1) = "1" Then
         Cut_Dist = 0.5 * Sta_Dist
         Call Trunc(Cut_Dist, Cut_Dist)
      End If
      If Mid(XSign(IP), 2, 1) = "1" Or Mid(XSign(IP - INCR), 2, 1) = "1" Then
         Fill_Dist = 0.5 * Sta_Dist
         Call Trunc(Fill_Dist, Fill_Dist)
      End If
      Cut_Vol(IP) = Int(DCut * Cut_Dist / IFACT + 0.5)
      Fill_Vol(IP) = Int(DFill * Fill_Dist / IFACT + 0.5)
      Tot_Cut = Tot_Cut + Cut_Vol(IP)
      Tot_Fill = Tot_Fill + Fill_Vol(IP)
      GTot_Cut = GTot_Cut + Cut_Vol(IP)
      GTot_Fill = GTot_Fill + Fill_Vol(IP)
      Tot_Dirt = Tot_Dirt + Cut_Vol(IP) - Fill_Vol(IP) * SFF
      If ((Bal_SW > 0 And Tot_Dirt < 0) Or (Bal_SW < 0 And Tot_Dirt > 0) Or (IP <> 0 And Tot_Dirt = 0)) Then
         Bal_SW = -1 * Bal_SW
         If Tot_Dist > DSkipDistance Then
            PTC = Tot_Cut - Cut_Vol(IP)
            PTF = Tot_Fill - Fill_Vol(IP)
            DA = (PTC - PTF * SFF) / (Fill_Vol(IP) * SFF - Cut_Vol(IP))
            If DA > 1 Then
               DA = 1
            End If
            DB = Sta_Dist * (1 - DA)
            If DB < tolerance And IP = JEP Then
               DA = 1
            End If
            DSTA = Sta_Dist * DA
            Bal_Sta = IStation(IP - 1) + DSTA
            If Bal_Dir = "B" Then
               Bal_Sta = IStation(IP + 1) - DSTA
            End If
            Call Trunc(Bal_Sta, Bal_Sta)
            DBC = Int(DSTA / Sta_Dist * Cut_Vol(IP) + 0.5)
            DBF2 = DSTA / Sta_Dist * Fill_Vol(IP)
            DBF = Int(DBF2 + 0.5)
            Bal_Cut = PTC + DBC
            Bal_Fill = PTF + DBF
            F_Bal_Fill = Int((PTF + DBF2) * SFF + 0.5)
            Tot_Cut = Cut_Vol(IP) - DBC
            Tot_Fill = Fill_Vol(IP) - DBF
            Tot_Dist = 0
            SDBC = "(" & DBC & ")"
            SDBF = "(" & DBF & ")"
            STot_Cut = "(" & Tot_Cut & ")"
            STot_Fill = "(" & Tot_Fill & ")"
            SBal_Sta = "** " & Bal_Sta
            Write #7, "", "", "", "", SDBC, "", "", "", "", SDBF, "", ""
            Write #7, SBal_Sta, "", "", "", "", Bal_Cut, "", "", "", "", Bal_Fill, F_Bal_Fill
            Write #7, "", "", "", "", STot_Cut, "", "", "", "", STot_Fill, "", ""
         End If
      End If
      Call Trunc(IStation(IP), IStation(IP))
      Call Trunc(ICut(IP), ICut(IP))
      Call Trunc(DCut, DCut)
      Call Trunc(Cut_Dist, Cut_Dist)
      Call Trunc(Cut_Vol(IP), Cut_Vol(IP))
      Call Trunc(Tot_Cut, Tot_Cut)
      Call Trunc(IFill(IP), IFill(IP))
      Call Trunc(DFill, DFill)
      Call Trunc(Fill_Dist, Fill_Dist)
      Call Trunc(Fill_Vol(IP), Fill_Vol(IP))
      Call Trunc(Tot_Fill, Tot_Fill)

      Write #7, IStation(IP), ICut(IP), DCut, Cut_Dist, Cut_Vol(IP), Tot_Cut, _
                              IFill(IP), DFill, Fill_Dist, Fill_Vol(IP), Tot_Fill, _
                              Int(Tot_Fill * SFF + 0.5)
      If Mid(XSign(IP), 1, 1) = "1" Then
         Write #7, "", "----", "", "", "", "", "", "", "", "", "", ""
      End If
      If Mid(XSign(IP), 2, 1) = "1" Then
         Write #7, "", "", "", "", "", "", "----", "", "", "", "", ""
      End If
      Pg_Tot_Cut = Pg_Tot_Cut + Cut_Vol(IP)
      Pg_Tot_Fill = Pg_Tot_Fill + Fill_Vol(IP)
' Forced Balance Processing
      If Mid(XSign(IP), 3, 1) = "B" Then
         If Tot_Dirt < 0 Then
            Tot_Dirt = Tot_Dirt * -1
            Write #7, "** Forced Balance ** Borrow:", Tot_Dirt, CUnits, "", "", "", "", "", "", "", "", ""
         Else
            Write #7, "** Forced Balance ** Waste: ", Tot_Dirt, CUnits, "", "", "", "", "", "", "", "", ""
         End If
         Tot_Cut = 0
         Tot_Fill = 0
         Tot_Dirt = 0
         Tot_Dist = 0
      
         If IP < NP Then
            Write #7, "", "", "", "", "", "0", "", "", "", "", "0", ""
            If Cut_Vol(IP - INCR) < Fill_Vol(IP_INCR) * SFF Then
               Bal_SW = 1
            End If
         End If
      End If
' Forced Balance processing finished
   Next IP
   Call Trunc(Tot_Dirt, Tot_Dirt)
   If Tot_Dirt < 0 Then
      Tot_Dirt = Tot_Dirt * -1
      Write #7, "** Finished Balance ** Borrow:", Tot_Dirt, CUnits, "", "", "", "", "", "", "", "", ""
   Else
      Write #7, "** Finished Balance ** Waste: ", Tot_Dirt, CUnits, "", "", "", "", "", "", "", "", ""
   End If
End Sub

Private Sub Trunc(DIN, DOUT)
   DOUT = Int(DIN * 100 + 0.000001) / 100
End Sub

Private Sub FWD_BAL(STA, NP, JBP, JEP, INCR)
   INCR = 1
   If optStation(1) = True Then
      JBP = 1
      JEP = NP
   Else
      Call Find_Sta(txtBegStation, JBP)
      Call Find_Sta(txtEndStation, JEP)
   End If
End Sub

Private Sub REV_BAL(STA, NP, JBP, JEP, INCR)
   INCR = -1
   If optStation(0) = True Then
      JBP = NP
      JEP = 1
   Else
      Call Find_Sta(txtBegStation, JBP)
      Call Find_Sta(txtEndStation, JEP)
   End If
End Sub

Private Sub Find_Sta(Station, I)
   For J = 0 To lstStatCNF.ListCount - 1
      If Int(Station) = IStation(J) Then
         I = J
      End If
   Next J
End Sub

Private Sub cmdSave_Click()
   For I = 0 To lstStatCNF.ListCount - 1
      sList = lstStatCNF.List(I)
      Call CreStr(sList, Station, Cut, Fill, Sign)
      
      Balance = Mid(Sign, 3, 1)
      Equality = Mid(Sign, 4, 1)
      Cut_Sign = "0"
      If Cut = "-" Then
         Cut = 0
         Cut_Sign = 1
      End If
      Fill_Sign = "0"
      If Fill = "-" Then
         Fill = 0
         Fill_Sign = 1
      End If
      Sign = Cut_Sign & Fill_Sign & Balance & Equality
      Write #2, Station, Cut, Fill, Sign
   Next I
   Save_Flag = False
   Close #1
   Close #2
   Open frmEbs.dlg1.filename For Output As #1
   Open "tmp.dat" For Input As #2
   Input #2, Units, County, Route, Comments, Shrink
   Write #1, Units, County, Route, Comments, Shrink
   Do While Not EOF(2)
      Input #2, Station, Cut, Fill, Sign
      Write #1, Station, Cut, Fill, Sign
   Loop
   cmdSave.Visible = False
End Sub

Private Sub Form_Load()
   While Not EOF(1)
      Input #1, Station, Cut, Fill, Sign
      Entry = Station & "* " & Cut & "* " & Fill & "* " & Sign & "*"
      lstStatCNF.AddItem Entry
   Wend
   Save_Flag = False
   Frame2.Visible = False
   optStation(0).Visible = False
   optStation(1).Visible = False
   optStation(2).Visible = False
   
   lblBegStation.Visible = False
   lblEndStation.Visible = False
   txtBegStation.Visible = False
   txtEndStation.Visible = False
   cmdRun.Visible = False
   cmdSave.Visible = False
End Sub

Private Sub lstStatCNF_DblClick()
   sList = lstStatCNF.List(lstStatCNF.ListIndex)
   Pos1 = InStr(1, sList, "* ", 0)
   Pos2 = InStr(Pos1 + 2, sList, "* ", 0)
   Pos3 = InStr(Pos2 + 2, sList, "* ", 0)
   Pos4 = InStr(Pos3 + 2, sList, "*", 0)
   
   txtStation = Mid(sList, 1, (Pos1 - 1))
   txtCut = Mid(sList, Pos1 + 2, (Pos2 - Pos1 - 2))
   txtFill = Mid(sList, Pos2 + 2, (Pos3 - Pos2 - 2))
   Sign = Mid(sList, Pos3 + 2, (Pos4 - Pos3 - 2))
   chkBalance.Value = 0
   If Mid(Sign, 3, 1) = "B" Then
      chkBalance.Value = 1
   End If
   chkEquality.Value = 0
   If Mid(Sign, 4, 1) = "E" Then
      chkEquality.Value = 1
   End If
   cmdReplace.Visible = True
   cmdAdd.Visible = False
   cmdSave.Visible = False
End Sub

Private Sub optDir_Click(Index As Integer)
   If optDir(0) = True Or optDir(1) = True Then
      cmdAdd.Visible = False
      cmdReplace.Visible = False
      cmdDelete.Visible = False
      cmdCancel.Visible = False
      cmdSave.Visible = False
      txtStation = ""
      txtCut = ""
      txtFill = ""
      Frame2.Visible = True
      optStation(2).Visible = True
      If optDir(0) = True Then
         Bal_Dir = "F"
         If optStation(0).Visible = True Then
            optStation(0).Visible = False
         End If
         optStation(1).Visible = True
      Else
         Bal_Dir = "B"
         If optStation(1).Visible = True Then
            optStation(1).Visible = False
         End If
         optStation(0).Visible = True
      End If
   Else
      lblBottomTop.Visible = False
      lblTopBottom.Visible = False
      lblStationing.Visible = False
      optStation(0).Visible = False
      optStation(1).Visible = False
      optStation(2).Visible = False
      lblBegStation.Visible = False
      lblEndStation.Visible = False
      txtBegStation.Visible = False
      txtEndStation.Visible = False
      cmdRun.Visible = False
   End If
End Sub

Private Sub optDir_GotFocus(Index As Integer)
optStation(0) = False
optStation(1) = False
optStation(2) = False
End Sub

Private Sub optStation_Click(Index As Integer)
   If optStation(2) = True Then
      lblBegStation.Visible = True
      lblEndStation.Visible = True
      txtBegStation.Visible = True
      txtEndStation.Visible = True
      txtBegStation.SetFocus
      txtBegStation = ""
      txtEndStation = ""
   Else
      lblBegStation.Visible = False
      lblEndStation.Visible = False
      txtBegStation.Visible = False
      txtEndStation.Visible = False
      cmdRun.Visible = False
   End If
   If optStation(0) = True Or optStation(1) = True Then
      cmdRun.Visible = True
   Else
      cmdRun.Visible = False
   End If
End Sub

Private Sub txtBegStation_Change()
   If txtBegStation <> "" And txtEndStation <> "" Then
      cmdRun.Visible = True
   Else
      If txtBegStation = "" Or txtEndStation = "" Then
         cmdRun.Visible = False
      End If
   End If
End Sub

Private Sub txtEndStation_Change()
   If txtBegStation <> "" And txtEndStation <> "" Then
      cmdRun.Visible = True
   End If
End Sub

Private Sub txtStation_GotFocus()
   cmdReplace.Visible = False
   cmdAdd.Visible = True
   txtStation.SelLength = 0
   txtStation.SelLength = Len(Station) + 1
End Sub

Private Sub txtCut_GotFocus()
   txtCut.SelLength = 0
   txtCut.SelLength = Len(Cut) + 1
End Sub

Private Sub txtFill_GotFocus()
   txtFill.SelLength = 0
   txtFill.SelLength = Len(Fill) + 1
End Sub

