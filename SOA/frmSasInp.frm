VERSION 5.00
Begin VB.Form frmSasInp 
   Caption         =   "SAS Data Entry Screen"
   ClientHeight    =   4410
   ClientLeft      =   8430
   ClientTop       =   5025
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   6330
   Begin VB.ComboBox comboUnits 
      Height          =   315
      ItemData        =   "frmSasInp.frx":0000
      Left            =   3285
      List            =   "frmSasInp.frx":000A
      TabIndex        =   37
      Top             =   2250
      Width           =   915
   End
   Begin VB.TextBox txtGP3Cnt 
      Height          =   285
      Left            =   2070
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ListBox lstSOESort 
      Height          =   1035
      ItemData        =   "frmSasInp.frx":001F
      Left            =   2385
      List            =   "frmSasInp.frx":0021
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   270
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   420
      Left            =   3375
      TabIndex        =   35
      Top             =   315
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtBegAlign 
      Height          =   285
      Left            =   2070
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   420
      Left            =   3195
      TabIndex        =   29
      Top             =   3600
      Width           =   850
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   420
      Left            =   2205
      TabIndex        =   28
      Top             =   3600
      Width           =   850
   End
   Begin VB.ListBox lstSASData 
      Height          =   1035
      Left            =   2385
      TabIndex        =   27
      Top             =   360
      Width           =   3615
   End
   Begin VB.Frame frameCurve 
      Caption         =   "Curve Types"
      Height          =   1275
      Left            =   225
      TabIndex        =   20
      Top             =   1845
      Width           =   2040
      Begin VB.OptionButton optCurve 
         Caption         =   "PC, POST, POST, PT"
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   26
         Top             =   855
         Width           =   1860
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "PC, POC, PT"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   25
         Top             =   585
         Width           =   1635
      End
      Begin VB.OptionButton optCurve 
         Caption         =   "PC, PI, PT"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   24
         Top             =   315
         Width           =   1635
      End
   End
   Begin VB.Frame frameType 
      Caption         =   "Record Type"
      Height          =   1140
      Left            =   180
      TabIndex        =   19
      Top             =   540
      Width           =   1590
      Begin VB.OptionButton optType 
         Caption         =   "Curve"
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   23
         Top             =   810
         Width           =   1185
      End
      Begin VB.OptionButton optType 
         Caption         =   "Line"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   540
         Width           =   1185
      End
      Begin VB.OptionButton optType 
         Caption         =   "Description"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.TextBox txtPnt4 
      Height          =   285
      Left            =   5265
      TabIndex        =   15
      Top             =   3150
      Width           =   750
   End
   Begin VB.TextBox txtPnt3 
      Height          =   285
      Left            =   5265
      TabIndex        =   13
      Top             =   2700
      Width           =   750
   End
   Begin VB.TextBox txtPnt2 
      Height          =   285
      Left            =   3465
      TabIndex        =   12
      Top             =   3150
      Width           =   750
   End
   Begin VB.TextBox txtPnt1 
      Height          =   285
      Left            =   3465
      TabIndex        =   11
      Top             =   2700
      Width           =   750
   End
   Begin VB.CommandButton cmdNewAlign 
      Caption         =   "New Alignment"
      Height          =   425
      Left            =   225
      TabIndex        =   10
      Top             =   3600
      Width           =   850
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   3420
      TabIndex        =   0
      Top             =   1485
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.TextBox txtAlignment 
      Height          =   285
      Left            =   3285
      TabIndex        =   2
      Top             =   1890
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtStation 
      Height          =   285
      Left            =   4275
      TabIndex        =   4
      Top             =   1890
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   425
      Left            =   4185
      TabIndex        =   9
      Top             =   3600
      Width           =   850
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   425
      Left            =   5175
      TabIndex        =   8
      Top             =   3600
      Width           =   850
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   425
      Left            =   1215
      TabIndex        =   7
      Top             =   3600
      Width           =   850
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   425
      Left            =   1215
      TabIndex        =   6
      Top             =   3600
      Width           =   850
   End
   Begin VB.Label lblUnits 
      Caption         =   "Units"
      Height          =   240
      Left            =   2475
      TabIndex        =   36
      Top             =   2295
      Width           =   870
   End
   Begin VB.Label lblGP3Cnt 
      Caption         =   "Beginning GP3 Point Number "
      Height          =   240
      Left            =   360
      TabIndex        =   34
      Top             =   765
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lblBegAlign 
      Caption         =   "Begining Alignment"
      Height          =   240
      Left            =   585
      TabIndex        =   30
      Top             =   405
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblPnt4 
      Height          =   240
      Left            =   4365
      TabIndex        =   18
      Top             =   3195
      Width           =   750
   End
   Begin VB.Label lblPnt2 
      Height          =   240
      Left            =   2565
      TabIndex        =   17
      Top             =   3195
      Width           =   750
   End
   Begin VB.Label lblPnt1 
      Height          =   240
      Left            =   2565
      TabIndex        =   16
      Top             =   2745
      Width           =   750
   End
   Begin VB.Label lblPnt3 
      Height          =   240
      Left            =   4365
      TabIndex        =   14
      Top             =   2745
      Width           =   750
   End
   Begin VB.Label lblStation 
      Caption         =   "Station"
      Height          =   240
      Left            =   3735
      TabIndex        =   5
      Top             =   1935
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblAlign 
      Caption         =   "Alignment"
      Height          =   240
      Left            =   2475
      TabIndex        =   3
      Top             =   1935
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblDesc 
      Caption         =   "Description"
      Height          =   240
      Left            =   2475
      TabIndex        =   1
      Top             =   1530
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Menu mnuRunAlign 
      Caption         =   "&Run Alignment"
   End
   Begin VB.Menu mnuWithSOE 
      Caption         =   "CenterLine Process &Using SOE file"
   End
   Begin VB.Menu mnuSasPrint 
      Caption         =   "&Print"
   End
End
Attribute VB_Name = "frmSasInp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   mnuWithSOE.Visible = False
   mnuSasPrint.Visible = False
   
   idxCnt = 0
   
   Open New_filename For Input As #2
   Open Aln_filename For Output As #4
   Open Pnt_filename For Append As #5
   Open SasRpt_filename For Output As #6
   Open Cla_filename For Output As #11
'
   Write #6, "Survey File :"; New_filename; " Control Points Listing"; ""; ""; ""
   Write #6, "Point"; "X Coord"; "Y Coord"; "Z Coord"; "Description"; ""
'
'   Line Input #2, Record
   Do While Not EOF(2)
      Line Input #2, Record
      If Mid(Record, 1, 1) <> "*" And Len(Record) > 0 Then
         Pnt = Str(Mid(Record, 1, 5))
         Desc = Mid(Record, 49, 20)
         iY = Str(Mid(Record, 7, 13))
         iX = Str(Mid(Record, 21, 13))
         iZ = Str(Mid(Record, 35, 13))
         
'write  #11 .cla file
         If Mid(Record, 49, 2) = "CL" Or _
            Mid(Record, 49, 3) = "CLP" And _
            Pnt > 0 Then
            Write #11, Pnt, iX, iY, iZ, Desc
         End If
'write  #5 .pnt file
         If Mid(Record, 49, 2) = "XR" Or _
            Mid(Record, 49, 2) = "XL" And _
            Pnt > 0 Then
            Write #5, Pnt, iX, iY, iZ, Desc
         End If
'load array from .pnt file
         If Mid(Record, 49, 3) = "POT" Or _
            Mid(Record, 49, 4) = "POST" Or _
            Mid(Record, 49, 2) = "PC" Or _
            Mid(Record, 49, 3) = "POC" Or _
            Mid(Record, 49, 2) = "PI" Or _
            Mid(Record, 49, 2) = "PT" Or _
            Mid(Record, 49, 2) = "RP" Or _
            Mid(Record, 49, 3) = "PRC" Or _
            Mid(Record, 49, 2) = "CP" Or _
            Mid(Record, 49, 2) = "BM" Or _
            Mid(Record, 49, 3) = "ICL" And _
            Pnt > 0 Then
            idxCnt = idxCnt + 1
            idx_Pnt(idxCnt) = Pnt
            idx_X(idxCnt) = iX
            idx_Y(idxCnt) = iY
            idx_Z(idxCnt) = iZ
            Write #6, Pnt, iX, iY, iZ, Desc, ""
         End If
      End If
   Loop
   Close #2
   Close #11
   Open Alg_filename For Append As #1
   MySize = LOF(1)
   Close #1
   If MySize = 0 Then
      Open Alg_filename For Append As #1
   Else
      Open Alg_filename For Input As #1
      While Not EOF(1)
         Input #1, RecType, Col1, Col2, Col3, Col4
         Entry = RecType & "* " & Col1 & "* " & Col2 & "* " & Col3 & "* " & Col4 & "*"
         lstSASData.AddItem Entry
      Wend
   End If
   Close #1
   cmdReplace.Visible = False
   cmdAdd.Visible = True
   cmdAdd.Default = True
   Save_Flag = "N"
End Sub

Private Sub cmdAdd_Click()
   If optType(0) = True Then
      Call Desc_Proc
      RC = "Y"
   ElseIf optType(1) = True Then
      Call Line_Proc
   ElseIf optType(2) = True Then
      Call Curve_Proc
   End If
   If RC = "Y" Then
      txtPnt1 = ""
      txtPnt2 = ""
      txtPnt3 = ""
      txtPnt4 = ""
      Entry = RecType & "* " & Col1 & "* " & Col2 & "* " & Col3 & "* " & Col4 & "*"
      lstSASData.AddItem Entry
      Col1 = ""
      Col2 = ""
      Col3 = ""
      Col4 = ""
      
      frameType.Visible = True
      optType(0) = False
      optType(1) = False
      optType(2) = False
      optType(0).Visible = True
      optType(1).Visible = True
      optType(2).Visible = True
   
      lstSASData.Refresh
   End If
   Save_Flag = "Y"
End Sub

Private Sub cmdReplace_Click()
   If optType(0) = True Then
      Call Desc_Proc
      RC = "Y"
   ElseIf optType(1) = True Then
      Call Line_Proc
   ElseIf optType(2) = True Then
      Call Curve_Proc
   End If
   
   If RC = "Y" Then
      txtPnt1 = ""
      txtPnt2 = ""
      txtPnt3 = ""
      txtPnt4 = ""
      Entry = RecType & "* " & Col1 & "* " & Col2 & "* " & Col3 & "* " & Col4 & "*"
      lstSASData.RemoveItem lstSASData.ListIndex
      lstSASData.AddItem Entry, lstIndex
      Col1 = ""
      Col2 = ""
      Col3 = ""
      Col4 = ""
      
      frameType.Visible = True
      optType(0) = False
      optType(1) = False
      optType(2) = False
      optType(0).Visible = True
      optType(1).Visible = True
      optType(2).Visible = True
   
      lstSASData.Refresh
   End If
   Save_Flag = "Y"
   cmdAdd.Visible = True
   cmdAdd.Default = True
   cmdReplace.Visible = False
   cmdReplace.Default = False
End Sub

Private Sub Desc_Proc()
   Col1 = txtDescription
   Col2 = txtAlignment
   Col3 = txtStation
   Col4 = comboUnits
   RecType = "D"
   txtDescription.Enabled = False
   txtAlignment.Enabled = False
   txtStation.Enabled = False
   comboUnits.Enabled = False
End Sub

Private Sub Line_Proc()
   Focus_Flag = "N"
   Col1 = txtPnt1
   Call Alignment.Point_Check(Val(Col1), RC, idx_Pnt(), idxCnt)
   If RC = "N" Then
      txtPnt1 = ""
      txtPnt1.SetFocus
      Focus_Flag = "Y"
   End If
   Col2 = txtPnt2
   Call Alignment.Point_Check(Val(Col2), RC, idx_Pnt(), idxCnt)
   If RC = "N" Then
      txtPnt2 = ""
      If Focus_Flag = "N" Then
         txtPnt2.SetFocus
      End If
   End If
   Col3 = ""
   Col4 = ""
   RecType = "L"
      
   If RC = "Y" Then
      txtPnt1.Visible = False
      txtPnt2.Visible = False
      lblPnt1.Visible = False
      lblPnt2.Visible = False
   End If
End Sub

Private Sub Curve_Proc()
   Focus_Flag = "N"
   If optCurve(0) = True Then
      RecType = "C0"
      Call Curve_Pnt_Proc(3)
   ElseIf optCurve(1) = True Then
      RecType = "C1"
      Call Curve_Pnt_Proc(3)
   ElseIf optCurve(2) = True Then
      RecType = "C2"
      Call Curve_Pnt_Proc(4)
   End If

   If RC = "Y" Then
      txtPnt1.Visible = False
      txtPnt2.Visible = False
      txtPnt3.Visible = False
      txtPnt4.Visible = False
      lblPnt1.Visible = False
      lblPnt2.Visible = False
      lblPnt3.Visible = False
      lblPnt4.Visible = False

      frameCurve.Visible = False
      optCurve(0) = False
      optCurve(1) = False
      optCurve(2) = False
      optCurve(0).Visible = False
      optCurve(1).Visible = False
      optCurve(2).Visible = False
   End If
End Sub

Private Sub Curve_Pnt_Proc(numPnt As Integer)
   Col1 = txtPnt1
   Call Point_Check(Val(Col1), RC, idx_Pnt(), idxCnt)
   If RC = "N" Then
      txtPnt1 = ""
      txtPnt1.SetFocus
      Focus_Flag = "Y"
   End If
   Col2 = txtPnt2
   Call Point_Check(Val(Col2), RC, idx_Pnt(), idxCnt)
   If RC = "N" Then
      txtPnt2 = ""
      If Focus_Flag = "N" Then
         txtPnt2.SetFocus
         Focus_Flag = "Y"
      End If
   End If
   Col3 = txtPnt3
   Call Point_Check(Val(Col3), RC, idx_Pnt(), idxCnt)
   If RC = "N" Then
      txtPnt3 = ""
      If Focus_Flag = "N" Then
         txtPnt3.SetFocus
         Focus_Flag = "Y"
      End If
   End If
   If numPnt = 4 Then
      Col4 = txtPnt4
      Call Point_Check(Val(Col4), RC, idx_Pnt, idxCnt)
      If RC = "N" Then
         txtPnt4 = ""
         If Focus_Flag = "N" Then
            txtPnt4.SetFocus
            Focus_Flag = "Y"
         End If
      End If
   End If
End Sub

Private Sub cmdCancel_Click()
   If optType(0) = True Then
      txtDescription = ""
      txtAlignment = ""
      txtStation = ""
      txtDescription.SetFocus
   ElseIf optType(1) = True Then
      txtPnt1 = ""
      txtPnt2 = ""
      txtPnt3 = ""
      txtPnt4 = ""
      txtPnt1.SetFocus
   ElseIf optType(2) = True Then
      txtPnt1 = ""
      txtPnt2 = ""
      txtPnt3 = ""
      txtPnt4 = ""
      txtPnt1.SetFocus
   End If
   cmdAdd.Visible = True
   cmdAdd.Default = True
   cmdReplace.Visible = False
   cmdReplace.Default = False
End Sub

Private Sub cmdDelete_Click()
   lstSASData.RemoveItem lstSASData.ListIndex
   lstSASData.Refresh
End Sub

Private Sub cmdQuit_Click()
   Close #2
   Close #4
   Close #5
   Close #6
   Close #7
   Close #11
   Close #20
   If Save_Flag = "Y" Then
      Call cmdSave_Click
   End If
   Unload frmSasInp
End Sub

Private Sub ReformData(sList, RecType, Col1, Col2, Col3, Col4)
   Pos1 = InStr(1, sList, "* ", 0)
   Pos2 = InStr(Pos1 + 2, sList, "* ", 0)
   Pos3 = InStr(Pos2 + 2, sList, "* ", 0)
   Pos4 = InStr(Pos3 + 2, sList, "* ", 0)
   Pos5 = InStr(Pos4 + 2, sList, "*", 0)
      
   RecType = Mid(sList, 1, (Pos1 - 1))
   Col1 = Mid(sList, Pos1 + 2, (Pos2 - Pos1 - 2))
   Col2 = Mid(sList, Pos2 + 2, (Pos3 - Pos2 - 2))
   Col3 = Mid(sList, Pos3 + 2, (Pos4 - Pos3 - 2))
   Col4 = Mid(sList, Pos4 + 2, (Pos5 - Pos4 - 2))
End Sub

Private Sub cmdSave_Click()
Dim I As Integer
   Open App.Path & "\tmp.dat" For Output As #30
   For I = 0 To lstSASData.ListCount - 1
      sList = lstSASData.List(I)
      Call ReformData(sList, RecType, Col1, Col2, Col3, Col4)
      Write #30, RecType, Col1, Col2, Col3, Col4
   Next I
   
   Close #30
   
   Open Alg_filename For Output As #1
   Open App.Path & "\tmp.dat" For Input As #30
   Input #30, RecType, Col1, Col2, Col3, Col4
   While Not EOF(30)
      Write #1, RecType, Col1, Col2, Col3, Col4
      Input #30, RecType, Col1, Col2, Col3, Col4
   Wend
' write last record...
   Write #1, RecType, Col1, Col2, Col3, Col4
   Close #1
   Close #30
   Kill App.Path & "\tmp.dat"
   Save_Flag = "N"
End Sub

Private Sub lstSASData_Click()
   cmdReplace.Visible = True
   cmdAdd.Visible = False
   cmdReplace.Default = True
   
   sList = lstSASData.List(lstSASData.ListIndex)
   lstIndex = lstSASData.ListIndex
   Pos1 = InStr(1, sList, "* ", 0)
   
   RecType = Mid(sList, 1, (Pos1 - 1))
   If RecType = "D" Then
      optType(0) = True
      
      Pos2 = InStr(Pos1 + 2, sList, "* ", 0)
      Pos3 = InStr(Pos2 + 2, sList, "* ", 0)
      Pos4 = InStr(Pos3 + 2, sList, "* ", 0)
      Pos5 = InStr(Pos4 + 2, sList, "*", 0)
      txtDescription = Mid(sList, Pos1 + 2, (Pos2 - Pos1 - 2))
      txtAlignment = Mid(sList, Pos2 + 2, (Pos3 - Pos2 - 2))
      txtStation = Mid(sList, Pos3 + 2, (Pos4 - Pos3 - 2))
      comboUnits = Mid(sList, Pos4 + 2, (Pos5 - Pos4 - 2))
      txtDescription.Enabled = True
      txtAlignment.Enabled = True
      txtStation.Enabled = True
      comboUnits.Enabled = True
      If txtPnt1.Visible = True Then
         txtPnt1.Visible = False
         txtPnt2.Visible = False
         txtPnt3.Visible = False
         txtPnt4.Visible = False
         lblPnt1.Visible = False
         lblPnt2.Visible = False
         lblPnt3.Visible = False
         lblPnt4.Visible = False
      End If
   ElseIf RecType = "L" Then
      optType(1) = True
      
      txtDescription.Enabled = False
      txtAlignment.Enabled = False
      txtStation.Enabled = False
      comboUnits.Enabled = False
         
      lblPnt1.Caption = "Beg Pnt"
      lblPnt2.Caption = "End Pnt"
      txtPnt1.Visible = True
      txtPnt2.Visible = True
      txtPnt3.Visible = False
      txtPnt4.Visible = False
      lblPnt1.Visible = True
      lblPnt2.Visible = True
      lblPnt3.Visible = False
      lblPnt4.Visible = False

      frameCurve.Visible = False
      optCurve(0).Visible = False
      optCurve(1).Visible = False
      optCurve(2).Visible = False
      
      Pos2 = InStr(Pos1 + 2, sList, "* ", 0)
      Pos3 = InStr(Pos2 + 2, sList, "* ", 0)
      txtPnt1 = Mid(sList, Pos1 + 2, (Pos2 - Pos1 - 2))
      txtPnt2 = Mid(sList, Pos2 + 2, (Pos3 - Pos2 - 2))
   ElseIf Mid(RecType, 1, 1) = "C" Then
      optType(2) = True
      If Mid(RecType, 2, 1) = "0" Then
         optCurve(0) = True
         lblPnt1.Caption = "PC"
         lblPnt2.Caption = "PI"
         lblPnt3.Caption = "PT"
         txtPnt1.Visible = True
         txtPnt2.Visible = True
         txtPnt3.Visible = True
         txtPnt4.Visible = False
         lblPnt1.Visible = True
         lblPnt2.Visible = True
         lblPnt3.Visible = True
         lblPnt4.Visible = False
      ElseIf Mid(RecType, 2, 1) = "1" Then
         optCurve(1) = True
         lblPnt1.Caption = "PC"
         lblPnt2.Caption = "POC"
         lblPnt3.Caption = "PT"
         txtPnt1.Visible = True
         txtPnt2.Visible = True
         txtPnt3.Visible = True
         txtPnt4.Visible = False
         lblPnt1.Visible = True
         lblPnt2.Visible = True
         lblPnt3.Visible = True
         lblPnt4.Visible = False
      ElseIf Mid(RecType, 2, 1) = "2" Then
         optCurve(2) = True
         lblPnt1.Caption = "PC"
         lblPnt2.Caption = "POST"
         lblPnt3.Caption = "POST"
         lblPnt4.Caption = "PT"
         txtPnt1.Visible = True
         txtPnt2.Visible = True
         txtPnt3.Visible = True
         txtPnt4.Visible = True
         lblPnt1.Visible = True
         lblPnt2.Visible = True
         lblPnt3.Visible = True
         lblPnt4.Visible = True
      End If
      
      Pos2 = InStr(Pos1 + 2, sList, "* ", 0)
      Pos3 = InStr(Pos2 + 2, sList, "* ", 0)
      Pos4 = InStr(Pos3 + 2, sList, "* ", 0)
      Pos5 = InStr(Pos4 + 2, sList, "*", 0)
      txtPnt1 = Mid(sList, Pos1 + 2, (Pos2 - Pos1 - 2))
      txtPnt2 = Mid(sList, Pos2 + 2, (Pos3 - Pos2 - 2))
      txtPnt3 = Mid(sList, Pos3 + 2, (Pos4 - Pos3 - 2))
      txtPnt4 = Mid(sList, Pos4 + 2, (Pos5 - Pos4 - 2))
   End If
End Sub

Private Sub mnuRunAlign_Click()

   For I = 0 To lstSASData.ListCount - 1
      sList = lstSASData.List(I)
      If Mid(sList, 1, 1) = "D" Then
         Call ReformData(sList, RecType, Col1, Col2, Col3, Col4)
         Desc = Col1
         bsta = Str(Col3)
         ESta = bsta
         Align = Str(Col2)
         Write #6, "", "", "", "", "", ""
         Write #6, "", "", "", "", "", ""
         Write #6, "", "", "", "", "", ""
         Write #6, "Description of Alignment", Align, Desc, "", "", ""
         Write #6, "", "", "", "", "", ""
      ElseIf Mid(sList, 1, 1) = "L" Then
         Call ReformData(sList, RecType, Col1, Col2, Col3, Col4)
         NPT = 2
         LDesc(1) = "BP"
         LDesc(2) = "EP"
         IPN(1) = Str(Col1)
         IPN(2) = Str(Col2)
         Call Find_XYZ(IPN, x, y, z, NPT)
         sta(1) = ESta
         sta(2) = sta(1) + Dist(x(1), y(1), x(2), y(2))
         ESta = sta(2)
         dY = y(2) - y(1)
         dX = x(2) - x(1)
         
         CC(1) = dY
         CC(1) = Format(CC(1), "#####0.000")
         CC(2) = -dX
         CC(2) = Format(CC(2), "#####0.000")
         CC(3) = -dY * x(1) + dX * y(1)
         CC(3) = Format(CC(3), "#####0.000")
         
         Call Find_Bearing(x(1), y(1), x(2), y(2), Bearing)
         Length = sta(2) - sta(1)
      
         NumElem = NumElem + 1
         TotElem = TotElem + 1
         AlignArray(NumElem) = Align
         ElemArray(NumElem) = "L"
         DirArray(NumElem) = ""
         For J = 1 To 3
            XArray(NumElem, J) = x(J)
            YArray(NumElem, J) = y(J)
            CArray(NumElem, J) = CC(J)
         Next J
         For J = 1 To 2
            StaArray(NumElem, J) = sta(J)
         Next J
'
' write out data to report here.
         Write #6, "", "", "", "", "", ""
         Write #6, "LINE ELEMENT", "", "", "", "", ""
         Write #6, "", "", "", "", "", ""
         Write #6, "Point"; "X"; "Y"; "Z"; "STA"; ""
         For K = 1 To NPT
            x(K) = Format(x(K), "#####0.000")
            y(K) = Format(y(K), "#####0.000")
            z(K) = Format(z(K), "#####0.000")
            sta(K) = Format(sta(K), "#####0.000")
            Write #6, LDesc(K), x(K), y(K), z(K), sta(K), ""
         Next K
'
         Write #4, Align, ElemArray(NumElem), Desc, _
                   x(1), y(1), x(2), y(2), x(3), y(3), _
                   sta(1), sta(2), sta(3), _
                   CC(1), CC(2), CC(3), _
                   DirArray(NumElem)
'
         Length = Format(Length, "#####0.000")
         Bearing = Format(Bearing, "#####0.000")
         Write #6, "Length=", Length, "Bearing=", Bearing, "", ""
'
      ElseIf Mid(sList, 1, 1) = "C" Then
         Call ReformData(sList, RecType, Col1, Col2, Col3, Col4)
         IPN(1) = Str(Col1)
         IPN(2) = Str(Col2)
         IPN(3) = Str(Col3)
         LDesc(1) = "PC"
         LDesc(2) = "PI"
         LDesc(3) = "PT"
         LDesc(4) = "CC"
         bsta = ESta
' process curve with PC,PI, PT
         If Mid(sList, 2, 1) = "0" Then
            NPT = 3
            Call Find_XYZ(IPN, x, y, z, NPT)
            Call Find_POC(IPN, x, y, NPT, x0, y0, Ang, dir, rad)
' process curve with PC,POC, PT
         ElseIf Mid(sList, 2, 1) = "1" Then
            NPT = 3
            Call Find_XYZ(IPN, x, y, z, NPT)
            Call Find_PI_CC(IPN, x, y, NPT, xPI, yPI, Ang, dir, rad)
            Call Find_POC(IPN, x, y, NPT, x0, y0, Ang, dir, rad)
' process curve with PC,POST, POST, PT
         ElseIf Mid(sList, 2, 1) = "2" Then
            IPN(4) = Str(Col4)
            NPT = 4
            Call Find_XYZ(IPN, x, y, z, NPT)
            Call Find_PI(IPN, x, y)
            Call Find_POC(IPN, x, y, NPT, x0, y0, Ang, dir, rad)
         End If
         TanLen = Dist(x(1), y(1), x(2), y(2))
         TanLen = Format(TanLen, "#####0.000")
         ChdLen = Dist(x(1), y(1), x(3), y(3))
         ChdLen = Format(ChdLen, "#####0.000")
         ExtLen = Dist(x0, y0, x(2), y(2)) - rad
         ExtLen = Format(ExtLen, "#####0.000")
         curveLength = rad * Ang
         curveLength = Format(curveLength, "#####0.000")
         Angle = Rad2Deg(Ang)
         Call DdmmssConv(Angle, Delta)
         sta(1) = ESta
         sta(2) = ESta + Dist(x(1), y(1), x(2), y(2))
         sta(3) = ESta + curveLength
         ESta = sta(3)
         
         CC(1) = x0
         CC(1) = Format(CC(1), "#####0.000")
         CC(2) = y0
         CC(2) = Format(CC(2), "#####0.000")
         CC(3) = rad
         CC(3) = Format(CC(3), "#####0.000")
         
         Angle = 5729.578 / rad
         Call DdmmssConv(Angle, Degree)
         Call Find_Bearing(x(1), y(1), x(2), y(2), Back)
         Call Find_Bearing(x(2), y(2), x(3), y(3), Ahead)
         
         NumElem = NumElem + 1
         TotElem = TotElem + 1
         AlignArray(NumElem) = Align
         ElemArray(NumElem) = "C"
         DirArray(NumElem) = dir
         For J = 1 To 3
            XArray(NumElem, J) = x(J)
            YArray(NumElem, J) = y(J)
            CArray(NumElem, J) = CC(J)
            StaArray(NumElem, J) = sta(J)
         Next J
'
' write out data to report here.
         Write #6, "", "", "", "", "", ""
         Write #6, "CURVE ELEMENT", "", "", "", "", ""
         Write #6, "", "", "", "", "", ""
         Write #6, "Point"; "X"; "Y"; "Z"; "STA"; ""
         For K = 1 To NPT
            x(K) = Format(x(K), "#####0.000")
            y(K) = Format(y(K), "#####0.000")
            z(K) = Format(z(K), "#####0.000")
            sta(K) = Format(sta(K), "#####0.000")
            Write #6, LDesc(K), x(K), y(K), z(K), sta(K), ""
         Next K
'
         Write #4, Align, ElemArray(NumElem), Desc, _
                   x(1), y(1), x(2), y(2), x(3), y(3), _
                   sta(1), sta(2), sta(3), _
                   CC(1), CC(2), CC(3), _
                   DirArray(NumElem)
'
         Write #6, "", "", "", "", "", ""
         Write #6, "Radius=", CC(3), " Delta=", Delta, "", ""
         Write #6, "Length=", curveLength, " Degree=", Degree, "", ""
         Write #6, "Back=", Back, " Ahead=", Ahead, "", ""
         Write #6, "Tangent=", TanLen; " Chord="; ChdLen, "", ""
         Write #6, "External=", ExtLen, "", "", "", ""
'
      End If
   Next I
   Msg = "Successfully completed"
   Style = vbOK
   Title = "Alignment Processing"
   Response = MsgBox(Msg, Style, Title)
   Close #6
   mnuWithSOE.Visible = True
   mnuSasPrint.Visible = True
End Sub

Private Sub mnuWithSOE_Click()

Dim IReturn As Integer

   cmdNewAlign.Visible = False
   cmdAdd.Visible = False
   cmdReplace.Visible = False
   cmdDelete.Visible = False
   cmdSave.Visible = False
   cmdCancel.Visible = False
   cmdQuit.Visible = False
   txtDescription.Visible = False
   txtAlignment.Visible = False
   txtStation.Visible = False
   comboUnits.Visible = False
   txtDescription.Enabled = False
   txtAlignment.Enabled = False
   txtStation.Enabled = False
   comboUnits.Enabled = False
   txtPnt1.Visible = False
   txtPnt2.Visible = False
   txtPnt3.Visible = False
   txtPnt4.Visible = False
   lblDesc.Visible = False
   lblAlign.Visible = False
   lblStation.Visible = False
   lblUnits.Visible = False
   lblPnt1.Visible = False
   lblPnt2.Visible = False
   lblPnt3.Visible = False
   lblPnt4.Visible = False
   lstSASData.Visible = False
   frameType.Visible = False
   frameCurve.Visible = False

   txtBegAlign.Visible = True
   lblBegAlign.Visible = True
   txtGP3Cnt.Visible = True
   lblGP3Cnt.Visible = True
   cmdProcess.Visible = True
   cmdProcess.Default = True
   txtBegAlign.SetFocus

'   Open SasRpt_filename For Append As #6
'   Open XYZ_filename For Output As #7
'   Open GP3_filename For Output As #20
'
'   Open SOE_filename For Input As #8
'   Open SOR_filename For Output As #14
   
   IReturn = Shell("c:\soesort.bat" & " " & SOE_filename & " " & " " & SOR_filename)
 
'   While Not EOF(8)
'      Line Input #8, Record
'      If Mid(Record, 1, 1) <> "*" Then
'         If Mid(Record, 1, 5) <> "     " Then
'            SoePnt = Mid(Record, 1, 5)
'            SoeSta = Mid(Record, 7, 13)
'            SoeOff = Mid(Record, 21, 13)
'            SoeElv = Mid(Record, 35, 13)
'            SoeDesc = Mid(Record, 49, 12)
'            If SoeOff < 0 Then
'               SoeOff = 500000000 - SoeOff
'            ElseIf SoeOff = 0 Then
'               SoeOff = 100000000
'            ElseIf SoeOff > 0 Then
'               SoeOff = SoeOff + 200000000
'            End If
'            Entry = SoeSta & "*" & SoeOff & "*" & SoePnt & "*" & SoeElv & "*" & SoeDesc & "*"
'            lstSOESort.AddItem Entry
'         End If
'      End If
'   Wend
'   Close #8
'   For Cnt = 0 To lstSOESort.ListCount - 1
'      sList = lstSOESort.List(Cnt)
'      Pos1 = InStr(1, sList, "*", 0)
'      Pos2 = InStr(Pos1 + 1, sList, "*", 0)
'      Pos3 = InStr(Pos2 + 1, sList, "*", 0)
'      Pos4 = InStr(Pos3 + 1, sList, "*", 0)
'      Pos5 = InStr(Pos4 + 1, sList, "*", 0)
'      SoeSta = Mid(sList, 1, Pos1 - 1)
'      SoeOff1 = Mid(sList, Pos1 + 1, Pos2 - Pos1 - 1)
'
'         If SoeOff1 > 500000000 Then
'            SoeOff = (SoeOff1 - 500000000) * -1
'         ElseIf SoeOff1 = 100000000 Then
'            SoeOff = 0
'         ElseIf SoeOff1 > 200000000 And SoeOff1 < 300000000 Then
'            SoeOff = SoeOff1 - 200000000
'         End If
'      SoeOff = Format(SoeOff, "######0.00000")
'      SoePnt = Mid(sList, Pos2 + 1, Pos3 - Pos2 - 1)
'      SoeElv = Mid(sList, Pos3 + 1, Pos4 - Pos3 - 1)
'      SoeDesc = Mid(sList, Pos4 + 1, Pos5 - Pos4 - 1)
'      Write #14, SoePnt, SoeSta, SoeOff, SoeElv, SoeDesc
'   Next Cnt
'   Close #14
'
End Sub

Private Sub cmdProcess_Click()
'
   Open SasRpt_filename For Append As #6
   Open XYZ_filename For Output As #7
   Open GP3_filename For Output As #20
'
   Open SO1_filename For Output As #10
   Open SOR_filename For Input As #14
   
   Write #6, "", "", "", "", "", ""
   Write #6, "", "", "", "", "", ""
   Write #6, "XYZ Conversion Report", "", "", "", "", ""
   Write #6, "X Coord", "Y Coord", "Z Coord", "Station", "Offset", "Elevation "
   
   If txtGP3Cnt = "" Then
      GP3Cnt = 1000
   Else
      GP3Cnt = Val(txtGP3Cnt)
   End If
   For Cnt = 1 To TotElem
      If Val(txtBegAlign) = AlignArray(Cnt) Then
         IC = Cnt
         While Not EOF(14)
            Line Input #14, Record
            If Mid(Record, 1, 1) <> "*" Then
'
               Pos1 = InStr(1, Record, ",", 0)
               Pos2 = InStr(Pos1 + 1, Record, ",", 0)
               Pos3 = InStr(Pos2 + 1, Record, ",", 0)
               Pos4 = InStr(Pos3 + 1, Record, ",", 0)
      
               SoePnt = Mid(Record, 1, Pos1 - 1)
               SoeSta = Mid(Record, Pos1 + 1, Pos2 - Pos1 - 1)
               SoeOff = Mid(Record, Pos2 + 1, Pos3 - Pos2 - 1)
               SoeElv = Mid(Record, Pos3 + 1, Pos4 - Pos3 - 1)
               SoeDesc = Mid(Record, Pos4 + 1, Len(Record) - Pos4)
'
200            If ElemArray(IC) = "C" Then
                  StaDist = StaArray(IC, 3) - SoeSta
                  If StaArray(IC, 3) > SoeSta Or Abs(StaDist) <= 1 Then
                     If SoeOff = 0 Then
                        CenterElev = SoeElv
                        Call curveStat(CArray(IC, 1), CArray(IC, 2), XArray(IC, 1), YArray(IC, 1), _
                           StaArray(IC, 1), SoeSta, DirArray(IC), CArray(IC, 3), xc, yc)
                     End If
                     Call xyzCoord(xc, yc, CenterElev, CArray(IC, 1), CArray(IC, 2), _
                        XArray(IC, 2), YArray(IC, 2), offset, elev, ElemArray(IC), _
                        DirArray(IC), xx, yy, zz)
                     Call writeRtn(xx, yy, zz, SoeSta, SoeOff, SoeElv, GP3Cnt)
                     GP3Cnt = GP3Cnt + 1
                  ElseIf IC < TotElem And StaArray(IC, 3) <= SoeSta Then
                     IC = IC + 1
                     GoTo 200
                  ElseIf IC >= TotElem And StaArray(IC, 3) <= SoeSta Then
                     IC = Cnt
                  End If
               ElseIf ElemArray(IC) = "L" Then
                  StaDist = StaArray(IC, 2) - SoeSta
                  If StaArray(IC, 2) > SoeSta Or Abs(StaDist) <= 1 Then
                     If SoeOff = 0 Then
                        CenterElev = SoeElv
                        Call lineStat(XArray(IC, 1), XArray(IC, 2), YArray(IC, 1), YArray(IC, 2), _
                           StaArray(IC, 1), SoeSta, xs, ys)
                     End If
                     Call xyzCoord(xs, ys, CenterElev, XArray(IC, 1), YArray(IC, 1), _
                        XArray(IC, 2), YArray(IC, 2), SoeOff, SoeElv, ElemArray(IC), _
                        DirArray(IC), xx, yy, zz)
                     Call writeRtn(xx, yy, zz, SoeSta, SoeOff, SoeElv, GP3Cnt)
                     GP3Cnt = GP3Cnt + 1
                  ElseIf IC < TotElem And StaArray(IC, 2) <= SoeSta Then
                     IC = IC + 1
                     GoTo 200
                  ElseIf IC >= TotElem And StaArray(IC, 2) <= SoeSta Then
                     IC = Cnt
                  End If
               End If
            End If
         Wend
      End If
   Next Cnt
   
   txtBegAlign.Visible = False
   lblBegAlign.Visible = False
   txtGP3Cnt.Visible = False
   lblGP3Cnt.Visible = False
   cmdProcess.Visible = False
   
   cmdNewAlign.Visible = True
   cmdAdd.Visible = True
   cmdReplace.Visible = True
   cmdDelete.Visible = True
   cmdSave.Visible = True
   cmdCancel.Visible = True
   cmdQuit.Visible = True
   txtDescription.Visible = True
   txtAlignment.Visible = True
   txtStation.Visible = True
   txtDescription.Enabled = True
   txtAlignment.Enabled = True
   txtStation.Enabled = True
   txtPnt1.Visible = True
   txtPnt2.Visible = True
   txtPnt3.Visible = True
   txtPnt4.Visible = True
   lblDesc.Visible = True
   lblAlign.Visible = True
   lblStation.Visible = True
   lblPnt1.Visible = True
   lblPnt2.Visible = True
   lblPnt3.Visible = True
   lblPnt4.Visible = True
   lstSASData.Visible = True
   frameType.Visible = True
   
   Close #6
   Close #10
   Close #14
End Sub

Private Sub optType_Click(Index As Integer)
   Select Case Index
      Case 0
         cmdNewAlign.Visible = True
   
         lblAlign.Visible = True
         lblDesc.Visible = True
         lblStation.Visible = True
         lblUnits.Visible = True
         comboUnits.Visible = True
         
         txtDescription.Visible = True
         txtAlignment.Visible = True
         txtStation.Visible = True
         txtDescription.Enabled = True
         txtAlignment.Enabled = True
         txtStation.Enabled = True
         
         txtDescription.SetFocus
         
         txtPnt1.Visible = False
         txtPnt2.Visible = False
         txtPnt3.Visible = False
         txtPnt4.Visible = False
         lblPnt1.Visible = False
         lblPnt2.Visible = False
         lblPnt3.Visible = False
         lblPnt4.Visible = False

         frameCurve.Visible = False
         optCurve(0).Visible = False
         optCurve(1).Visible = False
         optCurve(2).Visible = False
      Case 1
         
         txtDescription.Enabled = False
         txtAlignment.Enabled = False
         txtStation.Enabled = False
         
         lblPnt1.Caption = "Beg Pnt"
         lblPnt2.Caption = "End Pnt"
         txtPnt1.Visible = True
         txtPnt2.Visible = True
         txtPnt3.Visible = False
         txtPnt4.Visible = False
         lblPnt1.Visible = True
         lblPnt2.Visible = True
         lblPnt3.Visible = False
         lblPnt4.Visible = False

         txtPnt1.SetFocus

         frameCurve.Visible = False
         optCurve(0).Visible = False
         optCurve(1).Visible = False
         optCurve(2).Visible = False
      Case 2
         
         txtDescription.Enabled = False
         txtAlignment.Enabled = False
         txtStation.Enabled = False
         
         txtPnt1.Visible = False
         txtPnt2.Visible = False
         txtPnt3.Visible = False
         txtPnt4.Visible = False
         lblPnt1.Visible = False
         lblPnt2.Visible = False
         lblPnt3.Visible = False
         lblPnt4.Visible = False
         frameCurve.Visible = True
         optCurve(0) = False
         optCurve(1) = False
         optCurve(2) = False
         optCurve(0).Visible = True
         optCurve(1).Visible = True
         optCurve(2).Visible = True
   End Select
End Sub

Private Sub optCurve_Click(Index As Integer)
   txtPnt1 = ""
   txtPnt2 = ""
   txtPnt3 = ""
   txtPnt4 = ""
   Select Case Index
      Case 0
         lblPnt1.Caption = "PC"
         lblPnt2.Caption = "PI"
         lblPnt3.Caption = "PT"
         txtPnt1.Visible = True
         txtPnt2.Visible = True
         txtPnt3.Visible = True
         txtPnt4.Visible = False
         lblPnt1.Visible = True
         lblPnt2.Visible = True
         lblPnt3.Visible = True
         lblPnt4.Visible = False
      Case 1
         lblPnt1.Caption = "PC"
         lblPnt2.Caption = "POC"
         lblPnt3.Caption = "PT"
         txtPnt1.Visible = True
         txtPnt2.Visible = True
         txtPnt3.Visible = True
         txtPnt4.Visible = False
         lblPnt1.Visible = True
         lblPnt2.Visible = True
         lblPnt3.Visible = True
         lblPnt4.Visible = False
      Case 2
         lblPnt1.Caption = "PC"
         lblPnt2.Caption = "POST1"
         lblPnt3.Caption = "POST2"
         lblPnt4.Caption = "PT"
         txtPnt1.Visible = True
         txtPnt2.Visible = True
         txtPnt3.Visible = True
         txtPnt4.Visible = True
         lblPnt1.Visible = True
         lblPnt2.Visible = True
         lblPnt3.Visible = True
         lblPnt4.Visible = True
   End Select
End Sub


Private Sub mnuSasPrint_Click()

   Msg = "Continue Printing"
   Style = vbYesNo
   Title = "Print Proc"
   Response = MsgBox(Msg, Style, Title)
   
   Open SasRpt_filename For Input As #6
   
   If Response = vbNo Then
      Printer.KillDoc ' Terminate print job abruptly.
      Printer.EndDoc
   Else
      If Response = vbYes Then
         Printer.Orientation = 1       ' Portrait orientation
         Printer.FontSize = 10
         Printer.FontName = Screen.Fonts(12)
         Printer.FontSize = 10
         Page = 1
         Call Header(Page)
         
         IRow = 9
         xyzSwitch = "N"
         While Not EOF(6)
            Input #6, F1, F2, F3, F4, F5, F6
            IRow = IRow + 1
            If IRow > 60 Then
               IRow = 9
               Printer.NewPage
               Page = Page + 1
               Call Header(Page)
               If xyzSwitch = "Y" Then
                  F1a = "X Coord"
                  F2a = "Y Coord"
                  F3a = "Z Coord"
                  F4a = "Station"
                  F5a = "Offset"
                  F6a = "Elevation"
                  Printer.Print F1a, F2a, F3a, F4a, F5a, F6a
               End If
            ElseIf F1 = "LINE ELEMENT" And IRow > 54 Then
               IRow = 9
               Printer.NewPage
               Page = Page + 1
               Call Header(Page)
            ElseIf F1 = "CURVE ELEMENT" And IRow > 47 Then
               IRow = 9
               Printer.NewPage
               Page = Page + 1
               Call Header(Page)
            ElseIf F1 = "XYZ Conversion Report" Then
               xyzSwitch = "Y"
            End If
            Printer.Print F1, F2, F3, F4, F5, F6
         Wend
         
         Printer.EndDoc
         Msg = "Finished SAS Print"
         Style = vbOKOnly
         Title = "Print Proc"
         Response = MsgBox(Msg, Style, Title)
      End If
   End If
End Sub
Private Sub Header(Page)
   
   MyDate = "Date:" & Now
   MyPage = "Page:" & Page
   Printer.Print "                                                                                "; MyPage
   Printer.FontBold = True
   Printer.Print ""
   Printer.Print "                            SOUTH CAROLINA DEPARTMENT OF TRANSPORTATION"
   Printer.Print ""
   Printer.Print "                                     SURVEY AUTOMATION SYSTEM"
   Printer.Print ""
   Printer.Print "                         VERSION:WIN/NT 2.0                         "; MyDate
   Printer.Print ""
   Printer.FontBold = False
   Printer.Print "===================================================================================================="
End Sub

