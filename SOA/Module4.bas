Attribute VB_Name = "Soa"
Option Explicit

Public Sub CreateMap()

   Open New_filename For Append As #2
   Open Map_filename For Output As #3
   MySize = LOF(2)
   Close #2
   If MySize = 0 Then
      Msg = "Missing New File"
      Style = vbOK
      Title = "Critical Error"
      Response = MsgBox(Msg, Style, Title)
      Unload frmSoaInp
   Else
      Open New_filename For Input As #2
      Line Input #2, Record
      
      While Not EOF(2)
         If Mid(Record, 1, 1) <> "*" And _
            Mid(Record, 2, 1) <> "*" And _
            Mid(Record, 1, 5) <> "     " Then
            Pnt = Str(Mid(Record, 1, 5))
            iY = Str(Mid(Record, 7, 13))
            iX = Str(Mid(Record, 21, 13))
            iZ = Str(Mid(Record, 35, 13))
            Desc = Mid(Record, 49, 20)
         
'write  #3 .map file
            If Mid(Desc, 1, 2) = "CL" Or _
               Mid(Desc, 1, 3) = "X" Then
            Else
               Write #3, Pnt, iX, iY, iZ, Desc
            End If
         End If
         Line Input #2, Record
      Wend
'last record
      If Mid(Record, 1, 1) <> "*" And _
         Mid(Record, 2, 1) <> "*" And _
         Mid(Record, 1, 5) <> "     " Then
         Pnt = Str(Mid(Record, 1, 5))
         iY = Str(Mid(Record, 7, 13))
         iX = Str(Mid(Record, 21, 13))
         iZ = Str(Mid(Record, 35, 13))
         Desc = Mid(Record, 49, 20)
         
'write  #3 .map file
         If Mid(Desc, 1, 2) = "CL" Or _
            Mid(Desc, 1, 3) = "X" Then
         Else
            Write #3, Pnt, iX, iY, iZ, Desc
         End If
      End If
      Msg = "Map file created"
      Style = vbOK
      Title = "Create Map"
      Response = MsgBox(Msg, Style, Title)
      Close #2
      Close #3
   End If
End Sub

Public Sub CreateIdx()
   
   For I = 1 To 1000
      idx_Pnt(idxCnt) = Pnt
      idx_X(idxCnt) = 0
      idx_Y(idxCnt) = 0
      idx_Z(idxCnt) = 0
   Next I
   Open New_filename For Input As #2
   Line Input #2, Record
   While Not EOF(2)
      If Mid(Record, 1, 1) <> "*" Then
         Pnt = Str(Mid(Record, 1, 5))
         Desc = Mid(Record, 49, 20)
         iY = Str(Mid(Record, 7, 13))
         iX = Str(Mid(Record, 21, 13))
         iZ = Str(Mid(Record, 35, 13))
         
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
            Mid(Record, 49, 3) = "ICL" Or _
            Mid(Record, 49, 2) = "CP" Or _
            Mid(Record, 49, 3) = "ODL" Or _
            Mid(Record, 49, 4) = "ODLA" Or _
            Mid(Record, 49, 4) = "ODLB" Or _
            Mid(Record, 49, 2) = "BM" And _
            Pnt > 0 Then
            
            idxCnt = idxCnt + 1
            idx_Pnt(idxCnt) = Pnt
            idx_X(idxCnt) = iX
            idx_Y(idxCnt) = iY
            idx_Z(idxCnt) = iZ
            Write #9, Pnt, iX, iY, iZ, Desc, ""
         End If
      End If
      Line Input #2, Record
   Wend
   Close #2
End Sub

Public Sub SoaLine(x1, y1, X2, Y2, A, B, C, bsta, LOff, ROff, BegPnt, EndPnt, BegSta, _
                   EndSta, AlignNum, AlignDesc)
Dim xx As Double
Dim yy As Double
Dim zz As Double
   pntFlag = "Y"
   staFlag = "Y"
   If BegPnt = 0 And EndPnt = 0 Then
      pntFlag = "N"
   End If
   If BegSta = 0 And EndSta = 0 Then
      staFlag = "N"
   End If
   If A = 0 Then
      swCase = "1"
   ElseIf B = 0 Then
      swCase = "2"
   Else
      swCase = "3"
   End If
   For I = 1 To 1000
      rejPnt(I) = 0
      rejY(I) = 0
      rejX(I) = 0
      rejZ(I) = 0
      rejDesc(I) = ""
   Next I
   Open Rej_filename For Input As #11
   Do While Not EOF(11)
      pntCnt = pntCnt + 1
      validPnt(pntCnt) = "N"
      Input #11, Pnt, yy, xx, zz, Desc
      rejPnt(pntCnt) = Pnt
      rejY(pntCnt) = yy
      rejX(pntCnt) = xx
      rejZ(pntCnt) = zz
      rejDesc(pntCnt) = Desc
   Loop
   Close #11
   For I = 1 To pntCnt
      If pntFlag = "Y" Then
         If rejPnt(I) < BegPnt Or rejPnt(I) > EndPnt Then
            GoTo 49
         End If
      End If
      chk1X = "N"
      chk1Y = "N"
      If swCase = 1 Then
         lineX(I) = rejX(I)
         lineY(I) = -(C / B)
      ElseIf swCase = 2 Then
         lineX(I) = -(C / A)
         lineY(I) = rejY(I)
      Else
         lineX(I) = ((B / A) * rejX(I) - rejY(I) - (C / B)) / (A / B + B / A)
         lineY(I) = -(C / B) - (A / B) * lineX(I)
      End If
      
      result1 = x1
      If X2 < x1 Then    ' Min processing
         result1 = X2
      End If
      
      result2 = x1
      If X2 > x1 Then    ' Max processing
         result2 = X2
      End If
      res1 = result1 - 0.1
      res2 = result2 + 0.1
      If res1 <= lineX(I) And lineX(I) <= res2 Then
         chk1X = "Y"
      End If

      result1 = y1
      If Y2 < y1 Then
         result1 = Y2
      End If
      result2 = y1
      If Y2 > y1 Then
         result2 = Y2
      End If
      res1 = result1 - 0.1
      res2 = result2 + 0.1
      If res1 <= lineY(I) And lineY(I) <= res2 Then
         chk1Y = "Y"
      End If
      
      If chk1X = "Y" And chk1Y = "Y" Then
         Off(I) = Dist(rejX(I), rejY(I), lineX(I), lineY(I))
      Else
         GoTo 49
      End If
      ProjDist(I) = Dist(x1, y1, lineX(I), lineY(I))
      OffFlag = dirLR(X2, Y2, x1, y1, rejX(I), rejY(I))
      Off(I) = OffFlag * Off(I)
      
      If Off(I) < (ROff + 0.1) And Off(I) > (LOff - 0.1) Then
         OffSta(I) = BegSta + ProjDist(I)
      Else
         GoTo 49
      End If
      If staFlag = "Y" Then
         If OffSta(I) < BegSta Or OffSta(I) > EndSta Then
         GoTo 49
         End If
      End If
      validPnt(I) = "Y"
      Write #12, rejPnt(I), rejX(I), rejY(I), rejZ(I), rejDesc(I), _
                 OffSta(I), Off(I), AlignNum, AlignDesc
49    If validPnt(I) = "N" Then
         Write #22, rejPnt(pntCnt), rejY(pntCnt), rejX(pntCnt), rejZ(pntCnt), _
                    rejDesc(pntCnt)
      End If
   Next I
   pntCnt = 0
   
End Sub

Public Sub SoaCurve(x1, y1, X3, Y3, bsta, C1, C2, Radius, dir, LOff, ROff, BegPnt, _
                    EndPnt, BegSta, EndSta, AlignNum, AlignDesc)

Dim xx As Double
Dim yy As Double
Dim zz As Double
   
   pntFlag = "Y"
   staFlag = "Y"
   If BegPnt = 0 And EndPnt = 0 Then
      pntFlag = "N"
   End If
   If BegSta = 0 And EndSta = 0 Then
      staFlag = "N"
   End If
   If A = 0 Then
      swCase = "1"
   ElseIf B = 0 Then
      swCase = "2"
   Else
      swCase = "3"
   End If
   For I = 1 To 1000
      rejPnt(I) = 0
      rejY(I) = 0
      rejX(I) = 0
      rejZ(I) = 0
      rejDesc(I) = ""
      validPnt(I) = 0
   Next I
   
   Call curveDir(x1, y1, C1, C2, rejX(I), rejY(I), theta, TDir)
   theta1 = theta
   dir = TDir
   
   Open Rej_filename For Input As #11
   Do While Not EOF(11)
      pntCnt = pntCnt + 1
      validPnt(pntCnt) = "N"
      Input #11, Pnt, yy, xx, zz, Desc
      rejPnt(pntCnt) = Pnt
      rejY(pntCnt) = yy
      rejX(pntCnt) = xx
      rejZ(pntCnt) = zz
      rejDesc(pntCnt) = Desc
   Loop
   Close #11
   For I = 1 To pntCnt
      If pntFlag = "Y" Then
         If rejPnt(I) < BegPnt Or rejPnt(I) > EndPnt Then
            If validPnt(I) = "N" Then
               GoTo 49
            End If
         End If
      End If
      
      ChkOff(I) = Dist(C1, C2, rejX(I), rejY(I))
      If ChkOff(I) > (ROff + 0.1) Or ChkOff(I) < (LOff - 0.1) Then
         ChkOff(I) = 0
         Call tmpWrite(I)
      End If
      
      Call curveDir(x1, y1, C1, C2, rejX(I), rejY(I), theta, TDir)
      
      Theta4(I) = theta
      TDir4(I) = TDir
      
      If Theta4(I) <= theta1 And TDir4(I) = dir Then
         Off(I) = Dist(C1, C2, rejX(I), rejY(I)) - Radius
      Else
         GoTo 49
      End If
      OffFlag = dirLR(X2, Y2, x1, y1, rejX(I), rejY(I))
      Off(I) = OffFlag * Off(I)
      
      If Off(I) < (ROff + 0.1) And Off(I) > (LOff - 0.1) Then
         ProjDist(I) = Radius * Theta4(I)
         OffSta(I) = BegSta + ProjDist(I)
      Else
         GoTo 49
      End If
      
      If staFlag = "Y" Then
         If OffSta(I) < BegSta Or OffSta(I) > EndSta Then
            GoTo 49
         End If
      End If
      validPnt(I) = "Y"
      Write #12, rejPnt(I), rejX(I), rejY(I), rejZ(I), rejDesc(I), OffSta(I), _
                 Off(I), AlignNum, AlignDesc
49    If validPnt(I) = "N" Then
         Write #22, rejPnt(pntCnt), rejY(pntCnt), rejX(pntCnt), rejZ(pntCnt), _
                    rejDesc(pntCnt)
      End If
   Next I
   pntCnt = 0
 End Sub


Private Sub curveDir(x1, y1, C1, C2, rejX, rejY, theta, TDir)
   oa = (x1 - C1) * (rejX - C1)
   ob = (y1 - C2) * (rejY - C2)
   doa = Dist(x1, y1, C1, C2)
   dob = Dist(C1, C2, rejX, rejY)
   If doa = 0 Or dob = 0 Then
      theta = 0
   Else
      arg = (oa + ob) / (doa * dob)
      If Abs(1 - arg) < 0.000001 Then
         arg = 1
      End If
      If arg = 1 Then
         theta = 0
      Else
         theta = PI / 2 - Atn(arg / Sqr(1 - arg ^ 2))
      End If
   End If
   dir = dirLR(x1, y1, C1, C2, rejX, rejY)
   If dir > 0 Then
      TDir = "L"
   Else
      TDir = "R"
   End If
End Sub

Public Function dirLR(x1, y1, X2, Y2, X3, Y3)
   AXB = ((x1 - X2) * (Y3 - Y2)) - ((X3 - X2) * (y1 - Y2))
   If AXB = 0 Then
      dirLR = 0
   Else
      dirLR = -1 * (AXB / Abs(AXB))
   End If
End Function

Public Sub Output()
Dim parms As String
   
   Open Vld_filename For Input As #12
   Open Rt_filename For Output As #24
   Open Lt_filename For Output As #25
'
' list box used to sort file by align no. station then offset.
'
   While Not EOF(12)
      Input #12, Pnt, X3, Y3, Z3, Desc, SoeSta, offset, Align, AlignDesc
      If offset < 0 Then
         offset = offset + 500000000
      ElseIf offset = 0 Then
         offset = 100000000
      ElseIf offset > 0 Then
         offset = offset + 200000000
      End If
      Entry = Align & "*" & offset & "*" & SoeSta & "*" & Pnt & "*" & X3 & "*" & Y3 & "*" & Z3 & "*" & Desc & "*" & AlignDesc & "*"
      frmSoaInp.lstSoaSort.AddItem Entry
   Wend
   Close #12
   For Cnt = 0 To frmSoaInp.lstSoaSort.ListCount - 1
      sList = frmSoaInp.lstSoaSort.List(Cnt)
      Pos1 = InStr(1, sList, "*", 0)
      Pos2 = InStr(Pos1 + 1, sList, "*", 0)
      Pos3 = InStr(Pos2 + 1, sList, "*", 0)
      Pos4 = InStr(Pos3 + 1, sList, "*", 0)
      Pos5 = InStr(Pos4 + 1, sList, "*", 0)
      Pos6 = InStr(Pos5 + 1, sList, "*", 0)
      Pos7 = InStr(Pos6 + 1, sList, "*", 0)
      Pos8 = InStr(Pos7 + 1, sList, "*", 0)
      Pos9 = InStr(Pos8 + 1, sList, "*", 0)
      
      Align = Mid(sList, 1, Pos1 - 1)
      SoeOff = Mid(sList, Pos1 + 1, Pos2 - Pos1 - 1)
      SoeSta = Mid(sList, Pos2 + 1, Pos3 - Pos2 - 1)
      
      If SoeOff > 400000000 And SoeOff < 500000000 Then
         SoeOff = SoeOff - 500000000
      ElseIf SoeOff = 100000000 Then
         SoeOff = 0
      ElseIf SoeOff > 200000000 And SoeOff < 300000000 Then
         SoeOff = SoeOff - 200000000
      End If
      SoeOff = Format(SoeOff, "########0.00")
      
      Pnt = Mid(sList, Pos3 + 1, Pos4 - Pos3 - 1)
      X3 = Mid(sList, Pos4 + 1, Pos5 - Pos4 - 1)
      Y3 = Mid(sList, Pos5 + 1, Pos6 - Pos5 - 1)
      Z3 = Mid(sList, Pos6 + 1, Pos7 - Pos6 - 1)
      Desc = Mid(sList, Pos7 + 1, Pos8 - Pos7 - 1)
      AlignDesc = Mid(sList, Pos8 + 1, Pos9 - Pos8 - 1)
      If SoeOff < 0 Then
         Write #25, SoeSta, SoeOff, Desc, Align, AlignDesc
      ElseIf SoeOff > 0 Then
         Write #24, SoeSta, SoeOff, Desc, Align, AlignDesc
      End If
   Next Cnt
   Close #24
   Close #25
'
   frmSoaInp.lstSoaSort.Clear
'
   Open Rt_filename For Input As #24
   Open DgnRt_filename For Output As #34
   While Not EOF(24)
      Input #24, SoeSta, SoeOff, Desc, Align, AlignDesc
      Entry = Align & "*" & Format(SoeSta, "00000.00") & "*" & SoeOff & "*" & Desc & "*" & AlignDesc & "*"
      frmSoaInp.lstSoaSort.AddItem Entry
   Wend
   Close #24
'
   Print #34, ".LV=", 24
   Print #34, ".TH=", wHight
   Print #34, ".TW=", wWidth
   Print #34, ".FT=", wFont
   Print #34, ".LS=", 1.5
   Print #34, ".CO=", 1
   Print #34, ""
'
   prevAlign = 0
   For Cnt = 0 To frmSoaInp.lstSoaSort.ListCount - 1
      sList = frmSoaInp.lstSoaSort.List(Cnt)
      Pos1 = InStr(1, sList, "*", 0)
      Pos2 = InStr(Pos1 + 1, sList, "*", 0)
      Pos3 = InStr(Pos2 + 1, sList, "*", 0)
      Pos4 = InStr(Pos3 + 1, sList, "*", 0)
      Pos5 = InStr(Pos4 + 1, sList, "*", 0)
      
      Align = Mid(sList, 1, Pos1 - 1)
      SoeSta = Mid(sList, Pos1 + 1, Pos2 - Pos1 - 1)
      SoeOff = Mid(sList, Pos2 + 1, Pos3 - Pos2 - 1)
      Desc = Mid(sList, Pos3 + 1, Pos4 - Pos3 - 1)
      AlignDesc = Mid(sList, Pos4 + 1, Pos5 - Pos4 - 1)
      If prevAlign <> Align Then
         Print #34, ".GG"
         Print #34, "           RIGHT        "
         Print #34, AlignDesc
         Print #34, " Station-Offset   Description"
         Print #34, "------------------------------"
         Print #34, ".TH=", wHight
         Print #34, ".TW=", wWidth
         Print #34, ".FT=", wFont
         Print #34, ".LS=", 1
      End If
      If Units = "Metric" Then
         CSta = Int(SoeSta) & "+" & Format(SoeSta - Int(SoeSta), "########0.00") & "-" & SoeOff & "m" & " " & Desc
      Else
         CSta = Int(SoeSta) & "+" & Format(SoeSta - Int(SoeSta), "########0.00") & "-" & SoeOff & "'" & " " & Desc
      End If
      Print #34, CSta
      prevAlign = Align
   Next Cnt
   Close #34
'
   frmSoaInp.lstSoaSort.Clear
   Open Lt_filename For Input As #25
   Open DgnLt_filename For Output As #35
   While Not EOF(25)
      Input #25, SoeSta, SoeOff, Desc, Align, AlignDesc
      SoeSta = Format(SoeSta, "00000.00")
      Entry = Align & "*" & Format(SoeSta, "00000.00") & "*" & SoeOff & "*" & Desc & "*" & AlignDesc & "*"
      frmSoaInp.lstSoaSort.AddItem Entry
   Wend
   Close #25
'
   Print #35, ".LV=", 24
   Print #35, ".TH=", wHight
   Print #35, ".TW=", wWidth
   Print #35, ".FT=", wFont
   Print #35, ".LS=", 1.5
   Print #35, ".CO=", 1
   Print #35, ""
'
   prevAlign = 0
   For Cnt = 0 To frmSoaInp.lstSoaSort.ListCount - 1
      sList = frmSoaInp.lstSoaSort.List(Cnt)
      Pos1 = InStr(1, sList, "*", 0)
      Pos2 = InStr(Pos1 + 1, sList, "*", 0)
      Pos3 = InStr(Pos2 + 1, sList, "*", 0)
      Pos4 = InStr(Pos3 + 1, sList, "*", 0)
      Pos5 = InStr(Pos4 + 1, sList, "*", 0)
      
      Align = Mid(sList, 1, Pos1 - 1)
      SoeSta = Mid(sList, Pos1 + 1, Pos2 - Pos1 - 1)
      SoeOff = Mid(sList, Pos2 + 1, Pos3 - Pos2 - 1)
      Desc = Mid(sList, Pos3 + 1, Pos4 - Pos3 - 1)
      AlignDesc = Mid(sList, Pos4 + 1, Pos5 - Pos4 - 1)
      If prevAlign <> Align Then
         Print #35, ".GG"
         Print #35, "           LEFT           "
         Print #35, AlignDesc
         Print #35, " Station-Offset   Description"
         Print #35, "------------------------------"
         Print #35, ".TH=", wHight
         Print #35, ".TW=", wWidth
         Print #35, ".FT=", wFont
         Print #35, ".LS=", 1
      End If
      If Units = "Metric" Then
         CSta = Int(SoeSta) & "+" & Format(SoeSta - Int(SoeSta), "########0.00") & "-" & (SoeOff * -1) & "m" & " " & Desc
      Else
         CSta = Int(SoeSta) & "+" & Format(SoeSta - Int(SoeSta), "########0.00") & "-" & (SoeOff * -1) & "'" & " " & Desc
      End If
      Print #35, CSta
      prevAlign = Align
   Next Cnt
   Close #35
'
   Kill Rt_filename
   Kill Lt_filename
'
   Msg = "Finished Creating DGN Files"
   Style = vbOKOnly
   Title = "Print Proc"
   Response = MsgBox(Msg, Style, Title)
   frmSoaInp.mnuPrint.Enabled = True
End Sub









