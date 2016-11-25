VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmEbs 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EarthWork Balancing System"
   ClientHeight    =   4725
   ClientLeft      =   7755
   ClientTop       =   6540
   ClientWidth     =   7425
   Icon            =   "frmEbs.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4725
   ScaleWidth      =   7425
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   315
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      HelpFile        =   """C:\SOURCE\ebs\ebs.hlp"""
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuNewFile 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuDataProc 
      Caption         =   "&Data Proc"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmEbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim PCVol As Double
Dim PFVol As Double
Dim GCVol As Double
Dim GFVol As Double

Dim Msg, style, title, response

Private Sub dirFileOpen_Change()
   fileOpen.Path = dirFileOpen.Path
   ChDir dirFileOpen.Path
End Sub

Private Sub drvFileOpen_Change()
   dirFileOpen.Path = drvFileOpen.Drive
   ChDrive drvFileOpen.Drive
End Sub

Private Sub Form_Load()
'
   App.HelpFile = App.Path & "\ebs.hlp"
'
   mnuDataProc.Enabled = False
   mnuEdit.Enabled = False
   mnuPrint.Enabled = False
End Sub

Private Sub mnuDataProc_Click()
   frmInitDataEntry.Show vbModal, Me
End Sub

Private Sub mnuEdit_Click()
   Call Shell("notepad " & dlg1.filename, vbNormalFocus)
End Sub

Private Sub mnuExit_Click()
   End
End Sub

Private Sub mnuFileOpen_Click()
Dim sFile As String
   With dlg1
      .Filter = "CNF Files (*.CNF)|*.cnf"
      .ShowOpen
      If Len(.filename) = 0 Then
         Exit Sub
      End If
      sFile = .filename
   End With
   mnuDataProc.Enabled = True
   mnuEdit.Enabled = True
   mnuPrint.Enabled = True
End Sub

Private Sub mnuHelp_Click()
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, HelpConstants.cdlHelpContents, 0)

End Sub

Private Sub mnuNewFile_Click()
    ' Set CancelError is True
    dlg1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    dlg1.Flags = cdlOFNHideReadOnly
    ' Set filters
    dlg1.Filter = "All Files (*.*)|*.*|Ebs Files (*.cnf)|*.cnf"
    ' Specify default filter
    dlg1.FilterIndex = 2
    ' Display the Open dialog box
    dlg1.ShowOpen
    
   mnuDataProc.Enabled = True
   mnuEdit.Enabled = True
   mnuPrint.Enabled = True
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub mnuPrint_Click()
Dim Units As String
Dim County As String
Dim Route As String
Dim Comments As String
Dim Shrink As String

Dim F1 As String
Dim F2 As String
Dim F3 As String
Dim F4 As String
Dim F5 As String
Dim F6 As String
Dim F7 As String
Dim F8 As String
Dim F9 As String
Dim F10 As String
Dim F11 As String
Dim F12 As String
Dim IRow As Integer
   
   
   Open dlg1.filename For Input As #1
   Input #1, Units, County, Route, Comments, Shrink
   Close #1
   IPos = InStr(1, dlg1.filename, ".", 0)
   Rpt_filename = Mid(dlg1.filename, 1, IPos) & "Ewk"
   Open Rpt_filename For Input As #7
'
   Msg = "Continue Printing"
   style = vbYesNo
   title = "Print Proc"
   response = MsgBox(Msg, style, title)
   
   If response = vbNo Then
   Debug.Print "CANCEL"
      Printer.KillDoc ' Terminate print job abruptly.
      Printer.EndDoc
      Close #1
      Close #7
   Else
      If response = vbYes Then
         
         Printer.Orientation = 2
         Printer.FontSize = 7
         Printer.FontName = Screen.Fonts(12)
         Printer.FontSize = 7
         
         MyDate = "Date:" & Now
         MyCountyRoute = "                                                            County:" & County & "                          Route:" & Route
         Printer.FontBold = True
         Printer.Print ""
         Printer.Print "                                                                        SOUTH CAROLINA DEPARTMENT OF TRANSPORTATION"
         Printer.Print ""
         Printer.Print "                                                                                 EARTHWORK BALANCING SYSTEM"
         Printer.Print ""
         Printer.Print "                                                            VERSION:WIN/NT 2.0                         "; MyDate
         Printer.Print ""
         Printer.Print MyCountyRoute
         Printer.FontBold = False
         Printer.Print "===================================================================================================================================================================="
         Call PageHeader(Shrink)
         IRow = 14
         Input #7, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12
         Do While Not EOF(7)
            IRow = IRow + 1
            If IRow > 60 Then
               IRow = 4
               Call PageFooter(PCVol, PFVol, GCVol, GFVol)
               Printer.NewPage
               Call PageHeader(Shrink)
            End If
            F1 = Format(F1, "######0.00")
            F2 = Format(F2, "#######0.00")
            F3 = Format(F3, "#######0.00")
            F4 = Format(F4, "#######0.00")
            If Mid(F5, 1, 1) = "(" Then
               F5 = F5
            Else
               PCVol = PCVol + Val(F5)
               GCVol = GCVol + Val(F5)
               F5 = Format(F5, "#######0.00")
            End If
            F6 = Format(F6, "#######0.00")
            F7 = Format(F7, "#######0.00")
            F8 = Format(F8, "#######0.00")
            F9 = Format(F9, "#######0.00")
            If Mid(F10, 1, 1) = "(" Then
               F10 = F10
            Else
               PFVol = PFVol + Val(F10)
               GFVol = GFVol + Val(F10)
               F10 = Format(F10, "#######0.00")
            End If
            F11 = Format(F11, "#######0.00")
            F12 = Format(F12, "#######0.00")

mystring = "9999999999"
RSet mystring = F1
F1 = mystring
mystring = "9999999999"
RSet mystring = F2
F2 = mystring
mystring = "9999999999"
RSet mystring = F3
F3 = mystring
mystring = "9999999999"
RSet mystring = F4
F4 = mystring
mystring = "9999999999"
RSet mystring = F5
F5 = mystring
mystring = "9999999999"
RSet mystring = F6
F6 = mystring
mystring = "9999999999"
RSet mystring = F7
F7 = mystring
mystring = "9999999999"
RSet mystring = F8
F8 = mystring
mystring = "9999999999"
RSet mystring = F9
F9 = mystring
mystring = "9999999999"
RSet mystring = F10
F10 = mystring
mystring = "9999999999"
RSet mystring = F11
F11 = mystring
mystring = "9999999999"
RSet mystring = F12
F12 = mystring
            Printer.Print F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12
            Input #7, F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12
         Loop
         Printer.Print F1, F2, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12
         Call PageFooter(PCVol, PFVol, GCVol, GFVol)
         Printer.EndDoc
         Msg = "Finished Print"
         style = vbOKOnly
         title = "Print Proc"
         repsponse = MsgBox(Msg, style, title)
         Close #7
      End If
   End If
End Sub

Private Sub PageHeader(Shrink)
   Printer.Print "", "", "", "", "", "", "", "", "", "", "Page:", Printer.Page
   Printer.Print "", "", "Double", "", "", "Balance", "", "Double", "", "", "Balance", ""
   Printer.Print "Station", "Cut", "Area", "Distance", "Volume", "Cut", "Fill", "Area", "Distance", "Volume", "Fill", "F+" & Shrink
   Printer.Print "--------------------------------------------------------------------------------------------------------------------------------------------------------------------"
End Sub

Private Sub PageFooter(PCVol, PFVol, GCVol, GFVol)
Dim strPCVol As String
Dim strPFVol As String
Dim strGCVol As String
Dim strGFVol As String

   Printer.Print "===================================================================================================================================================================="
   
   
   strPCVol = Str(PCVol)
   strPFVol = Str(PFVol)
   strGCVol = Str(GCVol)
   strGFVol = Str(GFVol)
   
   strPCVol = Format(strPCVol, "######0.00")
   strPFVol = Format(strPFVol, "######0.00")
   strGCVol = Format(strGCVol, "######0.00")
   strGFVol = Format(strGFVol, "######0.00")
   
   mystring = "9999999999"
   RSet mystring = strPCVol
   strPCVol = mystring
   mystring = "9999999999"
   RSet mystring = strPFVol
   strPFVol = mystring
   mystring = "9999999999"
   RSet mystring = strGCVol
   strGCVol = mystring
   mystring = "9999999999"
   RSet mystring = strGFVol
   strGFVol = mystring
   
   Printer.Print "Page Total", "", "", "", strPCVol, "", "", "", "", strPFVol, "", ""
   Printer.Print "Grand Total", "", "", "", strGCVol, "", "", "", "", strGFVol, "", ""
   
   PCVol = 0
   PFVol = 0
End Sub
