VERSION 5.00
Begin VB.Form frmSoaInp 
   Caption         =   "Soa Options Screen"
   ClientHeight    =   4425
   ClientLeft      =   8430
   ClientTop       =   5025
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   6330
   Begin VB.ListBox lstSoaSort 
      Height          =   450
      ItemData        =   "frmSoaInp.frx":0000
      Left            =   2565
      List            =   "frmSoaInp.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   3735
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Continue"
      Height          =   420
      Left            =   4320
      TabIndex        =   16
      Top             =   2160
      Width           =   1005
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   4725
      TabIndex        =   12
      Top             =   1665
      Width           =   600
   End
   Begin VB.TextBox txtHight 
      Height          =   285
      Left            =   4725
      TabIndex        =   11
      Top             =   1260
      Width           =   600
   End
   Begin VB.TextBox txtFont 
      Height          =   285
      Left            =   4725
      TabIndex        =   10
      Top             =   855
      Width           =   600
   End
   Begin VB.CheckBox chkPnt 
      Caption         =   "Do you want to process every point"
      Height          =   330
      Left            =   450
      TabIndex        =   9
      Top             =   135
      Width           =   3030
   End
   Begin VB.Frame frameTypeProcOpt 
      Caption         =   "Type Proc Options"
      Height          =   2895
      Left            =   405
      TabIndex        =   1
      Top             =   675
      Width           =   3120
      Begin VB.OptionButton optTypeProc 
         Caption         =   "Drainage"
         Height          =   300
         Index           =   6
         Left            =   100
         TabIndex        =   8
         Top             =   2295
         Width           =   2850
      End
      Begin VB.OptionButton optTypeProc 
         Caption         =   "Single Alignment w/ Station Range"
         Height          =   300
         Index           =   5
         Left            =   100
         TabIndex        =   7
         Top             =   1980
         Width           =   2850
      End
      Begin VB.OptionButton optTypeProc 
         Caption         =   "Single Alignment w/ Point Range"
         Height          =   300
         Index           =   4
         Left            =   100
         TabIndex        =   6
         Top             =   1665
         Width           =   2850
      End
      Begin VB.OptionButton optTypeProc 
         Caption         =   "Single Alignment w/ All Points"
         Height          =   300
         Index           =   3
         Left            =   100
         TabIndex        =   5
         Top             =   1350
         Width           =   2850
      End
      Begin VB.OptionButton optTypeProc 
         Caption         =   "All Alignments w/ Station Range"
         Height          =   300
         Index           =   2
         Left            =   100
         TabIndex        =   4
         Top             =   1035
         Width           =   2850
      End
      Begin VB.OptionButton optTypeProc 
         Caption         =   "All Alignments w/ Point Range"
         Height          =   300
         Index           =   1
         Left            =   100
         TabIndex        =   3
         Top             =   720
         Width           =   2850
      End
      Begin VB.OptionButton optTypeProc 
         Caption         =   "All Alignments w/ All Points"
         Height          =   300
         Index           =   0
         Left            =   100
         TabIndex        =   2
         Top             =   405
         Width           =   2850
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   465
      Left            =   1215
      TabIndex        =   0
      Top             =   3735
      Width           =   1230
   End
   Begin VB.Label lblWidth 
      Caption         =   "Width"
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   1710
      Width           =   705
   End
   Begin VB.Label lblHight 
      Caption         =   "Hight"
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   1305
      Width           =   705
   End
   Begin VB.Label lblFont 
      Caption         =   "Font"
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   900
      Width           =   705
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
End
Attribute VB_Name = "frmSoaInp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Open SoaRpt_filename For Output As #9
   Open Vld_filename For Output As #12
   Open Map_filename For Append As #3
   MySize = LOF(3)
   
   Close #3
   If MySize = 0 Then
      Call CreateMap
   End If

   Call CreateIdx
   
   mnuPrint.Enabled = False

   Open Alg_filename For Append As #1
   MySize = LOF(1)
   Close #1
   If MySize = 0 Then
      Msg = "Missing ALG file. Please run the SAS option to create file."
      Style = vbOK
      Title = "File Processing Error"
      Response = MsgBox(Msg, Style, Title)
   Else
      Open Alg_filename For Input As #1
      Input #1, RecType, Desc, Align, sta, Units
      wUnits = Units
      Close #1
   End If
   Open Aln_filename For Append As #4
   MySize = LOF(4)
   Close #4
   If MySize = 0 Then
      Msg = "Missing ALN file. Please run the SAS option to create file."
      Style = vbOK
      Title = "File Processing Error"
      Response = MsgBox(Msg, Style, Title)
   Else
      Open Aln_filename For Input As #4
      AlnCntr = 1
      Do While Not EOF(4)
         Input #4, AlignNM(AlnCntr), ElemTyp(AlnCntr), AlgDesc(AlnCntr), _
                   x1, y1, X2, Y2, X3, Y3, bsta, ESta, CEsta, A, B, C, dir
'
         If AlnCntr > 200 Then
            Msg = "Error Trying to Process to many Alignments."
            Style = vbOK
            Title = "Critical Error-Alignment Processing"
            Response = MsgBox(Msg, Style, Title)
            Unload frmSoaInp
         End If
         AlnCntr = AlnCntr + 1
      Loop
      Close #4
      AlnCntr = AlnCntr - 1
   End If
' Set print parms
   txtFont = 3
   If wUnits = "Metric" Then
      txtWidth = 2
      txtHight = 2
   Else
      txtWidth = 5
      txtHight = 5
   End If
End Sub

Private Sub mnuPrint_Click()

   Msg = "Continue Printing"
   Style = vbYesNo
   Title = "Print Proc"
   Response = MsgBox(Msg, Style, Title)
   
   Open Vld_filename For Input As #12
   
   If Response = vbNo Then
      Printer.KillDoc ' Terminate print job abruptly.
      Printer.EndDoc
   Else
      If Response = vbYes Then
         Printer.Orientation = 2       ' Portrait orientation
         Printer.FontSize = 9
         Printer.FontName = Screen.Fonts(12)
         Printer.FontSize = 9
         Page = 1
         Call Header(Page)
         
         IRow = 9
         While Not EOF(12)
            Input #12, SF1, SF2, SF3, SF4, SF5, SF6, SF7, SF8, SF9
            IRow = IRow + 1
            If IRow > 55 Then
               IRow = 9
               Printer.NewPage
               Page = Page + 1
               Call Header(Page)
            End If
            SF1 = Format(SF1, "####0")
            SF2 = Format(SF2, "###0.00")
            SF3 = Format(SF3, "###0.00")
'            SF4 = Format(SF4, "###0.00")
            SF5 = Format(SF5, "###0.00")
            SF6 = Format(SF6, "###0.00")
            SF7 = Format(SF7, "###0.00")
            SF8 = Format(SF8, "###0")
'            SF9 = Format(SF9, "###0")
            Printer.Print SF1, SF2, SF3, SF5, SF6, SF7, SF4, SF8, SF9
         Wend
         
         Printer.EndDoc
         Msg = "Finished SOA Print"
         Style = vbOKOnly
         Title = "Print Proc"
         Response = MsgBox(Msg, Style, Title)
      End If
   End If
   Close #12
End Sub
Private Sub Header(Page)
   
   MyDate = "Date:" & Now
   MyPage = "Page:" & Page
   Printer.FontBold = True
   Printer.Print ""
   Printer.Print "     "; MyDate; "                                                                                                           "; MyPage
   Printer.Print ""
   Printer.Print "                                                  SOUTH CAROLINA DEPARTMENT OF TRANSPORTATION"
   Printer.Print ""
   Printer.Print "                                                      ALIGNMENT, STATION AND OFFSET REPORT"
   Printer.Print ""
   Printer.Print "                                                                VERSION:WIN/NT 2.0"
   Printer.Print ""
   Printer.FontBold = False
   F1a = "PNT NO."
   F2a = "X-COORD"
   F3a = "Y-COORD"
   F3a = "Z-COORD"
   F4a = "DESCRIPTION                "
   F5a = "STATION"
   F6a = "OFFSET"
   F7a = "ELEVATION"
   F8a = "ALIGN NO."
   F9a = "ALIGNMENT DESC."
   Printer.Print F1a, F2a, F3a, F4a, F5a, F6a, F7a, F8a, F9a
   Printer.Print "_____________________________________________________________________________________________________________________________________________"
End Sub

Private Sub optTypeProc_Click(Index As Integer)
   
   Open Map_filename For Input As #3
   Open Rej_filename For Output As #11
   Open Tmp_filename For Output As #22
'
   Do While Not EOF(3)
      Input #3, Pnt, iX, iY, iZ, Desc
      If optTypeProc(6) = True Then
         If Mid(Desc, 1, 3) = "CBN" Or Mid(Desc, 1, 2) = "DI" Or _
            Mid(Desc, 1, 3) = "MHD" Or Mid(Desc, 1, 3) = "MHS" Or _
            Mid(Desc, 1, 3) = "ODL" Or Mid(Desc, 1, 2) = "P " Or _
            Mid(Desc, 1, 2) = "PA" Or Mid(Desc, 1, 2) = "PB" Or _
            Mid(Desc, 1, 3) = "P1A" Or Mid(Desc, 1, 3) = "P1B" Or _
            Mid(Desc, 1, 3) = "P2A" Or Mid(Desc, 1, 3) = "P2B" Or _
            Mid(Desc, 1, 3) = "P3A" Or Mid(Desc, 1, 3) = "P3B" Or _
            Mid(Desc, 1, 3) = "P4A" Or Mid(Desc, 1, 3) = "P4B" Or _
            Mid(Desc, 1, 3) = "BRC" Or Mid(Desc, 1, 3) = "BRW" Or _
            Mid(Desc, 1, 3) = "CUL" Or Mid(Desc, 1, 2) = "HW" Or _
            Mid(Desc, 1, 2) = "WW" Or Mid(Desc, 1, 2) = "JB" Then
            Write #11, Pnt, iX, iY, iZ, Desc
         ElseIf Mid(Desc, 1, 4) = "*CBN" Or Mid(Desc, 1, 3) = "*DI" Or _
            Mid(Desc, 1, 4) = "*MHD" Or Mid(Desc, 1, 4) = "*MHS" Or _
            Mid(Desc, 1, 4) = "*ODL" Or Mid(Desc, 1, 3) = "*P " Or _
            Mid(Desc, 1, 3) = "*PA" Or Mid(Desc, 1, 3) = "*PB" Or _
            Mid(Desc, 1, 4) = "*P1A" Or Mid(Desc, 1, 4) = "*P1B" Or _
            Mid(Desc, 1, 4) = "*P2A" Or Mid(Desc, 1, 4) = "*P2B" Or _
            Mid(Desc, 1, 4) = "*P3A" Or Mid(Desc, 1, 4) = "*P3B" Or _
            Mid(Desc, 1, 4) = "*P4A" Or Mid(Desc, 1, 4) = "*P4B" Or _
            Mid(Desc, 1, 4) = "*BRC" Or Mid(Desc, 1, 4) = "*BRW" Or _
            Mid(Desc, 1, 4) = "*CUL" Or Mid(Desc, 1, 3) = "*HW" Or _
            Mid(Desc, 1, 3) = "*WW" Or Mid(Desc, 1, 3) = "*JB" Then
            Write #11, Pnt, iX, iY, iZ, Desc
         End If
      Else
         If chkPnt = False And IsNumeric(Pnt) And Mid(Desc, 1, 1) <> "*" Then
            Write #11, Pnt, iX, iY, iZ, Desc
         End If
      End If
   Loop
   
   Close #11
   Close #3
'
   Open Aln_filename For Input As #4
'
   frmSoaTypeProcInp.Show vbModal, Me
'
   Close #4
   
   Close #12

End Sub

Private Sub cmdQuit_Click()
   Close #3
   Close #9
   Close #12
   Close #22
   Unload frmSoaInp
End Sub
Private Sub cmdPrint_Click()
   wFont = txtFont
   wWidth = txtWidth
   wHight = txtHight
   Call Output
End Sub

