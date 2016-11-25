VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Survey Systems"
   ClientHeight    =   4410
   ClientLeft      =   8430
   ClientTop       =   5025
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   6315
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   450
      Top             =   495
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Menu mnuFIle 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSas 
      Caption         =   "&Sas"
   End
   Begin VB.Menu mnuSoa 
      Caption         =   "&Soa"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   App.HelpFile = App.Path & "\Soa.hlp"
   mnuSoa.Visible = False
   mnuSas.Visible = False
End Sub

Private Sub mnuExit_Click()
   End
End Sub

Private Sub mnuNew_Click()
   
   With dlg1
      .Flags = cdlOFNHideReadOnly
      .Filter = "Survey Files (*.Alg)|*.Alg"
      .FilterIndex = 2
      .ShowOpen
      Sfile = .filename
      
      IPos = InStr(1, frmMain.dlg1.filename, ".", 0)
      
      Alg_filename = Mid(Sfile, 1, IPos) & "Alg"
      Aln_filename = Mid(Sfile, 1, IPos) & "Aln"
      Cla_filename = Mid(Sfile, 1, IPos) & "Cla"
      Clp_filename = Mid(Sfile, 1, IPos) & "Clp"
      GP3_filename = Mid(Sfile, 1, IPos) & "Gp3"
      Map_filename = Mid(Sfile, 1, IPos) & "Map"
      New_filename = Mid(Sfile, 1, IPos) & "New"
      Pnt_filename = Mid(Sfile, 1, IPos) & "Pnt"
      Rej_filename = Mid(Sfile, 1, IPos) & "Rej"
      SasRpt_filename = "Sas.txt"
      SoaRpt_filename = "Soa.txt"
      SOE_filename = Mid(Sfile, 1, IPos) & "Soe"
      SOR_filename = Mid(Sfile, 1, IPos) & "Sor"
      Tmp_filename = Mid(Sfile, 1, IPos) & "Tmp"
      SO1_filename = Mid(Sfile, 1, IPos) & "So1"
      Vld_filename = Mid(Sfile, 1, IPos) & "Vld"
      XYZ_filename = Mid(Sfile, 1, IPos) & "Xyz"
   
      Rt_filename = Mid(Sfile, 1, IPos - 1) & "Rt.tmp"
      Lt_filename = Mid(Sfile, 1, IPos - 1) & "Lt.tmp"
      DgnRt_filename = Mid(Sfile, 1, IPos - 1) & "Rt.txt"
      DgnLt_filename = Mid(Sfile, 1, IPos - 1) & "Lt.txt"
   End With
   mnuSoa.Visible = True
   mnuSas.Visible = True
End Sub

'Private Sub dirFileOpen_Change()
'   fileOpen.Path = dirFileOpen.Path
'   ChDir dirFileOpen.Path
'End Sub

'Private Sub drvFileOpen_change()
'   dirFileOpen.Path = drvFileOpen.Drive
'   ChDrive drvFileOpen.Drive
'End Sub

Private Sub mnuOpen_Click()
   
   With dlg1
      .Filter = "Survey Files (*.Alg)|*.Alg|Map Files (*.Map)|*.Map"
      .FilterIndex = 1
      .ShowOpen
      Sfile = .filename
      
      IPos = InStr(1, frmMain.dlg1.filename, ".", 0)
      
      Alg_filename = Mid(Sfile, 1, IPos) & "Alg"
      Aln_filename = Mid(Sfile, 1, IPos) & "Aln"
      Cla_filename = Mid(Sfile, 1, IPos) & "Cla"
      Clp_filename = Mid(Sfile, 1, IPos) & "Clp"
      GP3_filename = Mid(Sfile, 1, IPos) & "Gp3"
      Map_filename = Mid(Sfile, 1, IPos) & "Map"
      New_filename = Mid(Sfile, 1, IPos) & "New"
      Pnt_filename = Mid(Sfile, 1, IPos) & "Pnt"
      Rej_filename = Mid(Sfile, 1, IPos) & "Rej"
      SasRpt_filename = "Sas.txt"
      SoaRpt_filename = "Soa.txt"
      SOE_filename = Mid(Sfile, 1, IPos) & "Soe"
      SOR_filename = Mid(Sfile, 1, IPos) & "Sor"
      Tmp_filename = Mid(Sfile, 1, IPos) & "Tmp"
      SO1_filename = Mid(Sfile, 1, IPos) & "So1"
      Vld_filename = Mid(Sfile, 1, IPos) & "Vld"
      XYZ_filename = Mid(Sfile, 1, IPos) & "Xyz"
   
      Rt_filename = Mid(Sfile, 1, IPos - 1) & "Rt.tmp"
      Lt_filename = Mid(Sfile, 1, IPos - 1) & "Lt.tmp"
      DgnRt_filename = Mid(Sfile, 1, IPos - 1) & "Rt.txt"
      DgnLt_filename = Mid(Sfile, 1, IPos - 1) & "Lt.txt"
   End With
   mnuSoa.Visible = True
   mnuSas.Visible = True
End Sub

Private Sub mnuSas_Click()
   frmSasInp.optType(0) = False
   frmSasInp.optType(1) = False
   frmSasInp.optType(2) = False
   
   frmSasInp.cmdNewAlign.Visible = False
   
   frmSasInp.txtAlignment.Visible = False
   frmSasInp.txtDescription.Visible = False
   frmSasInp.txtStation.Visible = False
   frmSasInp.lblAlign.Visible = False
   frmSasInp.lblDesc.Visible = False
   frmSasInp.lblStation.Visible = False
   
   frmSasInp.lblUnits.Visible = False
   frmSasInp.comboUnits.Visible = False
   
   frmSasInp.txtPnt1.Visible = False
   frmSasInp.txtPnt2.Visible = False
   frmSasInp.txtPnt3.Visible = False
   frmSasInp.txtPnt4.Visible = False
   frmSasInp.lblPnt1.Visible = False
   frmSasInp.lblPnt2.Visible = False
   frmSasInp.lblPnt3.Visible = False
   frmSasInp.lblPnt4.Visible = False
   
   frmSasInp.frameCurve.Visible = False
   
   frmSasInp.optCurve(0).Visible = True
   frmSasInp.optCurve(1).Visible = True
   frmSasInp.optCurve(2).Visible = True
   
   frmSasInp.Show vbModal, Me
   
End Sub

Private Sub mnuSoa_Click()
   frmSoaInp.Show vbModal, Me
End Sub

