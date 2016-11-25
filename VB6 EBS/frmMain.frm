VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "EBS"
   ClientHeight    =   2835
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   3810
      TabIndex        =   0
      Top             =   2565
      Width           =   3870
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   960
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBar6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuProcess 
      Caption         =   "&Process"
      Begin VB.Menu mnuEnter 
         Caption         =   "&Enter"
      End
      Begin VB.Menu mnuRun 
         Caption         =   "&Run"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About EBS..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuEnter_Click()
    frmEnter.Show vbModal, Me
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewStatusBar_Click()
    If mnuViewStatusBar.Checked Then
        sbStatusBar.Visible = False
        mnuViewStatusBar.Checked = False
    Else
        sbStatusBar.Visible = True
        mnuViewStatusBar.Checked = True
    End If
End Sub

Private Sub mnuViewToolbar_Click()
    If mnuViewToolbar.Checked Then
        tbToolBar.Visible = False
        mnuViewToolbar.Checked = False
    Else
        tbToolBar.Visible = True
        mnuViewToolbar.Checked = True
    End If
End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuHelpSearch_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    'To Do
    MsgBox "Refresh Code goes here!"
End Sub

Private Sub mnuEditCopy_Click()
    'To Do
    MsgBox "Copy Code goes here!"
End Sub

Private Sub mnuEditCut_Click()
    'To Do
    MsgBox "Cut Code goes here!"
End Sub

Private Sub mnuEditPaste_Click()
    'To Do
    MsgBox "Paste Code goes here!"
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'To Do
    MsgBox "Paste Special Code goes here!"
End Sub

Private Sub mnuEditUndo_Click()
    'To Do
    MsgBox "Undo Code goes here!"
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String
    With dlgCommonDialog
        'To Do
        'set the flags and attributes of the
        'common dialog control
        .Filter = "EBS Files (*.cnf;*.cnf)"
        .ShowOpen
        If Len(.filename) = 0 Then
            Exit Sub
        End If
        sFile = .filename
    End With
    'To Do
    'process the opened file
End Sub

Private Sub mnuFileClose_Click()
    'To Do
    MsgBox "Close Code goes here!"
End Sub

Private Sub mnuFileSave_Click()
    'To Do
    MsgBox "Save Code goes here!"
End Sub

Private Sub mnuFileSaveAs_Click()
    'To Do
    'Setup the common dialog control
    'prior to calling ShowSave
    dlgCommonDialog.ShowSave
End Sub

Private Sub mnuFileSaveAll_Click()
    'To Do
    MsgBox "Save All Code goes here!"
End Sub

Private Sub mnuFileProperties_Click()
    'To Do
    MsgBox "Properties Code goes here!"
End Sub

Private Sub mnuFilePageSetup_Click()
    dlgCommonDialog.ShowPrinter
End Sub

Private Sub mnuFilePrintPreview_Click()
    'To Do
    MsgBox "Print Preview Code goes here!"
End Sub

Private Sub mnuFilePrint_Click()
    'To Do
    MsgBox "Print Code goes here!"
End Sub

Private Sub mnuFileSend_Click()
    'To Do
    MsgBox "Send Code goes here!"
End Sub

Private Sub mnuFileMRU_Click(Index As Integer)
    'To Do
    MsgBox "MRU Code goes here!"
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    'To Do
    MsgBox "New File Code goes here!"
End Sub



