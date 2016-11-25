VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Browser"
   ClientHeight    =   4725
   ClientLeft      =   1395
   ClientTop       =   1290
   ClientWidth     =   6765
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4725
   ScaleWidth      =   6765
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   2400
      ScaleHeight     =   1275
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuFileExit_Click()
    End
End Sub


Private Sub mnuFileOpen_Click()
    frmFilePicker.Show 1
    Picture1.Picture = LoadPicture(frmFilePicker.txtFileName.Text)
End Sub


