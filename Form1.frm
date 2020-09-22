VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "ScreenCapture"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7650
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4005
      Left            =   75
      TabIndex        =   1
      Top             =   135
      Width           =   1665
      Begin VB.CommandButton Command1 
         Caption         =   "Capture Screen "
         Height          =   495
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Width           =   1395
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Print"
         Height          =   495
         Left            =   135
         TabIndex        =   6
         Top             =   3375
         Width           =   1395
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Clear"
         Height          =   495
         Left            =   135
         TabIndex        =   5
         Top             =   2745
         Width           =   1395
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Save to File"
         Height          =   495
         Left            =   135
         TabIndex        =   4
         Top             =   2115
         Width           =   1395
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Form Picture"
         Height          =   495
         Left            =   135
         TabIndex        =   3
         Top             =   1485
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Capture Form"
         Height          =   495
         Left            =   135
         TabIndex        =   2
         Top             =   870
         Width           =   1395
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1125
      Top             =   4710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".Bmp"
      DialogTitle     =   "Save Bitmap As"
      Filter          =   "*.bmp"
   End
   Begin VB.PictureBox Picture1 
      Height          =   1680
      Left            =   2430
      ScaleHeight     =   1620
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   270
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form1
' DateTime  : 10/21/2002 02:16
' Author    : Avaneesh Dvivedi
' Purpose   : Main form to capture
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : Command1_Click
' DateTime  : 10/21/2002 02:16
' Author    : Avaneesh Dvivedi
' Purpose   : It captures screen
'---------------------------------------------------------------------------------------
'
Private Sub Command1_Click()
    
    Set Picture1.Picture = CaptureScreen()
End Sub
      
      
Private Sub Command2_Click()
   Set Picture1.Picture = CaptureForm(Me)
End Sub
      
    
' Print the current contents of the picture box
Private Sub Command5_Click()
   PrintPicture Printer, Picture1.Picture
   Printer.EndDoc
End Sub
      
' Clear out the picture box
Private Sub Command6_Click()
   Set Picture1.Picture = Nothing
End Sub
      
Private Sub Command7_Click()
    On Error GoTo ErrorRoutineErr

    CommonDialog1.ShowSave
    SavePicture Picture1.Picture, CommonDialog1.FileName
    
ErrorRoutineResume:
    Exit Sub
ErrorRoutineErr:
    'User pressed cancel
    If Err = 32755 Then
        Resume ErrorRoutineResume
    End If
    MsgBox "Project1.Form1.Command7_Click" & Err & Error
    Resume Next
End Sub

Private Sub Command8_Click()
    Form2.Show
    Form2.Picture = Picture1.Picture
End Sub


      
Private Sub Form_Load()
   
   Picture1.AutoSize = True
   
End Sub
     


Private Sub mnuExit_Click()
    End
End Sub



