VERSION 5.00
Begin VB.Form frmAnimateTest 
   Caption         =   "Animate Test ©2000 Jeremy Stein"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   LinkTopic       =   "Form3"
   ScaleHeight     =   1770
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAni 
      Height          =   495
      Index           =   3
      Left            =   5280
      Picture         =   "FRMANI~1.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picAni 
      Height          =   495
      Index           =   2
      Left            =   4560
      Picture         =   "FRMANI~1.frx":69BC
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picAni 
      Height          =   495
      Index           =   1
      Left            =   3840
      Picture         =   "FRMANI~1.frx":D378
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picAni 
      Height          =   495
      Index           =   0
      Left            =   3120
      Picture         =   "FRMANI~1.frx":13D34
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "This tests animated backgrounds for windowed controls"
      Top             =   120
      Width           =   6135
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   285
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   6075
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6120
      Top             =   1200
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Here just to show another control with animation"
      Top             =   480
      Width           =   6375
   End
End
Attribute VB_Name = "frmAnimateTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Jeremy Stein
'jeremy11@mediaone.net
'Use anywhere anytime code
'Make sure your source picture box has autodraw set to true
'If you do anything cool with it let me know


Const MERGECOPY = &HC000CA
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long


Private Sub AnimateBackground()

  Dim cmbHdc As Long, txtHdc As Long
  Dim retval As Long
  Static picCounter As Integer
  
  cmbHdc = GetDC(Combo1.hWnd)
  txtHdc = GetDC(Text1.hWnd)
  picCounter = picCounter + 1
  Picture1.Cls
  If picCounter > 4 Then picCounter = 1
  
  'Uncomment following to load pictures dynamically
  'Picture1.Picture = LoadPicture(App.Path & "\images\" & picCounter & ".bmp")
  Picture1.Picture = picAni(picCounter - 1).Picture
  
  'Change the control.text here
  Picture1.Print Text1.Text

  'x,y,hieght and width values need to be adjusted according to the control
  retval = BitBlt(cmbHdc, 2, 3, Combo1.Width, Combo1.Height, Picture1.hDC, 0, 0, MERGECOPY)
  retval = BitBlt(txtHdc, 1, 1, Text1.Width, Text1.Height, Picture1.hDC, 0, 0, MERGECOPY)
  
End Sub

Private Sub Command1_Click()
  
  Timer1.Enabled = True
  
End Sub

Private Sub Command2_Click()

  Timer1.Enabled = False
  Text1.Refresh
  Combo1.Refresh
  
End Sub

Private Sub Timer1_Timer()

  AnimateBackground
  
End Sub
