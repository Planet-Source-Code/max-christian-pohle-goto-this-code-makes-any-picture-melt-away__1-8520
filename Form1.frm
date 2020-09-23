VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   2595
   ClientTop       =   2745
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   4845
   Begin VB.CommandButton Command2 
      Caption         =   "(re-)start"
      Default         =   -1  'True
      Height          =   465
      Left            =   630
      TabIndex        =   2
      Top             =   3870
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Stop"
      Height          =   465
      Left            =   1845
      TabIndex        =   1
      Top             =   3870
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Left            =   135
      Top             =   90
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   540
      ScaleHeight     =   3195
      ScaleWidth      =   3960
      TabIndex        =   0
      Top             =   450
      Width           =   4020
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "This should be a Picture!"
         ForeColor       =   &H0000FF00&
         Height          =   735
         Left            =   270
         TabIndex        =   3
         Top             =   0
         Width           =   3345
      End
      Begin VB.Image Image1 
         Height          =   2835
         Left            =   135
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   315
         Width           =   3675
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code-Example by: McP
'...and do not forget to visit www.coderonline.de !!


Const LineWidth = 2     'change this and watch the effect!


Dim Cancel As Boolean

Private Sub Command1_Click()
Cancel = True
End Sub

Private Sub Command2_Click()
Picture1.Refresh

Command1_Click
Cancel = False
Melt
End Sub

Private Sub Form_Load()
On Error Resume Next
Picture1.BackColor = RGB(0, 0, 0)
End Sub

Sub Melt()
On Error Resume Next
Picture1.AutoRedraw = False    'is very important if you wanna use an image-control inside (this is useful, because you can stretch it)
Picture1.DrawWidth = LineWidth 'Do not change this right here ... but look for CONST LINEWIDTH=1

Do While Not blackwhole = True 'This will exit procedure, when the picture is black
    
        blackwhole = True                                                      '...We want to belive the screen is black
    
    For i% = firstcoloredpoint To Picture1.Height
            GETX = CInt(Rnd * Picture1.Width) 'Get a Random POINT X
            gety = CInt(Rnd * i%)             'Get a Random POINT Y (but it should melt!!)
            
    If Cancel = True Then Exit For  'IF USER CHECKS THE STOP-BUTTON...
            
            Picture1.ForeColor = Picture1.Point(GETX, gety)
            
                If Not Picture1.ForeColor = QBColor(0) Then blackwhole = False: firstcoloredpoint = gety '...If any other color then black is on the screen go on!
            
            Picture1.DrawWidth = LineWidth + 0:   Picture1.Line (GETX, gety)-(GETX, gety + 100) ' a real raindrop seems to be a line
            Picture1.DrawWidth = LineWidth + 2:   Picture1.PSet (GETX, gety + 100)              ' with a drop (point) ....
DoEvents
    Next i%

If Cancel = True Then Exit Do   'IF USER CHECKS THE STOP-BUTTON...
Loop
End Sub
