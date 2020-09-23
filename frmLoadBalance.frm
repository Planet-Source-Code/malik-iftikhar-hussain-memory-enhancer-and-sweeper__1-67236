VERSION 5.00
Begin VB.Form frmLoadBalance 
   Caption         =   "Load Balancing"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   2580
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmLoadBalance.frx":0000
      Top             =   900
      Width           =   480
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   2160
      X2              =   1920
      Y1              =   1320
      Y2              =   1620
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   2160
      X2              =   2340
      Y1              =   1320
      Y2              =   1620
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   2340
      X2              =   1920
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   300
      X2              =   4320
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "frmLoadBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Line1.X1 = 300
    
End Sub

Private Sub Image1_Click()
    MsgBox Image1.Left
End Sub

Private Sub Timer1_Timer()
    Image1.Left = Image1.Left + 100
    If Image1.Left = 4040 Then
        Image1.Left = 1940
    End If
End Sub
