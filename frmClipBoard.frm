VERSION 5.00
Begin VB.Form frmClipBoard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clip Board Contents"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3915
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   3375
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   360
         Width           =   4395
      End
   End
End
Attribute VB_Name = "frmClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

