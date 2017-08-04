VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "timer"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2340
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   2340
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   240
   End
   Begin VB.Label Label2 
      Caption         =   "00:00:00"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "现在时间："
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Label2.Caption = Time
End Sub
