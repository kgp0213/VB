VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Í¸Ã÷´°Ìå£¬ºÇºÇ"
      BeginProperty Font 
         Name            =   "Ó×Ô²"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim Bmp
Me.AutoRedraw = True
Bmp = CreateCompatibleBitmap(Me.hdc, 0, 0)
SelectObject Me.hdc, Bmp
Me.Refresh
End Sub

