VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   2640
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   960
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    MsgBox Me.Winsock1.LocalIP
End Sub
