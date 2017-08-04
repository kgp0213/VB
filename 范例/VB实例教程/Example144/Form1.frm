VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Resource"
   ClientHeight    =   1830
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   2970
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu mnu_File 
      Caption         =   "文件/File"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.mnu_File.Caption = LoadResString(101)
End Sub
