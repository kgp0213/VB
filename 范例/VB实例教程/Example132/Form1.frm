VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SHARD_PATH = &H2&
  
Private Declare Function SHAddToRecentDocs Lib "shell32.dll" _
                (ByVal dwFlags As Long, _
                ByVal dwData As String) _
                As Long
  
Public Sub AddRecent(strFile As String)
    Dim lRetVal As Long
    If strFile = "" Then
        lRetVal = SHAddToRecentDocs(SHARD_PATH, vbNullString)
    Else
        lRetVal = SHAddToRecentDocs(SHARD_PATH, strFile)
    End If
End Sub

Private Sub Form_Load()
    If Right(App.Path, 1) = "\" Then
        AddRecent (App.Path + App.EXEName + ".exe")
    Else
        AddRecent (App.Path + "\" + App.EXEName + ".exe")
    End If
End Sub
