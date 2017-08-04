VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ani"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command_Start 
      Caption         =   "Start"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GCL_HCURSOR = (-12)

Private Declare Function LoadCursorFromFile Lib "user32" _
                Alias "LoadCursorFromFileA" _
                (ByVal lpFileName As String) _
                As Long
                
Private Declare Function SetClassLong Lib "user32" _
                Alias "SetClassLongA" _
                (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) _
                As Long
                
Private Declare Function GetClassLong Lib "user32" _
                Alias "GetClassLongA" _
                (ByVal hwnd As Long, _
                ByVal nIndex As Long) _
                As Long

Private Sub Command_Start_Click()
    Dim mhBaseCursor As Long, mhAniCursor As Long
    Dim lResult As Long
    If Right(App.Path, 1) = "\" Then
        mhAniCursor = LoadCursorFromFile(App.Path + "horse.ani")
    Else
        mhAniCursor = LoadCursorFromFile(App.Path + "\horse.ani")
    End If
    lResult = SetClassLong(Me.hwnd, GCL_HCURSOR, mhAniCursor)
End Sub
