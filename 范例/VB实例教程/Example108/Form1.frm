VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Icon"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   2745
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1515
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" _
                Alias "ExtractAssociatedIconA" _
                (ByVal hInst As Long, _
                ByVal lpIconPath As String, _
                lpiIcon As Long) _
                As Long

Private Declare Function DrawIcon Lib "user32" _
                (ByVal hdc As Long, _
                ByVal x As Long, _
                ByVal y As Long, _
                ByVal hIcon As Long) _
                As Long

Private Sub Form_Paint()
    Dim hIcon As Long
    Dim IconIndex As Long
    IconIndex = 0
    hIcon = ExtractAssociatedIcon(App.hInstance, "C:\AutoExec.bat", IconIndex)
    Set Me.Picture1.Picture = LoadPicture()
    Me.Picture1.AutoRedraw = True
    DrawIcon Me.Picture1.hdc, 10, 10, hIcon
    Picture1.AutoRedraw = False
    Picture1.Refresh
End Sub
