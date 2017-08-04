VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T Shaped Form"
   ClientHeight    =   2790
   ClientLeft      =   1290
   ClientTop       =   1935
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdT 
      Caption         =   "Change to &T"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
   X As Long
   Y As Long
End Type
Dim XY() As POINTAPI

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Sub cmdT_Click()
Dim hRgn As Long
Dim lRes As Long
ReDim XY(7) As POINTAPI 'T shape has 8 points
'
' Points must be in order like connecting the dots. Start at the origin
' and following from point to point. You don't need to specify the origin
' point as the last entry.
'
With Me
    XY(0).X = 0
    XY(0).Y = 0
    XY(1).X = .ScaleWidth
    XY(1).Y = 0
    XY(2).X = .ScaleWidth
    XY(2).Y = .ScaleHeight / 2
    XY(3).X = .ScaleWidth - (.ScaleWidth / 3)
    XY(3).Y = .ScaleHeight / 2
    XY(4).X = .ScaleWidth - (.ScaleWidth / 3)
    XY(4).Y = .ScaleHeight
    XY(5).X = .ScaleWidth / 3
    XY(5).Y = .ScaleHeight
    XY(6).X = .ScaleWidth / 3
    XY(6).Y = .ScaleHeight / 2
    XY(7).X = 0
    XY(7).Y = .ScaleHeight / 2
End With

hRgn = CreatePolygonRgn(XY(0), 8, 2)
lRes = SetWindowRgn(Me.hWnd, hRgn, True)
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.ScaleMode = vbPixels
End Sub
