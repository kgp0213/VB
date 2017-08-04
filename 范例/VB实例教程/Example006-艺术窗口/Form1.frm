VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Shape Shape2 
      Height          =   2535
      Left            =   3960
      Shape           =   2  'Oval
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   255
   End
   Begin VB.Shape Shape7 
      Height          =   615
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   255
   End
   Begin VB.Shape Shape8 
      Height          =   615
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   735
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Left            =   1800
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      Height          =   2535
      Left            =   120
      Shape           =   2  'Oval
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   3375
      Left            =   480
      Shape           =   2  'Oval
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CombineRgn Lib "gdi32" _
                (ByVal hDestRgn As Long, _
                ByVal hSrcRgn1 As Long, _
                ByVal hSrcRgn2 As Long, _
                ByVal nCombineMode As Long) _
                As Long

Private Declare Function GetWindowRgn Lib "user32" _
                (ByVal hWnd As Long, _
                ByVal hRgn As Long) _
                As Long

Private Declare Function CreateEllipticRgn Lib "gdi32" _
                (ByVal X1 As Long, _
                ByVal Y1 As Long, _
                ByVal X2 As Long, _
                ByVal Y2 As Long) _
                As Long

Private Declare Function SetWindowRgn Lib "user32" _
                (ByVal hWnd As Long, _
                ByVal hRgn As Long, _
                ByVal bRedraw As Boolean) _
                As Long

Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_DIFF = 4
Private Const RGN_OR = 2
Private Const RGN_XOR = 3

Private Sub Form_Load()
    Me.ScaleMode = 3
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim hRgn1, hRgn2, hRgn3 As Long
    hRgn1 = CreateEllipticRgn(Me.Shape1.Left, Me.Shape1.Top, _
                Me.Shape1.Left + Me.Shape1.Width, Me.Shape1.Top + Me.Shape1.Height)

    hRgn2 = CreateEllipticRgn(Me.Shape2.Left, Me.Shape2.Top, _
                Me.Shape2.Left + Me.Shape2.Width, Me.Shape2.Top + Me.Shape2.Height)
    hRgn3 = CreateEllipticRgn(Me.Shape3.Left, Me.Shape3.Top, _
                Me.Shape3.Left + Me.Shape3.Width, Me.Shape3.Top + Me.Shape3.Height)
    
    Call CombineRgn(hRgn1, hRgn1, hRgn2, RGN_OR)
    Call CombineRgn(hRgn1, hRgn1, hRgn3, RGN_OR)
    Call SetWindowRgn(Me.hWnd, hRgn1, True)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetWindowRgn(Me.hWnd, 0, True)
End Sub
