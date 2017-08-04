VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Factorial"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   2385
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Factorial"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function Factorial(ByVal N As Double) As Double
    If N <= 1 Then
        Factorial = 1
    Else
        Factorial = Factorial(N - 1) * N
    End If
End Function

Private Sub Command1_Click()
    Dim N As Double
    N = Val(Me.Text1.Text)
    MsgBox Str(N) + "!=" + Str(Factorial(N))
End Sub
