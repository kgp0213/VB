VERSION 5.00
Begin VB.Form frmPrime 
   Caption         =   "Find the Prime Numbers"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstResults 
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrime 
      Caption         =   "Find Prime"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblResults 
      Caption         =   "Results:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmPrime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrime_Click()
    lstResults.Clear
    lblResults = "Results:"
    lngStart = InputBox("Enter the beginning number:", "Start Number")
    lngEnd = InputBox("Enter the ending number:", "End Number")
    'Since this is for inbetween, add one to the starting number
    lngStart = lngStart + 1
    Do Until lngStart = lngEnd
        If PrimeStatus(lngStart) = True Then
            lstResults.AddItem Str(lngStart)
            lngCount = lngCount + 1
        Else
            'Not Prime don't care
        End If
        lngStart = lngStart + 1
    Loop
    lblResults = lblResults & " " & Str(lngCount)
    lngCount = 0
    
End Sub

