VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Record"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   3225
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Person
    ID As Integer
    Name As String
End Type

Sub WriteData()
    Dim MyRecord As Person
    Dim recordNumber As Integer
    
    Dim FileNum As Integer
    FileNum = FreeFile()
    Open App.Path + "\TestFile.dat" For Random As #FileNum
    For recordNumber = 1 To 5 Step 1
        MyRecord.ID = recordNumber
        MyRecord.Name = "My Name" & recordNumber
        
        Put #FileNum, recordNumber, MyRecord
    Next recordNumber
    Close #FileNum
End Sub

Sub ReadData()
    Dim MyRecord As Person
    Dim recordNumber As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile()
    Open App.Path + "\TestFile.dat" For Random As #FileNum

    Me.Text1.Text = ""
    Dim i As Integer
    i = 1
    Do While Not EOF(1)
        Seek #FileNum, i
        Get #FileNum, i, MyRecord
        
        Me.Text1.Text = Me.Text1 + Str(MyRecord.ID) + Chr(13) + Chr(10)
        Me.Text1.Text = Me.Text1 + MyRecord.Name + Chr(13) + Chr(10)
        Me.Text1.Text = Me.Text1 + "==================" + Chr(13) + Chr(10)
        i = i + 1
    Loop

    Close #FileNum
End Sub

Private Sub Form_Load()
    WriteData
    ReadData
End Sub
