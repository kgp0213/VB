VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "rtf"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   2940
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command_Write 
      Caption         =   "Write"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3201
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Write_Click()
    Dim dbTest As Database
    Dim rdTest As Recordset
    If Right(App.Path, 1) = "\" Then
        Set dbTest = OpenDatabase(App.Path + "db1.mdb")
    Else
        Set dbTest = OpenDatabase(App.Path + "\db1.mdb")
    End If
    Set rdTest = dbTest.OpenRecordset("rtf表")
    rdTest.AddNew
    If Right(App.Path, 1) = "\" Then
        RichTextBox1.LoadFile App.Path + "test.rtf"
    Else
        RichTextBox1.LoadFile App.Path + "\test.rtf"
    End If
    rdTest("文档内容").AppendChunk RichTextBox1.TextRTF
    rdTest.Update
End Sub
