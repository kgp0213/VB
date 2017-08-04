VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "ÍÏ·Å"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   4755
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub RichTextBox1_OLEDragDrop( _
                    Data As RichTextLib.DataObject, _
                    Effect As Long, Button As Integer, _
                    Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFText) Then
        Me.RichTextBox1.SelText = Data.GetData(vbCFText)
    ElseIf Data.GetFormat(vbCFFiles) Then
        Me.RichTextBox1.LoadFile Data.Files.Item(1)
    Else
        Effect = vbDropEffectNone
    End If
End Sub
