VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "图像拖放"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4155
   ScaleWidth      =   3450
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3855
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   3795
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handle
    If Data.GetFormat(vbCFBitmap) Or _
        Data.GetFormat(vbCFMetafile) Or _
        Data.GetFormat(vbCFDIB) Or _
        Data.GetFormat(vbCFEMetafile) Then
        Me.Picture1.Picture = Data.GetData(vbCFBitmap)
    ElseIf Data.GetFormat(vbCFFiles) Then
        Me.Picture1.Picture = LoadPicture(Data.Files.Item(1))
    Else
        Effect = vbDropEffectNone
    End If
    Exit Sub
Err_Handle:
    MsgBox Err.Description
    Exit Sub
End Sub


Private Sub Picture1_Resize()
'窗口随着Picture控件改变大小
Me.Width = Picture1.Width + 600
Me.Height = Picture1.Height + 600
End Sub
