VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4230
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command_Temp 
      Caption         =   "Temp"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Temp_Click()
    MsgBox Environ("temp")
End Sub

Private Sub Form_Load()
    Me.ListView1.View = lvwReport
    Me.ListView1.ColumnHeaders.Add , , "环境变量"
    Me.ListView1.ColumnHeaders.Add , , "内容"
    
    Dim i As Integer
    Dim Env As String
    i = 1
    Env = Environ(i)
    Do Until Env = ""
        Env = Environ(i)
        If Env = "" Then Exit Do
        Dim str_Item As String
        str_Item = Left(Env, InStr(1, Env, "=") - 1)
        Me.ListView1.ListItems.Add i, , str_Item
        Me.ListView1.ListItems(i).ListSubItems.Add , , _
                Right(Env, Len(Env) - InStr(1, Env, "="))
        i = i + 1
    Loop
End Sub
