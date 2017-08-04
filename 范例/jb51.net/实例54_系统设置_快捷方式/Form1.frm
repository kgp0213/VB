VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "加入"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "清除"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "选择加入，则选择文件，将其快捷方式加入到Windows开始菜单－>文档菜单中去"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub SHAddToRecentDocs Lib "shell32.dll" _
        (ByVal uFlags As Integer, ByVal pv As String)
    
Const SHARD_PATH = 2
Private Sub Command1_Click()
    Dim l As Long
    
    SHAddToRecentDocs SHARD_PATH, vbNullString
End Sub

Private Sub Command2_Click()
    Dim l As Long
    
    CommonDialog1.ShowOpen
    SHAddToRecentDocs SHARD_PATH, Trim(CommonDialog1.FileName)
End Sub

Private Sub Label1_Click()

End Sub
