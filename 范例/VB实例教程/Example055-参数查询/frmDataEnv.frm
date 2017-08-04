VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDataEnv 
   Caption         =   "参数查询"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5760
   StartUpPosition =   3  '窗口缺省
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmDataEnv.frx":0000
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "frmDataEnv.frx":0014
      TabIndex        =   3
      Top             =   600
      Width           =   5415
   End
   Begin VB.Data Data1 
      Caption         =   "记录"
      Connect         =   "Access"
      DatabaseName    =   "D:\Microsoft Visual Studio\VB98\NWIND.MDB"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Employees"
      Top             =   3480
      Width           =   4095
   End
   Begin VB.CommandButton CmdF 
      Caption         =   "查询"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1560
      TabIndex        =   1
      Text            =   "London"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "选择城市："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdF_Click()
    Dim questr As String
    questr = "Select EmployeeID,LastName,FirstName,City" + " " & _
             "From Employees" + " " & _
             "Where" + " " + "city=" + "'" + Text1.Text + "'"
    Data1.RecordSource = questr
    Data1.Refresh
End Sub

