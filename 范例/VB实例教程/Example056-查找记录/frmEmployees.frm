VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEmployees 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employees"
   ClientHeight    =   4035
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5775
   Begin VB.CommandButton CmdFind 
      Caption         =   "查找"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "City"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Country"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1995
      Width           =   2775
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LastName"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   1350
      Width           =   2775
   End
   Begin VB.TextBox txtFields 
      DataField       =   "FirstName"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   705
      Width           =   2775
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EmployeeID"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   3705
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=D:\Microsoft Visual Studio\VB98\NWIND.MDB;"
      OLEDBString     =   "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=D:\Microsoft Visual Studio\VB98\NWIND.MDB;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select EmployeeID,FirstName,LastName,Country,City from Employees Order by EmployeeID"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      Caption         =   "City:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Country:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1995
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LastName:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1350
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "FirstName:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   705
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "EmployeeID:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cursorPos As String

Private Sub CmdFind_Click()
'实现查找功能
    On Error Resume Next
    If Len(cursorPos) = 0 Then
        MsgBox "请选择要查找的字段！"
        Exit Sub
    End If
    Dim FindStr As String
    FindStr = InputBox("请输入要查找的内容", cursorPos)
    FindStr = cursorPos + "='" + FindStr + "'"
    datPrimaryRS.Recordset.Find (FindStr)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_MoveComplete( _
            ByVal adReason As ADODB.EventReasonEnum, _
            ByVal pError As ADODB.Error, _
            adStatus As ADODB.EventStatusEnum, _
            ByVal pRecordset As ADODB.Recordset)
'显示当前记录位置
  datPrimaryRS.Caption = "Record: " & _
                        CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
'设置cursorPos变量
    cursorPos = ""
    Select Case Index
        Case 0
            cursorPos = "EmployeeId"
        Case 1
            cursorPos = "FirstName"
        Case 2
            cursorPos = "LastName"
        Case 3
            cursorPos = "Country"
        Case 4
            cursorPos = "City"
    End Select
  End Sub
