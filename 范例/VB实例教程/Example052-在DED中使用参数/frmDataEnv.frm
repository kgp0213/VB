VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDataEnv 
   Caption         =   "参数"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5685
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   360
      Top             =   2280
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.TextBox txtcity 
      DataField       =   "city"
      DataMember      =   "CmdNWind"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1860
      TabIndex        =   7
      Top             =   1515
      Width           =   2475
   End
   Begin VB.TextBox txtFirstName 
      DataField       =   "FirstName"
      DataMember      =   "CmdNWind"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1860
      TabIndex        =   5
      Top             =   1140
      Width           =   1650
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "LastName"
      DataMember      =   "CmdNWind"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   1860
      TabIndex        =   3
      Top             =   750
      Width           =   3300
   End
   Begin VB.TextBox txtEmployeeID 
      DataField       =   "EmployeeID"
      DataMember      =   "CmdNWind"
      DataSource      =   "DataEnvironment1"
      Height          =   405
      Left            =   1860
      TabIndex        =   1
      Top             =   255
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "city:"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "FirstName:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "LastName:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   795
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EmployeeID:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1080
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Adodc1.Caption = "移动记录指针"
    Set Adodc1.Recordset = DataEnvironment1.rsCmdNWind
End Sub
