VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DataEnvironment1 
   ClientHeight    =   9255
   ClientLeft      =   1080
   ClientTop       =   1500
   ClientWidth     =   10350
   _ExtentX        =   18256
   _ExtentY        =   16325
   FolderFlags     =   3
   TypeLibGuid     =   "{A5DC9AF5-9235-11D1-B067-00DD01144174}"
   TypeInfoGuid    =   "{A5DC9AF6-9235-11D1-B067-00DD01144174}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "CnnNWind"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PVB98\NWIND.MDB;Persist Security Info=False"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   1
   BeginProperty Recordset1 
      CommandName     =   "CmdNWind"
      CommDispId      =   1002
      RsDispId        =   1004
      CommandText     =   $"DataEnvironment1.dsx":0000
      ActiveConnectionName=   "CnnNWind"
      CommandType     =   1
      Prepared        =   -1  'True
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "EmployeeID"
         Caption         =   "EmployeeID"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "LastName"
         Caption         =   "LastName"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   202
         Name            =   "FirstName"
         Caption         =   "FirstName"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "city"
         Caption         =   "city"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "pracity"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
         ParamValue      =   "Seattle"
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DataEnvironment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataEnvironment_Initialize()
    pracity = InputBox("请输入城市", "城市", "Seattle")
End Sub

Private Sub DataEnvironment_Terminate()
Set Adodc1.Recordset = DataEnvironment1.rsCmdNWind

End Sub
