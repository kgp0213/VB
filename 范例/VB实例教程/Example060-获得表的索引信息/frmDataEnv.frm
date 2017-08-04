VERSION 5.00
Begin VB.Form frmDataEnv 
   Caption         =   "索引信息"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   4590
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmDataEnv.frx":0000
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim cn As ADODB.Connection
    Dim rsSchema As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim rCriteria As Variant

    Set cn = New ADODB.Connection
    
    With cn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .CursorLocation = adUseClient
        .ConnectionString = "Data Source=D:\Microsoft Visual Studio\VB98\NWIND.MDB"
        .Open
    End With

    rCriteria = Array(Empty, Empty, Empty, Empty, "employees")
    
    Set rsSchema = cn.OpenSchema(adSchemaIndexes, rCriteria)
    
    Me.Text1.Text = ""
    Me.Text1.Text = Me.Text1.Text + "Index Count: " & Str(rsSchema.RecordCount) + Chr(13) + Chr(10)
    'rsSchema.RecordCount返回索引数目
    While Not rsSchema.EOF
    '使用While语句显示每个索引的信息
       Me.Text1.Text = Me.Text1.Text + "==============================" + Chr(13) + Chr(10)
       '显示==================并换行
       For Each fld In rsSchema.Fields
       '显示当前索引中各属性名称和相应的属性值
          Me.Text1.Text = Me.Text1.Text + fld.Name + ":"
          'fld.Name为索引的属性名称
          '以下代码根据属性值决定如何显示
          If IsNull(fld.Value) Then
              Me.Text1.Text = Me.Text1.Text + "Null" + Chr(13) + Chr(10)
              '如果为Null值则在Text1中显示Null
          Else
              If VarType(fld.Value) = vbBoolean Then
                  If fld.Value = True Then
                      Me.Text1.Text = Me.Text1.Text + "True" + Chr(13) + Chr(10)
                      '如果为True则显示True
                  Else
                      Me.Text1.Text = Me.Text1.Text + "False" + Chr(13) + Chr(10)
                      '如果为False则显示False
                  End If
              ElseIf VarType(fld.Value) = vbString Then
                  Me.Text1.Text = Me.Text1.Text + fld.Value + Chr(13) + Chr(10)
                  '如果为字符类型则直接显示
              Else
                  Me.Text1.Text = Me.Text1.Text + Str(fld.Value) + Chr(13) + Chr(10)
                  '如果为其它类型则转换为字符类型后显示

              End If
          End If
          Me.Text1.Text = Me.Text1.Text + "--------------------------------" + Chr(13) + Chr(10)
       Next
       rsSchema.MoveNext
    Wend
    
    rsSchema.Close
    Set rsSchema = Nothing
    cn.Close
    Set cn = Nothing
    Set fld = Nothing
End Sub
