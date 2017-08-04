VERSION 5.00
Begin VB.UserControl labelshape 
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   840
   ScaleWidth      =   1890
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ActiveX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Shape Shapeback 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "labelshape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event dbclick()

Private Sub Label1_DblClick()
RaiseEvent dbclick
End Sub


Private Sub UserControl_Initialize()
Debug.Print "初始化"
End Sub

Private Sub UserControl_InitProperties()
Debug.Print "属性初始化"
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Point(X, Y) = Shapeback.FillColor Then
RaiseEvent dbclick
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Debug.Print "开始写入属性"
Caption = PropBag.ReadProperty("caption", Extender.Name)
End Sub

Private Sub UserControl_Resize()
'#####################################
'控件的关键事件实例所用的代码
'Static test As Integer
'test = test + 1
'Debug.Print "改变" & test; "次"
'######################################
Shapeback.Move 0, 0, ScaleWidth, ScaleHeight
Label1.Move 0, (ScaleHeight - Label1.Height) / 2, ScaleWidth, ScaleHeight
End Sub

Private Sub UserControl_Terminate()
Debug.Print "中止"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Debug.Print "开始写入属性"
PropBag.WriteProperty Caption, "Caption", Extender.Name
End Sub

Public Property Get Caption() As String
Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
Label1.Caption = vNewValue
PropertyChanged "caption"
End Property
