VERSION 5.00
Begin VB.UserControl MyMenu 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2280
   PropertyPages   =   "MyMenu.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   2280
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4085
      Left            =   500
      TabIndex        =   0
      Top             =   0
      Width           =   2740
      Begin VB.Label mLabel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Top             =   45
         Width           =   2775
      End
   End
   Begin VB.Image mImage 
      Height          =   240
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "MyMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'Download by http://down.liehuo.net
Public Event Click(SelectedItem As Integer) '定义菜单项单击事件

Dim mItemSum As Integer       '当前菜单项总数（运行时只读）
Dim sItemSum As Integer       '上一次菜单项总数
Dim mCaption() As String      '菜单项文本
Dim mindex As Integer         '上一次鼠标经过的菜单编号
Dim mStartColor As Long       '渐变起始色
Dim mCeaseColor As Long       '渐变结束色
Dim i As Integer

Private Sub mLabel_Click(Index As Integer)
If left(mLabel(Index).Caption, 1) <> "-" Then RaiseEvent Click(Index) '转发单击事件
End Sub

Private Sub UserControl_Click()
If left(mLabel(mindex).Caption, 1) <> "-" Then RaiseEvent Click(mindex) '转发单击事件
End Sub

Private Sub mLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SetMenu (Index)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetMenu (Y * sItemSum \ Height + 1)
End Sub

Private Sub SetMenu(Index As Integer)
If left(mLabel(Index).Caption, 1) = "-" Then Exit Sub
mLabel(mindex).BackColor = &HFFFFC0
mLabel(mindex).ForeColor = 0
mLabel(Index).BackColor = &H8000&
mLabel(Index).ForeColor = &HFFFFFF
mindex = Index
End Sub

Public Property Get Caption(ByVal Index As Integer) As String '各菜单项的文本
Attribute Caption.VB_Description = "返回/设置菜单图像区背景颜色渐变的结束色。"
Attribute Caption.VB_ProcData.VB_Invoke_Property = "myMenuPage"
Caption = mCaption(Index)
End Property

Public Property Let Caption(ByVal Index As Integer, ByVal newVal As String)
mCaption(Index) = newVal
mLabel(Index).Caption = newVal
PropertyChanged "sCaption" & Index
End Property

Public Property Get Font() As Font
Set Font = mLabel(1).Font
End Property

Public Property Set Font(ByVal NewFont As Font) '注意：字体只能用“Set”而不能用“Let”
Set mLabel(1).Font = NewFont
PropertyChanged "Font"
End Property

Public Property Get ItemS() As Integer '菜单项总数
Attribute ItemS.VB_Description = "菜单项总数（运行时只读）。"
Attribute ItemS.VB_ProcData.VB_Invoke_Property = "myMenuPage"
ItemS = mItemSum
End Property

Public Property Let ItemS(ByVal NewValue As Integer) '运行时只读
If Not Ambient.UserMode And NewValue > 0 Then
  mItemSum = NewValue
  ReDim Preserve mCaption(1 To mItemSum) '重新设置文本数组
  PropertyChanged "ItemSum"
  DrawMenu
  sItemSum = mItemSum
End If
End Property

Public Property Get CeaseColor() As OLE_COLOR
Attribute CeaseColor.VB_Description = "返回/设置菜单图像区背景颜色渐变的结束色。"
CeaseColor = mCeaseColor
End Property

Public Property Let CeaseColor(ByVal NewColor As OLE_COLOR)
mCeaseColor = NewColor
PropertyChanged "CeaseColor"
DrawMenu
End Property

Public Property Get StartColor() As OLE_COLOR
Attribute StartColor.VB_Description = "返回/设置菜单图像区背景颜色渐变的起始色。"
StartColor = mStartColor
End Property

Public Property Let StartColor(ByVal NewColor As OLE_COLOR)
mStartColor = NewColor
PropertyChanged "StartColor"
DrawMenu
End Property

Private Sub UserControl_Initialize() '该事件有点类似于窗体页面的Form_Load事件
mindex = 1
End Sub

Private Sub UserControl_InitProperties() '该事件只在刚将控件画到窗体上时运行一次
mItemSum = 1: sItemSum = 1
ReDim mCaption(1 To 1) As String
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  mStartColor = .ReadProperty("StartColor", &HFFFFC0)
  mCeaseColor = .ReadProperty("CeaseColor", &H8000&)
  Set mLabel(1).Font = .ReadProperty("Font", Ambient.Font)
  mItemSum = .ReadProperty("ItemSum", 1)
  sItemSum = mItemSum
  ReDim Preserve mCaption(1 To mItemSum)
  For i = 1 To mItemSum: mCaption(i) = .ReadProperty("sCaption" & i, ""): Next
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  .WriteProperty "StartColor", mStartColor, &HFFFFC0
  .WriteProperty "CeaseColor", mCeaseColor, &H8000&
  .WriteProperty "Font", mLabel(1).Font, Ambient.Font
  .WriteProperty "ItemSum", mItemSum, 1
  For i = 1 To mItemSum: .WriteProperty "sCaption" & i, mCaption(i), "": Next
End With
End Sub

Private Sub UserControl_Show()
DrawMenu
End Sub

Private Sub DrawMenu() '绘制菜单
Dim iPath As String '图标路径
Height = 270 * mItemSum
Frame1.Move 500, 0, Width - 500, Height
mLabel(1).Move 15, 30, Frame1.Width
Cls
mLabel(1).BackColor = &HFFFFC0
mLabel(1).ForeColor = 0
颜色渐变
If mItemSum > 1 Then
  If mLabel.Count > 1 Then For i = 2 To sItemSum: Unload mLabel(i): Next '如果有上一次建立的标签控件数组，就删除
  For i = 2 To mItemSum: Load mLabel(i): mLabel(i).Move 15, (i - 1) * 265 + 30: mLabel(i).Visible = True: Next '建立当前的标签控件数组
  For i = 1 To mItemSum
    mLabel(i).Caption = mCaption(i)
    iPath = App.Path & "\图片\10" & i & ".ico"
    If Len(Dir(iPath)) Then mImage.Picture = LoadPicture(iPath): PaintPicture mImage, 100, (i - 1) * 265 + 30
  Next
End If
End Sub

Private Sub 颜色渐变()
Dim rSta, gSta, bSta, rEnd, gEnd, bEnd, rInfo, gInfo, bInfo
rSta = mStartColor Mod 256: gSta = mStartColor \ 256 Mod 256: bSta = mStartColor \ 256 \ 256
rEnd = mCeaseColor Mod 256: gEnd = mCeaseColor \ 256 Mod 256: bEnd = mCeaseColor \ 256 \ 256
rInfo = (rEnd - rSta) / Height: gInfo = (gEnd - gSta) / Height: bInfo = (bEnd - bSta) / Height
For i = 0 To Height: Line (0, i)-(500, i), RGB(rSta + i * rInfo, gSta + i * gInfo, bSta + i * bInfo): Next
End Sub
