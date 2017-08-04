VERSION 5.00
Begin VB.UserControl MenuControl 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   ScaleHeight     =   960
   ScaleWidth      =   1800
   Begin VB.PictureBox menuitem 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   320
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "MenuControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type menuRect  '创建菜单项矩形，鼠标移动时，菜单项就出现在对应的矩形位置
  left As Integer
  top As Integer
  right As Integer
  bottom As Integer
End Type

Public Event Click(SelectedItem As Integer) '定义菜单项鼠标点击事件

Dim SelectedItem As Integer  '被选中的菜单项编号
Attribute SelectedItem.VB_VarDescription = "被选中的菜单项（只读）。"
Dim mItemSum As Integer      '菜单项总数（运行时只读）
Dim mRects() As menuRect     '菜单项矩形数组，用来记录所有菜单项的位置
Dim mStartColor As Long      '渐变起始色
Dim mCeaseColor As Long      '渐变结束色

Private Sub menuitem_Click()
RaiseEvent Click(SelectedItem) '转发单击事件
End Sub

Public Property Get ItemSum() As Integer '菜单项总数
Attribute ItemSum.VB_Description = "菜单项总数（运行时只读）。"
ItemSum = mItemSum
End Property

Public Property Let ItemSum(ByVal vNewValue As Integer) '该属性运行时只读
If Not Ambient.UserMode And vNewValue > 1 Then
  mItemSum = vNewValue
  PropertyChanged "ItemSum"
  UserControl_Resize
End If
End Property

Public Property Get CeaseColor() As OLE_COLOR
Attribute CeaseColor.VB_Description = "返回/设置菜单背景颜色渐变的结束色。"
CeaseColor = mCeaseColor
End Property

Public Property Let CeaseColor(ByVal NewColor As OLE_COLOR)
mCeaseColor = NewColor
PropertyChanged "CeaseColor"
颜色渐变
End Property

Public Property Get StartColor() As OLE_COLOR
Attribute StartColor.VB_Description = "返回/设置菜单背景颜色渐变的起始色。"
StartColor = mStartColor
End Property

Public Property Let StartColor(ByVal NewColor As OLE_COLOR)
mStartColor = NewColor
PropertyChanged "StartColor"
颜色渐变
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
  mItemSum = .ReadProperty("ItemSum", 1)
  mStartColor = .ReadProperty("StartColor", &HFFFFC0)
  mCeaseColor = .ReadProperty("CeaseColor", &H8000&)
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  .WriteProperty "ItemSum", mItemSum, 1
  .WriteProperty "StartColor", mStartColor, &HFFFFC0
  .WriteProperty "CeaseColor", mCeaseColor, &H8000&
End With
End Sub

Private Sub UserControl_Show()
PopAndLoadpic
End Sub

Private Sub UserControl_InitProperties()
mItemSum = 1
mStartColor = &HFFFFC0
mCeaseColor = &H8000&
End Sub

Private Sub UserControl_Resize() '调整控件大小
UserControl.Height = 320 * mItemSum
menuitem.Width = UserControl.Width
End Sub

Sub PopAndLoadpic() '生成菜单项矩形
Dim i As Integer
ReDim mRects(mItemSum) '重新定义菜单项矩形总数

For i = 1 To mItemSum
  mRects(i).left = 0       '菜单项矩形左边距
  mRects(i).right = menuitem.Width '菜单项矩形右边距
  mRects(i).top = IIf(i = 1, 0, mRects(i - 1).bottom) '如果是第一个菜单项，其矩形顶距＝0，否则＝上一个菜单项矩形的底边
  mRects(i).bottom = mRects(i).top + menuitem.Height  '菜单项底边高度
Next i
颜色渐变
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer, mName As String
For i = 1 To mItemSum
  If y > mRects(i).top And y <= mRects(i).bottom Then
    SelectedItem = i: mName = App.Path & "\图片\" & i & ".bmp"
    If Len(Dir(mName)) Then menuitem.Picture = LoadPicture(mName)
    menuitem.Move 0, mRects(i).top
    menuitem.Visible = True
    Exit Sub
  End If
Next
End Sub

Private Sub 颜色渐变()
Dim rSta, gSta, bSta, rEnd, gEnd, bEnd, rInfo, gInfo, bInfo
Dim i As Integer
rSta = mStartColor Mod 256: gSta = mStartColor \ 256 Mod 256: bSta = mStartColor \ 256 \ 256
rEnd = mCeaseColor Mod 256: gEnd = mCeaseColor \ 256 Mod 256: bEnd = mCeaseColor \ 256 \ 256
rInfo = (rEnd - rSta) / Height: gInfo = (gEnd - gSta) / Height: bInfo = (bEnd - bSta) / Height
For i = 0 To Height: Line (0, i)-(Width, i), RGB(rSta + i * rInfo, gSta + i * gInfo, bSta + i * bInfo): Next
End Sub
