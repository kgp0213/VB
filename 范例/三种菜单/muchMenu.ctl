VERSION 5.00
Begin VB.UserControl muchMenu 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FDEAD9&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2280
   FillColor       =   &H00FF8080&
   PropertyPages   =   "muchMenu.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   2280
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   450
      Left            =   500
      TabIndex        =   0
      Top             =   0
      Width           =   1905
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
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "muchMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'Download by http://down.liehuo.net
Public Event Click(SelectedItem As Integer) '定义菜单项单击事件

Dim mCaption(1 To 10, 1 To 20) As String '菜单项文本
Dim mItemSum(1 To 10) As Integer '当前菜单层的菜单项总数（运行时只读）
Dim sRep As Integer              '菜单层总数
Dim vRep As Integer              '当前菜单层
Dim sItemSum As Integer          '当前菜单层的上一次菜单项总数
Dim mindex As Integer            '上一次鼠标经过的菜单编号
Dim mStartColor As Long          '渐变起始色
Dim mCeaseColor As Long          '渐变结束色
Dim i As Integer

Private Sub mLabel_Click(Index As Integer)
If left(mLabel(Index).Caption, 1) <> "-" Then RaiseEvent Click(Index) '转发单击事件
End Sub

Private Sub UserControl_Click()
If left(mLabel(mindex).Caption, 1) <> "-" Then RaiseEvent Click(mindex) '转发单击事件
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetMenu (Y * sItemSum \ Height + 1)
End Sub

Private Sub mLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SetMenu (Index)
End Sub

Private Sub SetMenu(Index As Integer)
If left(mLabel(Index).Caption, 1) = "-" Then Exit Sub
mLabel(mindex).BackColor = &HFFFFC0
mLabel(mindex).ForeColor = 0
mLabel(Index).BackColor = &H8000&
mLabel(Index).ForeColor = &HFFFFFF
mindex = Index
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "返回/设置菜单项文本的字体。"
Set Font = mLabel(1).Font
End Property

Public Property Set Font(ByVal NewFont As Font) '注意：字体只能用“Set”而不能用“Let”
Set mLabel(1).Font = NewFont
PropertyChanged "Font"
End Property

Public Property Get RepeatCount() As Integer '菜单层总数
Attribute RepeatCount.VB_Description = "返回/设置菜单控件的总层数，运行时只读。"
RepeatCount = sRep
End Property

Public Property Let RepeatCount(ByVal NewValue As Integer) '运行时只读
If Not Ambient.UserMode And NewValue > 0 And NewValue < 11 Then
  sRep = NewValue
  PropertyChanged "RepeatCount"
End If
End Property

Public Property Get RepeatCurrent() As Integer '当前菜单层
Attribute RepeatCurrent.VB_Description = "返回/设置当前的菜单层编号。"
RepeatCurrent = vRep
End Property

Public Property Let RepeatCurrent(ByVal NewValue As Integer)
If NewValue > 0 And NewValue <= sRep Then
  mindex = 1
  vRep = NewValue
  PropertyChanged "RepeatCurrent"
  DrawMenu
End If
End Property

Public Property Get ItemS() As Integer '当前菜单层的菜单项总数
Attribute ItemS.VB_Description = "返回/设置当前菜单层的菜单项总数，运行时只读。"
ItemS = mItemSum(vRep)
End Property

Public Property Let ItemS(ByVal NewValue As Integer) '运行时只读
If Not Ambient.UserMode And NewValue > 0 And NewValue < 21 Then
  mItemSum(vRep) = NewValue
  PropertyChanged "ItemSum"
  DrawMenu
End If
End Property

Public Property Get Caption(ByVal Index As Integer) As String '当前菜单层各菜单项的文本
Attribute Caption.VB_Description = "返回/设置菜单图像区背景颜色渐变的结束色。"
Caption = mCaption(vRep, Index)
End Property

Public Property Let Caption(ByVal Index As Integer, ByVal newVal As String)
mCaption(vRep, Index) = newVal
mLabel(Index).Caption = newVal
PropertyChanged "sCaption"
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
For i = 1 To 10: mItemSum(i) = 1: Next
mindex = 1: sItemSum = 1: sRep = 1: vRep = 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim j As Integer
With PropBag
  mStartColor = .ReadProperty("StartColor", &HFFFFC0)
  mCeaseColor = .ReadProperty("CeaseColor", &H8000&)
  Set mLabel(1).Font = .ReadProperty("Font", Ambient.Font)
  sRep = .ReadProperty("RepeatCount", 1)
  vRep = .ReadProperty("RepeatCurrent", 1)
  For i = 1 To sRep: mItemSum(i) = .ReadProperty("ItemSum" & i, 1): Next
  For j = 1 To sRep: For i = 1 To mItemSum(j): mCaption(j, i) = .ReadProperty("sCaption" & j * 10 & i, ""): Next: Next
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim j As Integer
With PropBag
  .WriteProperty "StartColor", mStartColor, &HFFFFC0
  .WriteProperty "CeaseColor", mCeaseColor, &H8000&
  .WriteProperty "Font", mLabel(1).Font, Ambient.Font
  .WriteProperty "RepeatCount", sRep, 1
  .WriteProperty "RepeatCurrent", vRep, 1
  For i = 1 To sRep: .WriteProperty "ItemSum" & i, mItemSum(i), 1: Next
  For j = 1 To sRep: For i = 1 To mItemSum(j): .WriteProperty "sCaption" & j * 10 & i, mCaption(j, i), "": Next: Next
End With
End Sub

Private Sub UserControl_Show()
DrawMenu
End Sub

Private Sub DrawMenu() '绘制菜单
Dim iPath As String '图标路径
Height = 270 * mItemSum(vRep)
Frame1.Move 500, 0, Width - 500, Height
mLabel(1).Move 15, 30, Frame1.Width
Cls
mLabel(1).BackColor = &HFFFFC0
mLabel(1).ForeColor = 0
颜色渐变
If mItemSum(vRep) > 0 Then
  If mLabel.Count > 1 Then For i = 2 To sItemSum: Unload mLabel(i): Next '如果有上一次建立的标签控件数组，就删除
  For i = 2 To mItemSum(vRep): Load mLabel(i): mLabel(i).Move 15, (i - 1) * 265 + 30: mLabel(i).Visible = True: Next '建立当前的标签控件数组
  For i = 1 To mItemSum(vRep)
    mLabel(i).Caption = mCaption(vRep, i)
    iPath = App.Path & "\图片\" & vRep * 10 & i & ".ico"
    If Len(Dir(iPath)) Then mImage.Picture = LoadPicture(iPath): PaintPicture mImage, 100, (i - 1) * 265 + 30
  Next
End If
sItemSum = mItemSum(vRep)
End Sub

Private Sub 颜色渐变()
Dim rSta, gSta, bSta, rEnd, gEnd, bEnd, rInfo, gInfo, bInfo
rSta = mStartColor Mod 256: gSta = mStartColor \ 256 Mod 256: bSta = mStartColor \ 256 \ 256
rEnd = mCeaseColor Mod 256: gEnd = mCeaseColor \ 256 Mod 256: bEnd = mCeaseColor \ 256 \ 256
rInfo = (rEnd - rSta) / Height: gInfo = (gEnd - gSta) / Height: bInfo = (bEnd - bSta) / Height
For i = 0 To Height: Line (0, i)-(500, i), RGB(rSta + i * rInfo, gSta + i * gInfo, bSta + i * bInfo): Next
End Sub
