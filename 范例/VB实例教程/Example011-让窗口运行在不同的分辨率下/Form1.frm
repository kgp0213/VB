VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "分辨率设置"
   ClientHeight    =   3945
   ClientLeft      =   4125
   ClientTop       =   2745
   ClientWidth     =   6015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "设置分辨率"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dxSet As New DirectX7
'声明DirectX7对象
Dim ddSet As DirectDraw7
'声明DirectDraw7对象
Dim DisModesEnum As DirectDrawEnumModes
'声明DirectDrawEnumModes对象
Dim dds2 As DDSURFACEDESC2

'以下四个数组存储显示模式的相关数据
Dim lntWid(100) As Integer
'存储宽度
Dim lntHig(100) As Integer
'存储高度
Dim lntBB(100) As Integer
'存储颜色位数
Dim lntRefR(100) As Integer
'存储刷新频率
Private Sub Command1_Click()
    Dim intSel As Integer
    intSel = List1.ListIndex
    '取得在列表框中选择的显示模式
    Call ddSet.SetCooperativeLevel(Me.hWnd, DDSCL_ALLOWMODEX Or DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE)
    '设置协作水平
    ddSet.SetDisplayMode lntWid(intSel), lntHig(intSel), lntBB(intSel), lntRefR(intSel), DDSDM_DEFAULT
    '设置显示模式
    Me.Height = 0
    Me.Width = 3660
    Me.Caption = "关闭窗口恢复原来的分辨率"
End Sub

Private Sub Form_Load()
    Set ddSet = dxSet.DirectDrawCreate("")
    'dxSet建立DirectDraw对象ddSet
    ddSet.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
    '设置协作水平
    Set DisModesEnum = ddSet.GetDisplayModesEnum(DDEDM_DEFAULT, dds2)
    'DisModesEnum获得支持的显示模式
     
     For i = 1 To DisModesEnum.GetCount()
        DisModesEnum.GetItem i, dds2
        '将指定的显示模式的相关数据存入dds2
        lntWid(i) = dds2.lWidth
        '将该显示模式下的宽度存入数组lntWid
        lntHig(i) = dds2.lHeight
        '将该显示模式下的高度存入数组lntHig
        lntBB(i) = dds2.ddpfPixelFormat.lRGBBitCount
        '将该显示模式下的色彩深度存入数组lntBB
        lntRefR(i) = dds2.lRefreshRate
        '将该显示模式下的刷新率存入数组lntRefR
        List1.AddItem "显示模式：" + Str(i - 1) + _
                      "      宽度" + Str(lntWid(i)) + _
                      "      高度" + Str(lntHig(i)) + _
                      "      颜色位数" + Str(lntBB(i)) + _
                      "      刷新率" + Str(lntRefR(i))
    Next
    '在列表框中显示各种显示模式的宽度、高度、色彩深度、刷新率，并为各显示模式编号
End Sub

