VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "菜单图标"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture4 
      Height          =   495
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   1800
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   1200
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   3000
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Menu File 
      Caption         =   "文件"
      Index           =   0
      Begin VB.Menu New 
         Caption         =   "新建"
         Index           =   0
      End
      Begin VB.Menu Open 
         Caption         =   "打开"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "保存"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu Print 
         Caption         =   "打印"
         Index           =   3
         Shortcut        =   ^P
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
End Sub

Private Sub Form_Load()
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
    '隐藏各Picture Box控件
    Picture1.AutoSize = True
    Picture2.AutoSize = True
    Picture3.AutoSize = True
    Picture4.AutoSize = True
    '使各Picture Box控件能够根据图片自动改变大小
     
     hMenu& = GetMenu(Form1.hwnd)
    '得到窗口中菜单的句柄
    hSubMenu& = GetSubMenu(hMenu&, 0)
    '得到第一个子菜单File的句柄
    
    hID& = GetMenuItemID(hSubMenu&, 0)
    Picture1.Picture = LoadPicture(App.Path + "\New.bmp")
    SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, _
                        Picture1.Picture, Picture1.Picture
    '得到File菜单中的第一个菜单项的ID
    '然后为其添加图标
    '也可以添加两个图片
    '一个表示菜单项被选中状态
    '另一个表示菜单项没有被选中
    
        hID& = GetMenuItemID(hSubMenu&, 1)
    Picture2.Picture = LoadPicture(App.Path + "\Open.bmp")
    SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, _
                        Picture2.Picture, Picture2.Picture
    '得到File菜单中的第二个菜单项的ID并为其添加图标
    
    hID& = GetMenuItemID(hSubMenu&, 2)
    Picture3.Picture = LoadPicture(App.Path + "\Save.bmp")
    SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, _
                        Picture3.Picture, Picture3.Picture
    '得到File菜单中的第三个菜单项的ID并为其添加图标
    
    hID& = GetMenuItemID(hSubMenu&, 3)
    Picture4.Picture = LoadPicture(App.Path + "\Print.bmp")
    SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, _
                        Picture4.Picture, Picture4.Picture
        '得到File菜单中的第四个菜单项的ID并为其添加图标
End Sub
