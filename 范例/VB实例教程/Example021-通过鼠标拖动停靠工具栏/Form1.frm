VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "工具栏"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   5070
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0224
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0336
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      DragMode        =   1  'Automatic
      Height          =   630
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   1111
      ButtonWidth     =   1138
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "黑体"
            Key             =   "FontBold"
            Description     =   "黑体"
            Object.ToolTipText     =   "黑体"
            ImageIndex      =   1
            Style           =   1
            Object.Width           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "斜体"
            Key             =   "FontItalic "
            Description     =   "斜体"
            Object.ToolTipText     =   "斜体"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "下划线"
            Key             =   "FontULine"
            Description     =   "下划 线"
            Object.ToolTipText     =   "下划 线"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除线"
            Key             =   "FontStrikthr"
            Description     =   "删除线"
            Object.ToolTipText     =   "删除线"
            ImageIndex      =   3
            Style           =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Toolbar1.DragMode = vbAutomatic
    Toolbar1.TextAlignment = tbrTextAlignBottom
End Sub

Private Sub Form_DragDrop(Source As Control, _
                        X As Single, Y As Single)
If Source Is Toolbar1 Then
    If X > Form1.ScaleWidth - 200 And _
       (Y > 200 And Y < Form1.ScaleHeight - 200) Then
        Toolbar1.Align = vbAlignRight
    End If
    '如果此时位于窗口的右边缘则将工具栏停靠在窗口右侧
    If X < 200 And (Y > 200 And _
       Y < Form1.ScaleHeight - 5) Then
        Toolbar1.Align = vbAlignLeft
    End If
    If Y > Form1.ScaleHeight - 200 And _
       (X > 200 And X < Form1.ScaleWidth - 200) Then
        Toolbar1.Align = vbAlignBottom
    End If
    If Y < 200 And _
       (X > 200 And X < Form1.ScaleWidth - 200) Then
        Toolbar1.Align = vbAlignTop
    End If
    Toolbar1.Refresh
Else
   Exit Sub
End If
End Sub

