VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "工具栏"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
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
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":0448
      Top             =   720
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "工具栏风格"
      Height          =   1455
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      Begin VB.OptionButton Option3 
         Caption         =   "标准风格"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Office风格"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "IE4风格"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   741
      ButtonWidth     =   1667
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
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
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "删除线"
            Key             =   "FontStrikthr"
            Description     =   "删除线"
            Object.ToolTipText     =   "删除线"
            ImageIndex      =   4
            Style           =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()
'Office风格的工具栏
    If Option1.Value = True Then
        Toolbar1.Style = tbrFlat
        Toolbar1.TextAlignment = tbrTextAlignRight
    End If
End Sub

Private Sub Option2_Click()
'IE4风格的工具栏
    If Option2.Value = True Then
        Toolbar1.TextAlignment = tbrTextAlignBottom
        Toolbar1.Style = tbrStandard
    End If
End Sub

Private Sub Option3_Click()
'标准风格的工具栏
    If Option1.Value = True Then
        Toolbar1.Style = tbrStandard
        Toolbar1.TextAlignment = tbrTextAlignRight
        Toolbar1.BorderStyle = ccFixedSingle
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
        '如果单击的按钮的Index属性为1
            If Button.Value = tbrPressed Then
                Text1.FontBold = True
                '如果此时按钮处于选中状态
              '将Text1的FontBold属性设置为True
            Else
                Text1.FontBold = False
                '如果此时按钮处于非选中状态
                '将Text1的FontBold属性设置为False
            End If
        Case 2
        '如果单击的按钮的Index属性为2
            If Button.Value = tbrPressed Then
            '如果此时按钮处于选中状态
               Text1.FontItalic = True
               '将Text1的FontItalic属性设置为True
            Else
                Text1.FontItalic = False
            End If
        Case 3
            If Button.Value = tbrPressed Then
               Text1.FontUnderline = True
            Else
                Text1.FontUnderline = False
            End If
        Case 4
            If Button.Value = tbrPressed Then
               Text1.FontStrikethru = True
            Else
                Text1.FontStrikethru = False
            End If
    End Select

End Sub

