VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   1200
      Top             =   120
   End
   Begin VB.Menu Pop 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu Reading 
         Caption         =   "阅读"
      End
      Begin VB.Menu Pleased 
         Caption         =   "高兴"
      End
      Begin VB.Menu Writing 
         Caption         =   "书写"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dot0 As IAgentCtlCharacterEx
Dim AgentRequest As IAgentCtlRequest
Private Sub Agent1_Click(ByVal CharacterID As String, _
                         ByVal Button As Integer, _
                         ByVal Shift As Integer, _
                         ByVal x As Integer, _
                         ByVal y As Integer)
'如果单击鼠标右键显示自定义的菜单
    If Button = 2 Then
       PopupMenu Pop
    End If
    
End Sub

Private Sub Form_Load()
'通过Set语句对其进行设置
    Set AgentRequest = Agent1.Characters.Load("Dot0", _
                       App.Path + "\merlin.acs")
    Set Dot0 = Agent1.Characters.Character("Dot0")
    Dot0.Show (0)
    Dot0.AutoPopupMenu = False
End Sub
Private Sub Pleased_Click()
'播放Pleased动作
    Dot0.StopAll
    Dot0.Play "Pleased"
End Sub

Private Sub Reading_Click()
'播放Reading动作
    Dot0.StopAll
    Dot0.Play "Reading"
End Sub

Private Sub Writing_Click()
'播放Writing动作
    Dot0.StopAll
    Dot0.Play "Writing"
End Sub

