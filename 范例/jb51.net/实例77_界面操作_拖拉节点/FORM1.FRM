VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   1785
   ClientTop       =   1455
   ClientWidth     =   5415
   LinkTopic       =   "Form10"
   ScaleHeight     =   4875
   ScaleWidth      =   5415
   Begin ComctlLib.TreeView OrgTree 
      DragIcon        =   "Form1.frx":0000
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7646
      _Version        =   327682
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   3
      Left            =   4800
      Picture         =   "Form1.frx":0152
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   2
      Left            =   4800
      Picture         =   "Form1.frx":02A4
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   1
      Left            =   4800
      Picture         =   "Form1.frx":03F6
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   6
      Left            =   5160
      Picture         =   "Form1.frx":0548
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   5
      Left            =   5160
      Picture         =   "Form1.frx":064A
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   4
      Left            =   5160
      Picture         =   "Form1.frx":098C
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   3
      Left            =   4800
      Picture         =   "Form1.frx":0CCE
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   2
      Left            =   4800
      Picture         =   "Form1.frx":0DD0
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image TreeImage 
      Height          =   240
      Index           =   1
      Left            =   4800
      Picture         =   "Form1.frx":1112
      Top             =   1440
      Width           =   240
   End
   Begin ComctlLib.ImageList TreeImages 
      Left            =   4800
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ObjectType
    otNone = 0
    otFactory = 1
    otGroup = 2
    otPerson = 3
    otFactory2 = 4
    otGroup2 = 5
    otPerson2 = 6
End Enum

Private SourceNode As Object
Private SourceType As ObjectType
Private TargetNode As Object
' ***********************************************
' Return the node's object type.
' ***********************************************
Private Function NodeType(test_node As Node) As ObjectType
    If test_node Is Nothing Then Exit Function
    Select Case Left$(test_node.Key, 1)
        Case "f"
            NodeType = otFactory
        Case "g"
            NodeType = otGroup
        Case "p"
            NodeType = otPerson
    End Select
End Function
' ***********************************************
' Prepare the ImageList and TreeView controls.
' ***********************************************
Private Sub Form_Load()
Dim i As Integer
Dim factory As Node
Dim group As Node
Dim person As Node

    ' Load pictures into the ImageList.
    For i = 1 To 6
        TreeImages.ListImages.Add , , TreeImage(i).Picture
    Next i
    
    ' Attach the TreeView to the ImageList.
    OrgTree.ImageList = TreeImages
    
    ' Create some nodes.
    Set factory = OrgTree.Nodes.Add(, , "f R & D", "R & D", otFactory, otFactory2)
    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Engineering", "Engineering", otGroup, otGroup2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Cameron, Charlie", "Cameron, Charlie", otPerson, otPerson2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Davos, Debbie", "Davos, Debbie", otPerson, otPerson2)
    person.EnsureVisible
    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Test", "Test", otGroup, otGroup2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Able, Andy", "Andy, Able", otPerson, otPerson2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Baker, Betty", "Baker, Betty", otPerson, otPerson2)
    person.EnsureVisible
    
    Set factory = OrgTree.Nodes.Add(, , "f Sales & Support", "Sales & Support", otFactory, otFactory2)
    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Showroom Sales", "Showroom Sales", otGroup, otGroup2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Gaines, Gina", "Gaines, Gina", otPerson, otPerson2)
    person.EnsureVisible
    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Field Service", "Field Service", otGroup, otGroup2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Helms, Harry", "Helms, Harry", otPerson, otPerson2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Ives, Irma", "Ives, Irma", otPerson, otPerson2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Jackson, Josh", "Jackson, Josh", otPerson, otPerson2)
    person.EnsureVisible
    Set group = OrgTree.Nodes.Add(factory, tvwChild, "g Customer Support", "Customer Support", otGroup, otGroup2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Klug, Karl", "Klug, Karl", otPerson, otPerson2)
    Set person = OrgTree.Nodes.Add(group, tvwChild, "p Landau, Linda", "Landau, Linda", otPerson, otPerson2)
    person.EnsureVisible
End Sub
' ***********************************************
' Make the TreeView as large as possible.
' ***********************************************
Private Sub Form_Resize()
    OrgTree.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


' ***********************************************
' Save the node pressed so we can drag it later.
' ***********************************************
Private Sub OrgTree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set SourceNode = OrgTree.HitTest(x, y)
End Sub

' ***********************************************
' Start a drag if one is not in progress.
' ***********************************************
Private Sub OrgTree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SourceNode Is Nothing Then Exit Sub

    If Button = vbLeftButton Then
        ' Start a new drag. Note that we do not get
        ' other MouseMove events while the drag is
        ' in progress.
        
        ' See what node we are dragging.
        SourceType = NodeType(SourceNode)

        ' Select this node. When no node is highlighted,
        ' this node will be displayed as selected. That
        ' shows where it will land if dropped.
        Set OrgTree.SelectedItem = SourceNode

        ' Set the drag icon for this source.
        OrgTree.DragIcon = IconImage(SourceType)
        OrgTree.Drag vbBeginDrag
    End If
End Sub
' ***********************************************
' The user is dropping. See if the drop is valid.
' ***********************************************
Private Sub OrgTree_DragDrop(Source As Control, x As Single, y As Single)
    If SourceNode Is Nothing Then Exit Sub

    If Not (OrgTree.DropHighlight Is Nothing) Then
        ' It's a valid drop. Set source node's
        ' parent to be the target node.
        Set SourceNode.Parent = OrgTree.DropHighlight
        Set OrgTree.DropHighlight = Nothing
    End If

    Set SourceNode = Nothing
    SourceType = otNone
End Sub
' ***********************************************
' The mouse is being dragged over the control.
' Highlight the appropriate node.
' ***********************************************
Private Sub OrgTree_DragOver(Source As Control, x As Single, y As Single, State As Integer)
Dim target As Node
Dim highlight As Boolean

    If SourceNode Is Nothing Then Exit Sub

    ' See what node we're above.
    Set target = OrgTree.HitTest(x, y)
    
    ' If it's the same as last time, do nothing.
    If target Is TargetNode Then Exit Sub
    Set TargetNode = target
    
    highlight = False
    If Not (TargetNode Is Nothing) Then
        ' See what kind of node were above.
        If NodeType(TargetNode) + 1 = SourceType Then _
            highlight = True
    End If
    
    If highlight Then
        Set OrgTree.DropHighlight = TargetNode
    Else
        Set OrgTree.DropHighlight = Nothing
    End If
End Sub

