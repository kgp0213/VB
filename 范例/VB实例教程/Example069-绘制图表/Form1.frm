VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4125
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSChart20Lib.MSChart MSChart1 
      Bindings        =   "Form1.frx":0000
      Height          =   3135
      Left            =   0
      OleObjectBlob   =   "Form1.frx":0015
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.MSChart1.chartType = VtChChartType3dCombination
    Me.MSChart1.ColumnCount = 5
    Me.MSChart1.RowCount = 3
    Dim column, row As Integer
    For column = 1 To Me.MSChart1.ColumnCount
        For row = 1 To Me.MSChart1.RowCount
            Me.MSChart1.column = column
            Me.MSChart1.row = row
            Me.MSChart1.Data = Rnd * Rnd
        Next
    Next
End Sub
