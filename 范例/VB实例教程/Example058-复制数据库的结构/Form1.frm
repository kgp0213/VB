VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5415
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim sourceDB As DAO.Database
    Set sourceDB = OpenDatabase("D:\Microsoft Visual Studio\VB98\BIBLIO.MDB")
    Dim destDB As DAO.Database
    Set destDB = CreateDatabase("D:\Microsoft Visual Studio\VB98\Copy of BIBLIO.MDB", dbLangGeneral)
    
    Dim sourceT As TableDef
    For Each sourceT In sourceDB.TableDefs
        If sourceT.Attributes = 0 Then
            Dim destT As TableDef
            Set destT = destDB.CreateTableDef(sourceT.Name, sourceT.Attributes, _
                               sourceT.SourceTableName, sourceT.Connect)
                               
            Dim sourceF As Field
            For Each sourceF In sourceT.Fields
                Dim destF As Field
                Set destF = destT.CreateField(sourceF.Name, sourceF.Type, sourceF.Size)
                destT.Fields.Append destF
            Next
            destDB.TableDefs.Append destT
        End If
    Next
    destDB.Close
    sourceDB.Close
End Sub
