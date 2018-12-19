Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object

Sub Main()
    Dim currentDoc As ModelDoc2
    Dim selMgr As SelectionMgr
    Dim selData As SelectData
    Dim selCount As Integer
    Dim comp As Component2
    Dim compCount As Integer
    Dim components() As Component2
    Dim i As Integer
    
    Set swApp = Application.SldWorks
    Set currentDoc = swApp.ActiveDoc
    If Not currentDoc Is Nothing Then
        If currentDoc.GetType = swDocASSEMBLY Then
            Set selMgr = currentDoc.SelectionManager
            selCount = selMgr.GetSelectedObjectCount2(-1)
            If selCount > 0 Then
                ReDim components(1 To selCount)
                compCount = 0
                For i = 1 To selCount
                    Set comp = selMgr.GetSelectedObjectsComponent4(i, -1)
                    If Not comp Is Nothing Then
                        compCount = compCount + 1
                        Set components(compCount) = comp
                    End If
                Next
                If compCount >= 1 Then
                    Set selData = selMgr.CreateSelectData
                    currentDoc.ClearSelection2 True  'Magic: Need clear before every copy
                    For i = 1 To compCount
                        selMgr.AddSelectionListObject components(i), selData
                        currentDoc.EditCopy
                        currentDoc.ClearSelection2 True
                        currentDoc.Paste
                    Next
                End If
            End If
        End If
    End If
End Sub
