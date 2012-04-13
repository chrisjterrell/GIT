Attribute VB_Name = "Module1"
Sub GoToMacro()
    GoToForm.Show
End Sub

Sub addMacro()

    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        Sheet1.ComboBox1.AddItem ws.Name
    Next
End Sub
