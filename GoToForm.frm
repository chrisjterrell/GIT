VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GoToForm 
   Caption         =   "Choose your Segment"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4080
   OleObjectBlob   =   "GoToForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GoToForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBsheet_Change()
    On Error Resume Next
    nme = CBsheet
    Sheets(nme).Select

    'Remove comment to to Hide the Form
    'GoToForm.Hide
End Sub
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            GoToForm.CBsheet.AddItem ws.Name
        End If
    Next
End Sub
