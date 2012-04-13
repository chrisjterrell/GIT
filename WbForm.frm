VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WbForm 
   Caption         =   "UserForm1"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   OleObjectBlob   =   "WbForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WbForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button_Click()
    If Me.CB.ListIndex = -1 Then Exit Sub
    nme = Me.CB
    Set fswb = Workbooks(nme)
    Me.Hide
End Sub

Private Sub CB_Change()
    nme = Me.CB
    Windows(nme).Activate
End Sub

Private Sub UserForm_Initialize()

    
    For Each fswb In Application.Workbooks
        If wb.Name <> fswb.Name Then
            Me.CB.AddItem fswb.Name
        End If
    
    Next
End Sub
