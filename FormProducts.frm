VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProducts 
   Caption         =   "Report Filter"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6045
   OleObjectBlob   =   "FormProducts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ButtonProducts_Click()
    Application.ScreenUpdating = False
    ws.Columns.Hidden = False
    
    For lrw = 0 To Me.LBProducts.ListCount - 1
        If Me.LBProducts.Selected(lrw) = False Then
            nme = LBProducts.List(lrw)
            ecl = ws.Cells.SpecialCells(xlCellTypeLastCell).Column
            For rpcl = cl To ecl
                If nme = ws.Cells(rw, rpcl).Text Then
                    ws.Columns(rpcl).Hidden = True
                    rpcl = rpcl + 1
                    Do While ws.Cells(rw, rpcl).Text = "" And rpcl < ecl + 1

                        ws.Columns(rpcl).Hidden = True
                        rpcl = rpcl + 1
                    Loop
                End If
            Next
        End If
    Next
    
    For lrw = 0 To Me.LBmetric.ListCount - 1
        If Me.LBmetric.Selected(lrw) = False Then
            nme = LBmetric.List(lrw)
            ecl = ws.Cells.SpecialCells(xlCellTypeLastCell).Column
            For rpcl = cl To ecl
                If Right(nme, Len(ws.Cells(mrw, rpcl))) = ws.Cells(mrw, rpcl) Then
                    ws.Columns(rpcl).Hidden = True
                End If
            Next
        End If
    Next
    
    Me.Hide
    On Error Resume Next
    ActiveSheet.CB.ListIndex = -1
    ActiveSheet.CB.ListIndex = 0

    
    Unload Me
    Application.ScreenUpdating = True

End Sub

Private Sub BoxMetric_Click()
    If Me.BoxMetric.Caption = "Select All" Then
        
        For lrw = 0 To Me.LBmetric.ListCount - 1
            Me.LBmetric.Selected(lrw) = True
        Next
        
        Me.BoxMetric.Caption = "Unselect All"
    Else
        
        For lrw = 0 To Me.LBmetric.ListCount - 1
            Me.LBmetric.Selected(lrw) = False
        Next
        
        Me.BoxMetric.Caption = "Select All"
    End If
End Sub

Private Sub BoxProd_Click()
    If Me.BoxProd.Caption = "Select All" Then
        
        For lrw = 0 To Me.LBProducts.ListCount - 1
            Me.LBProducts.Selected(lrw) = True
        Next
        
        Me.BoxProd.Caption = "Unselect All"
    Else
        
        For lrw = 0 To Me.LBProducts.ListCount - 1
            Me.LBProducts.Selected(lrw) = False
        Next
        
        Me.BoxProd.Caption = "Select All"
    End If
End Sub

Private Sub CBcancel_Click()
    Unload Me
End Sub

Private Sub lbl_Click()

End Sub

Private Sub UserForm_Initialize()
        
    ecl = Application.WorksheetFunction.CountA(ws.Rows(mrw)) + 10
    
    Me.lbChan.Caption = lbl
    
    For rcl = cl To ecl
        If ws.Cells(rw, rcl) <> "" And ws.Cells(rw + 1, rcl) <> "" Then
            If ws.Cells(rw, rcl).Text = "" Then
                Me.LBProducts.AddItem Format(ws.Cells(rw, rcl), "mmmm yyyy")
            Else
                Me.LBProducts.AddItem ws.Cells(rw, rcl).Text
            End If
        End If
    Next
    
    For lrw = 0 To Me.LBProducts.ListCount - 1
        'tst = ws.Cells(8, (lrw * 13) + cl)
        mNme = Me.LBProducts.List(lrw)
        fnd = False
        wscl = cl
        Do While Cells(9, wscl) <> ""
            If Cells(8, wscl) <> "" Then prod = Cells(8, wscl).Text
            If prod = mNme And Columns(wscl).Hidden = False Then
                fnd = True
                Exit Do
            End If
            wscl = wscl + 1
        Loop
        
        If fnd = True Then
            Me.LBProducts.Selected(lrw) = True
        End If
    Next
    
    fvar = Cells(mrw, mcl)
    
    For rcl = cl To ecl
        If ws.Cells(mrw, rcl) <> "" And ws.Cells(mrw, rcl) <> "Totals" Then
            If fvar = ws.Cells(mrw, rcl) And rcl <> mcl Then Exit For
            Me.LBmetric.AddItem ws.Cells(mrw, rcl)
        End If
    Next
    
   
    For lrw = 0 To Me.LBmetric.ListCount - 1
        mNme = Me.LBmetric.List(lrw)
        fnd = False
        wscl = cl
        Do While Cells(9, wscl) <> ""
            If Cells(9, wscl) = mNme And Columns(wscl).Hidden = False Then
                fnd = True
                Exit Do
            End If
            wscl = wscl + 1
        Loop
        
        If fnd = True Then
            Me.LBmetric.Selected(lrw) = True
        End If
    Next
    
End Sub

