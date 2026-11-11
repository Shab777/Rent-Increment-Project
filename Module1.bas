Attribute VB_Name = "Module1"
Option Explicit

Sub NewRentPaid()

Dim Response As Variant

Response = MsgBox("Did Tenant pay new rent?", vbYesNoCancel)

If Response = vbCancel Then

   Exit Sub
    
    ElseIf Response = vbYes Then
        
                 
        ActiveCell.EntireRow.Select
                    
        With Selection.Font
        .Color = -11489280
        .TintAndShade = 0
    End With
   
    ElseIf Response = vbNo Then
       
       ActiveCell.EntireRow.Select

       With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
End If


    
End Sub

