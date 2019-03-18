Attribute VB_Name = "ContactMacros"
Option Explicit
Dim ContRow As Long
Dim ContCol As Long

Sub Cont_Load()
With Sheet1
If .Range("B5").Value = Empty Then Exit Sub
.Range("B7").Value = True 'Set Contact Load to True
ContRow = .Range("B5").Value 'Contact Row
For ContCol = 4 To 18
     .Range(.Cells(35, ContCol).Value).Value = .Cells(ContRow, ContCol).Value
Next ContCol
    On Error Resume Next
     .Shapes("ThumbPic").Delete 'Delete thumbnail picture (if any)
    On Error GoTo 0
'If .Range("T8").Value <> Empty Then Cont_Display Thumb
.Range("B6").Value = False
.Shapes("ExitContGrp").Visible = msoCTrue
.Shapes("NewContGrp").Visible = msoFalse

.Range("B7").Value = False 'Set Contact Load to False

End With
End Sub


Sub Cont_New()
With Sheet1
.Range("B7").Value = True 'New Contact to true
.Range("B6").Value = True 'New Contact
.Range("E11,E13,E15,E17,E19,E21,E23,E25,I11,I13,I15,I17,I19,I21,I23").ClearContents
.Shapes("ExitContGrp").Visible = msoFalse
.Shapes("NewContGrp").Visible = msoCTrue

.Range("B7").Value = False 'New Contact to False
End With

End Sub



Sub Cont_Save()
With Sheet1
If .Range("E11").Value = Empty Then
   MsgBox "Please Don't Leave this Blank"
   Exit Sub
   End If
   ContRow = .Range("D9999").End(xlUp).Row + 1 'First Available Row
   For ContCol = 4 To 18
   .Cells(ContRow, ContCol).Value = .Range(.Cells(35, ContCol).Value).Value
   
   Next ContCol
   .Shapes("ExitContGrp").Visible = msoCTrue
   .Shapes("NewContGrp").Visible = msoFalse
   .Range("B6").Value = False
   



End With

End Sub


Sub Cont_Delete()
With Sheet1
If MsgBox("Are you sure you want to delete nga walay duha2?", vbYesNo, "Delete Profile") = vbNo Then Exit Sub
If .Range("B5").Value = Empty Then Exit Sub
ContRow = .Range("B5").Value
.Range(ContRow & ":" & ContRow).EntireRow.Delete
.Range("D37").Select




End With

End Sub

Sub Cont_CancelNew()
With Sheet1
If .Range("D37").Value <> Empty Then .Range("D37").Select



End With

End Sub
