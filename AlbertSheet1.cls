VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Target.Count > 1 Then Exit Sub
If Not Intersect(Target, Range("D37:R999999")) Is Nothing And Range("D" & Target.Row).Value <> Empty Then
   Range("B5").Value = Target.Row
   Cont_Load
End If

End Sub
Sub SearchBox()
'PURPOSE: Filter Data on User-Determined Column & Text/Numerical value
'SOURCE: www.TheSpreadsheetGuru.com

Dim myButton As OptionButton
Dim SearchString As String
Dim ButtonName As String
Dim sht As Worksheet
Dim myField As Long
Dim DataRange As Range
Dim mySearch As Variant

'Load Sheet into A Variable
  Set sht = ActiveSheet

'Unfilter Data (if necessary)
  On Error Resume Next
    sht.ShowAllData
  On Error GoTo 0
  
'Filtered Data Range (include column heading cells)
  Set DataRange = sht.Range("A4:E31") 'Cell Range
  'Set DataRange = sht.ListObjects("Table1").Range 'Table

'Retrieve User's Search Input
  mySearch = sht.Shapes("UserSearch").TextFrame.Characters.Text 'Control Form
  'mySearch = sht.OLEObjects("UserSearch").Object.Text 'ActiveX Control
  'mySearch = sht.Range("A1").Value 'Cell Input
   
'Determine if user is searching for number or text
  If IsNumeric(mySearch) = True Then
    SearchString = "=" & mySearch
  Else
    SearchString = "=*" & mySearch & "*"
  End If
    
'Loop Through Option Buttons
  For Each myButton In sht.OptionButtons
    If myButton.Value = 1 Then
      ButtonName = myButton.Text
      Exit For
    End If
  Next myButton
   
'Determine Filter Field
  On Error GoTo HeadingNotFound
    myField = Application.WorksheetFunction.Match(ButtonName, DataRange.Rows(1), 0)
  On Error GoTo 0
   
'Filter Data
  DataRange.AutoFilter _
    Field:=myField, _
    Criteria1:=SearchString, _
    Operator:=xlAnd
  
'Clear Search Field
  sht.Shapes("UserSearch").TextFrame.Characters.Text = "" 'Control Form
  'sht.OLEObjects("UserSearch").Object.Text = "" 'ActiveX Control
  'sht.Range("A1").Value = "" 'Cell Input
  
Exit Sub

'ERROR HANDLERS
HeadingNotFound:
  MsgBox "The column heading [" & ButtonName & "] was not found in cells " & DataRange.Rows(1).Address & ". " & _
    vbNewLine & "Please check for possible typos.", vbCritical, "Header Name Not Found!"
  
End Sub
