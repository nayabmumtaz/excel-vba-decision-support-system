Attribute VB_Name = "Module1"
Dim Valid As Boolean

Sub RunSimulation()

Valid = True

ErrorCheck

If Valid = False Then
    Worksheets("Inputs").Select
    Exit Sub
End If

    With Sheets("Pharma")
        .Visible = True
        .Activate
    End With

' RunSimulation Macro

'Dimension Variables
Dim n As Integer
Dim z As Double
Dim c As Integer
Dim simRange As Range

n = Range("B34")
z = Range("B35")

For sim = 1 To 7

Range("B13") = Range("E42").Offset(sim, 0)

'Clear previous simulation
    Range("A43:B43").Select
    Range(Selection, Selection.End(xlDown)).ClearContents

' Create loop to complete n itterations
For c = 1 To n
    Range("A42").Offset(c, 0) = c
    Range("A42").Offset(c, 1) = Range("G32")
Next c

'Defines simRange
    Range("B43").Select
    Set simRange = Range(Selection, Selection.End(xlDown))

'Calculate the Stats
Range("E42").Offset(sim, 1) = Application.WorksheetFunction.Average(simRange)
Range("E42").Offset(sim, 2) = Application.WorksheetFunction.Min(simRange)
Range("E42").Offset(sim, 3) = Application.WorksheetFunction.Max(simRange)
Range("E42").Offset(sim, 4) = Application.WorksheetFunction.StDev(simRange)

Next sim

'HideInputs
    Sheets("Inputs").Visible = False
    
'HidePharma
  Sheets("Pharma").Visible = False

'ShowOutputs
    Sheets("Outputs").Visible = True
    Sheets("Outputs").Select
    

End Sub

Sub ErrorCheck()
    Dim ws As Worksheet
    Dim inputRange As Range, minCell As Range, maxCell As Range
    Dim meanCell As Range, stdCell As Range, discountRateCell As Range
    Dim iterationsCell As Range
    Dim cell As Range
    Dim errorMsg As String
    Dim allValid As Boolean
    Dim userResponse As VbMsgBoxResult

    Set ws = ThisWorkbook.Sheets("Inputs")
    Set inputRange = ws.Range("G14,G16,G22,I22,G26,G28,G30,G34,G38,G42,G47")
    Set minCell = ws.Range("G22")
    Set maxCell = ws.Range("I22")
    Set discountRateCell = ws.Range("G38")
    Set meanCell = ws.Range("G14")
    Set stdCell = ws.Range("G16")
    Set iterationsCell = ws.Range("G42")

    allValid = True
    errorMsg = "Input Validation Errors:" & vbCrLf

    ' Check for numeric errors, blank cells, and negatives
    For Each cell In inputRange
        If IsEmpty(cell.Value) Then
            allValid = False
            errorMsg = errorMsg & cell.Address & " is empty." & vbCrLf
        ElseIf Not IsNumeric(cell.Value) Then
            allValid = False
            errorMsg = errorMsg & cell.Address & " contains a non-numeric value (" & cell.Value & ")." & vbCrLf
        ElseIf cell.Value < 0 Then
            allValid = False
            errorMsg = errorMsg & cell.Address & " cannot contain a negative value (" & cell.Value & ")." & vbCrLf
        End If
    Next cell

    ' Check that Min is less than Max
    If Not IsEmpty(minCell.Value) And Not IsEmpty(maxCell.Value) Then
        If Not IsNumeric(minCell.Value) Or Not IsNumeric(maxCell.Value) Then
            allValid = False
            errorMsg = errorMsg & "Min or Max contains a non-numeric value." & vbCrLf
        ElseIf minCell.Value >= maxCell.Value Then
            allValid = False
            errorMsg = errorMsg & "Min value (" & minCell.Value & ") must be less than Max value (" & maxCell.Value & ")." & vbCrLf
        End If
    Else
        allValid = False
        errorMsg = errorMsg & "Min and Max values cannot be blank." & vbCrLf
    End If

   ' Check if Mean is greater than Standard Deviation
    If Not IsEmpty(meanCell.Value) And Not IsEmpty(stdCell.Value) Then
        If meanCell.Value <= stdCell.Value Then
            allValid = False
            errorMsg = errorMsg & "Mean value (" & meanCell.Value & ") must be greater than Standard Deviation value (" & stdCell.Value & ")." & vbCrLf
        End If
    End If
   
    If Not allValid Then
        Valid = False

        ' Error message
        ws.Activate
        userResponse = MsgBox(errorMsg & vbCrLf & "Press Cancel to fix the inputs or OK to dismiss this message.", vbExclamation + vbOKCancel, "Input Validation Errors")

        ' If user presses Cancel, stay on Inputs sheet
        If userResponse = vbCancel Then
            ws.Activate
            Exit Sub
        End If
    Else
        ' Once all inputs are valid
        MsgBox "All inputs are valid!", vbInformation, "Validation Successful"
    End If
End Sub






