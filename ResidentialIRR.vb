Option Explicit

Public Sub ResidentialIRR()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim assumptions As Worksheet
    Dim spillRange As Range
    Dim inputRange As Range
    Dim outputRange As Range

    Dim data As Variant
    Dim output() As Variant
    Dim passings As Variant
    Dim estimate As Variant
    Dim irr As Variant

    Dim rowCount As Long
    Dim lastRow As Long
    Dim oldLastRow As Long
    Dim i As Long

    Dim previousCalculation As XlCalculation
    Dim completedSuccessfully As Boolean

    MsgBox _
        "Please don't use your computer while the script runs. " & _
        "Wait for the completion message.", _
        vbExclamation

    previousCalculation = Application.Calculation
    On Error GoTo CleanFail

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Resi_IRR_ALL_LOB")
    Set assumptions = wb.Worksheets("Assumptions")

    'Refresh the spill before determining its size.
    Application.Calculate

    'A2 is the anchor cell for the residential spill.
    On Error Resume Next
    Set spillRange = ws.Range("A2").SpillingToRange
    On Error GoTo CleanFail

    If spillRange Is Nothing Then
        Err.Raise vbObjectError + 1000, , _
                  "No spill range was found from Resi_IRR_ALL_LOB!A2."
    End If

    rowCount = spillRange.Rows.Count
    lastRow = spillRange.Row + rowCount - 1


    If ws.Range("Y2").HasSpill Then
        Err.Raise vbObjectError + 1001, , _
                  "Column Y is still part of a spill range."
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    'Clear old results, including rows left from a previously larger spill.
    oldLastRow = ws.Cells(ws.Rows.Count, "Y").End(xlUp).Row

    If oldLastRow >= 2 Then
        ws.Range("Y2:Y" & Application.Max(lastRow, oldLastRow)).ClearContents
    End If

    'Reset assumption inputs.
    With assumptions
        .Range("F7:F12").ClearContents
        .Range("F21:F27").ClearContents
        .Range("F7").Value2 = 6
    End With

    'Read columns A:W into memory.
    Set inputRange = ws.Range("A2:W" & lastRow)
    data = inputRange.Value2

    ReDim output(1 To rowCount, 1 To 1)

    For i = 1 To rowCount

        'K = Passings, Q = Total Estimate
        passings = ToNumber(data(i, 11))
        estimate = ToNumber(data(i, 17))

        If IsEmpty(passings) Or IsEmpty(estimate) Then
            output(i, 1) = "Unavailable"

        ElseIf CDbl(passings) <= 0 Then
            output(i, 1) = "Unavailable"

        Else
            assumptions.Range("F8").Value2 = CDbl(passings)
            assumptions.Range("F10").Value2 = CDbl(estimate)

            Application.Calculate

            irr = assumptions.Range("F15").value

            If IsError(irr) Or Not IsNumeric(irr) Then
                output(i, 1) = "Unavailable"
            Else
                output(i, 1) = CDbl(irr)
            End If
        End If

    Next i

    Set outputRange = ws.Range("Y2:Y" & lastRow)
    outputRange.Value2 = output
    outputRange.NumberFormat = "0.0%"

    completedSuccessfully = True

CleanExit:
    Application.Calculation = previousCalculation
    Application.Calculate
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If completedSuccessfully Then
        MsgBox rowCount & " residential IRRs updated.", _
               vbInformation, "Script Finished"
    End If

    Exit Sub

CleanFail:
    MsgBox "Residential IRR update failed: " & Err.Description, _
           vbCritical, "ResidentialIRR"
    Resume CleanExit

End Sub

Private Function ToNumber(ByVal value As Variant) As Variant

    Dim cleaned As String

    If IsError(value) Or IsEmpty(value) Then
        ToNumber = Empty
        Exit Function
    End If

    If IsNumeric(value) Then
        ToNumber = CDbl(value)
        Exit Function
    End If

    cleaned = CStr(value)
    cleaned = Replace(cleaned, "$", "")
    cleaned = Replace(cleaned, ",", "")
    cleaned = Replace(cleaned, " ", "")
    cleaned = Replace(cleaned, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")

    If Len(cleaned) > 0 And IsNumeric(cleaned) Then
        ToNumber = CDbl(cleaned)
    Else
        ToNumber = Empty
    End If

End Function



