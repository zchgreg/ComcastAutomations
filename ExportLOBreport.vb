Option Explicit

Public Sub Export_LOB_Report()

    Dim sourceWb As Workbook
    Dim reportWb As Workbook
    Dim savePath As Variant
    Dim oldCalculation As XlCalculation

    On Error GoTo ErrorHandler

    Set sourceWb = ThisWorkbook
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    'Create report workbook with Weekly_Callouts.
    CopyFirstSheet sourceWb.Worksheets("Weekly_Callouts"), reportWb
    ConvertSheetToValues reportWb.Worksheets(1)

    'BP Stack Rank: copy exactly as-is.
    CopySheetComplete _
        sourceWb.Worksheets("BP_Stack_Rank"), reportWb

    'Full Residential sheet:
    'A:AG becomes values; AI:AP formulas remain.
    CopyDataSheet _
        sourceWb.Worksheets("Resi_IRR_ALL_LOB"), _
        reportWb, _
        "AG"

    'Full Commercial sheet:
    'A:W becomes values; Z:AG formulas remain.
    CopyDataSheet _
        sourceWb.Worksheets("Com_ALL_LOB"), _
        reportWb, _
        "W"

    'Residential LOB sheets.
    SplitSheetByLOB _
        sourceWb.Worksheets("Resi_IRR_ALL_LOB"), _
        reportWb, _
        "AG"

    'Commercial LOB sheets.
    SplitSheetByLOB _
        sourceWb.Worksheets("Com_ALL_LOB"), _
        reportWb, _
        "W"

    reportWb.Worksheets(1).Activate
    Application.CalculateFull

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="LOB Report.xlsx", _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx")

    If savePath <> False Then

        reportWb.SaveAs _
            Filename:=CStr(savePath), _
            FileFormat:=xlOpenXMLWorkbook

        MsgBox _
            "The report was created and saved.", _
            vbInformation, "Export Complete"

    Else

        MsgBox _
            "The report was created but was not saved.", _
            vbExclamation, "Export Complete"

    End If

CleanExit:
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox _
        "Export failed: " & Err.Description, _
        vbCritical, "LOB Report"

    Resume CleanExit

End Sub

Private Sub CopyFirstSheet( _
    ByVal sourceWs As Worksheet, _
    ByRef reportWb As Workbook)

    Dim freezePanes As Boolean
    Dim frozenRows As Long
    Dim frozenColumns As Long

    GetFreezeSettings _
        sourceWs, freezePanes, frozenRows, frozenColumns

    sourceWs.Copy
    Set reportWb = ActiveWorkbook

    ApplyFreezeSettings _
        reportWb.Worksheets(1), _
        freezePanes, frozenRows, frozenColumns

End Sub

Private Sub CopySheetComplete( _
    ByVal sourceWs As Worksheet, _
    ByVal reportWb As Workbook)

    Dim newWs As Worksheet
    Dim freezePanes As Boolean
    Dim frozenRows As Long
    Dim frozenColumns As Long

    GetFreezeSettings _
        sourceWs, freezePanes, frozenRows, frozenColumns

    sourceWs.Copy After:=reportWb.Worksheets( _
        reportWb.Worksheets.Count)

    Set newWs = reportWb.Worksheets( _
        reportWb.Worksheets.Count)

    ApplyFreezeSettings _
        newWs, freezePanes, frozenRows, frozenColumns

End Sub

Private Sub CopyDataSheet( _
    ByVal sourceWs As Worksheet, _
    ByVal reportWb As Workbook, _
    ByVal detailLastColumn As String)

    Dim newWs As Worksheet
    Dim lastRow As Long

    Dim freezePanes As Boolean
    Dim frozenRows As Long
    Dim frozenColumns As Long

    GetFreezeSettings _
        sourceWs, freezePanes, frozenRows, frozenColumns

    lastRow = sourceWs.Cells( _
        sourceWs.Rows.Count, "A").End(xlUp).Row

    sourceWs.Copy After:=reportWb.Worksheets( _
        reportWb.Worksheets.Count)

    Set newWs = reportWb.Worksheets( _
        reportWb.Worksheets.Count)

    'Only convert the detail section.
    'Summary formulas and charts remain intact.
    With newWs.Range( _
        "A1:" & detailLastColumn & lastRow)

        .Value2 = .Value2

    End With

    RepointCharts sourceWs, newWs

    newWs.Calculate

    ApplyFreezeSettings _
        newWs, freezePanes, frozenRows, frozenColumns

End Sub

Private Sub SplitSheetByLOB( _
    ByVal sourceWs As Worksheet, _
    ByVal reportWb As Workbook, _
    ByVal detailLastColumn As String)

    Dim lobDictionary As Object
    Dim sourceData As Variant
    Dim outputData() As Variant

    Dim lob As Variant
    Dim lobValue As String
    Dim newWs As Worksheet

    Dim lastRow As Long
    Dim lastColumn As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim matchCount As Long
    Dim outputRow As Long

    Dim freezePanes As Boolean
    Dim frozenRows As Long
    Dim frozenColumns As Long

    Set lobDictionary = CreateObject( _
        "Scripting.Dictionary")

    lobDictionary.CompareMode = vbTextCompare

    lastRow = sourceWs.Cells( _
        sourceWs.Rows.Count, "A").End(xlUp).Row

    lastColumn = sourceWs.Range( _
        detailLastColumn & "1").Column

    If lastRow < 2 Then
        Err.Raise vbObjectError + 1000, , _
            "No data rows were found on '" & _
            sourceWs.Name & "'."
    End If

    sourceData = sourceWs.Range( _
        sourceWs.Cells(1, 1), _
        sourceWs.Cells(lastRow, lastColumn)).Value2

    GetFreezeSettings _
        sourceWs, freezePanes, frozenRows, frozenColumns

    'Collect unique LOBs from column A.
    For rowIndex = 2 To UBound(sourceData, 1)

        If Not IsError(sourceData(rowIndex, 1)) Then

            lobValue = Trim$( _
                CStr(sourceData(rowIndex, 1)))

            If Len(lobValue) > 0 Then

                If Not lobDictionary.Exists(lobValue) Then
                    lobDictionary.Add lobValue, lobValue
                End If

            End If

        End If

    Next rowIndex

    For Each lob In lobDictionary.Keys

        matchCount = CountLOBRows( _
            sourceData, CStr(lob))

        If matchCount > 0 Then

            ReDim outputData( _
                1 To matchCount, _
                1 To lastColumn)

            outputRow = 0

            For rowIndex = 2 To UBound(sourceData, 1)

                If Not IsError(sourceData(rowIndex, 1)) Then

                    If StrComp( _
                        Trim$(CStr(sourceData(rowIndex, 1))), _
                        CStr(lob), _
                        vbTextCompare) = 0 Then

                        outputRow = outputRow + 1

                        For columnIndex = 1 To lastColumn

                            outputData( _
                                outputRow, columnIndex) = _
                                sourceData( _
                                    rowIndex, columnIndex)

                        Next columnIndex

                    End If

                End If

            Next rowIndex

            'Copy complete source sheet to preserve:
            'formatting, summaries, charts and objects.
            sourceWs.Copy After:=reportWb.Worksheets( _
                reportWb.Worksheets.Count)

            Set newWs = reportWb.Worksheets( _
                reportWb.Worksheets.Count)

            newWs.Name = GetUniqueSheetName( _
                reportWb, _
                CleanSheetName(CStr(lob)))

            'Convert the entire detail area to values.
            With newWs.Range( _
                "A1:" & detailLastColumn & lastRow)

                .Value2 = .Value2

            End With

            'Remove all mixed-LOB rows.
            newWs.Range( _
                "A2:" & detailLastColumn & lastRow _
                ).ClearContents

            'Write only the selected LOB.
            newWs.Range("A2").Resize( _
                matchCount, lastColumn).Value2 = outputData

            'Charts copied from the source should point
            'to this new LOB sheet.
            RepointCharts sourceWs, newWs

            'Summary formulas remain live because they
            'were outside the converted detail area.
            newWs.Calculate

            ApplyFreezeSettings _
                newWs, _
                freezePanes, _
                frozenRows, _
                frozenColumns
                
            RemoveAllFilters newWs
            
        End If

    Next lob

End Sub

Private Function CountLOBRows( _
    ByVal sourceData As Variant, _
    ByVal requestedLOB As String) As Long

    Dim rowIndex As Long
    Dim result As Long

    For rowIndex = 2 To UBound(sourceData, 1)

        If Not IsError(sourceData(rowIndex, 1)) Then

            If StrComp( _
                Trim$(CStr(sourceData(rowIndex, 1))), _
                requestedLOB, _
                vbTextCompare) = 0 Then

                result = result + 1

            End If

        End If

    Next rowIndex

    CountLOBRows = result

End Function

Private Sub RepointCharts( _
    ByVal sourceWs As Worksheet, _
    ByVal destinationWs As Worksheet)

    Dim chartObject As chartObject
    Dim chartSeries As Series
    Dim seriesFormula As String

    Dim oldLocalReference As String
    Dim oldExternalReference As String
    Dim newReference As String

    oldLocalReference = _
        "'" & EscapeSheetName(sourceWs.Name) & "'!"

    oldExternalReference = _
        "'[" & EscapeSheetName(sourceWs.Parent.Name) & "]" & _
        EscapeSheetName(sourceWs.Name) & "'!"

    newReference = _
        "'" & EscapeSheetName(destinationWs.Name) & "'!"

    For Each chartObject In destinationWs.ChartObjects

        For Each chartSeries In _
            chartObject.Chart.SeriesCollection

            seriesFormula = chartSeries.Formula

            seriesFormula = Replace( _
                seriesFormula, _
                oldExternalReference, _
                newReference, _
                1, -1, vbTextCompare)

            seriesFormula = Replace( _
                seriesFormula, _
                oldLocalReference, _
                newReference, _
                1, -1, vbTextCompare)

            On Error Resume Next
            chartSeries.Formula = seriesFormula
            On Error GoTo 0

        Next chartSeries

    Next chartObject

End Sub

Private Function EscapeSheetName( _
    ByVal text As String) As String

    EscapeSheetName = Replace(text, "'", "''")

End Function

Private Sub ConvertSheetToValues( _
    ByVal ws As Worksheet)

    With ws.usedRange
        .Value2 = .Value2
    End With

End Sub

Private Sub GetFreezeSettings( _
    ByVal ws As Worksheet, _
    ByRef freezePanes As Boolean, _
    ByRef frozenRows As Long, _
    ByRef frozenColumns As Long)

    ws.Parent.Activate
    ws.Activate

    With ActiveWindow
        freezePanes = .freezePanes
        frozenRows = .SplitRow
        frozenColumns = .SplitColumn
    End With

End Sub

Private Sub ApplyFreezeSettings( _
    ByVal ws As Worksheet, _
    ByVal freezePanes As Boolean, _
    ByVal frozenRows As Long, _
    ByVal frozenColumns As Long)

    ws.Parent.Activate
    ws.Activate

    With ActiveWindow

        .freezePanes = False
        .SplitRow = 0
        .SplitColumn = 0

        If freezePanes Then
            .SplitRow = frozenRows
            .SplitColumn = frozenColumns
            .freezePanes = True
        End If

    End With

End Sub

Private Function CleanSheetName( _
    ByVal requestedName As String) As String

    Dim invalidCharacters As Variant
    Dim character As Variant

    invalidCharacters = Array( _
        "\", "/", "*", "[", "]", ":", "?")

    For Each character In invalidCharacters

        requestedName = Replace( _
            requestedName, _
            CStr(character), _
            "-")

    Next character

    requestedName = Trim$(requestedName)

    Do While Len(requestedName) > 0 And _
             Left$(requestedName, 1) = "'"

        requestedName = Mid$(requestedName, 2)

    Loop

    Do While Len(requestedName) > 0 And _
             Right$(requestedName, 1) = "'"

        requestedName = Left$( _
            requestedName, _
            Len(requestedName) - 1)

    Loop

    If Len(requestedName) = 0 Then
        requestedName = "LOB"
    End If

    If Len(requestedName) > 31 Then
        requestedName = Left$(requestedName, 31)
    End If

    CleanSheetName = requestedName

End Function

Private Function GetUniqueSheetName( _
    ByVal wb As Workbook, _
    ByVal requestedName As String) As String

    Dim baseName As String
    Dim candidateName As String
    Dim suffix As Long
    Dim maximumBaseLength As Long

    baseName = requestedName
    candidateName = baseName
    suffix = 1

    Do While WorksheetExists(wb, candidateName)

        suffix = suffix + 1

        maximumBaseLength = _
            31 - Len(CStr(suffix)) - 1

        candidateName = _
            Left$(baseName, maximumBaseLength) & _
            "_" & suffix

    Loop

    GetUniqueSheetName = candidateName

End Function

Private Function WorksheetExists( _
    ByVal wb As Workbook, _
    ByVal worksheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(worksheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0

End Function

Private Sub RemoveAllFilters(ByVal ws As Worksheet)

    Dim lo As ListObject

    On Error Resume Next

    'Clear active filtering and remove worksheet filter arrows.
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False

    'Remove filter arrows from any retained Excel tables.
    For Each lo In ws.ListObjects
        lo.AutoFilter.ShowAllData
        lo.ShowAutoFilter = False
    Next lo

    On Error GoTo 0

End Sub


