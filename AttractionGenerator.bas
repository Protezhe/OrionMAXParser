Attribute VB_Name = "AttractionGenerator"
Option Explicit

' Indices for sheets
Private Const SHEET_SETTINGS As Long = 2
Private Const SHEET_INPUT As Long = 3
Private Const SHEET_ATTRACTIONS As Long = 4
Private Const SHEET_SCHEDULE As Long = 5
Private Const SHEET_STATS As Long = 6

Public Sub GenerateReport()
    On Error GoTo Fail
    Dim yearNumber As Long, monthNumber As Long, wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    yearNumber = CLng(wsSettings.Range("B2").value)
    monthNumber = GetMonthNumber(CStr(wsSettings.Range("B3").value))
    If monthNumber < 1 Or monthNumber > 12 Then
        MsgBox ChrW(1052) & ChrW(1077) & ChrW(1089) & ChrW(1103) & ChrW(1094) & " " & ChrW(1085) & ChrW(1077) & " " & ChrW(1088) & ChrW(1072) & ChrW(1089) & ChrW(1087) & ChrW(1086) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1085) & ".", vbExclamation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    GenerateMonthSheet yearNumber, monthNumber
    UpdateStatistics yearNumber
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Sub GenerateMonthSheet(ByVal yearNumber As Long, ByVal monthNumber As Long)
    Dim ws As Worksheet, sheetName As String, attrNames() As String, attrDisplay() As String, attrCount As Long
    Dim firstDay As Date, lastDay As Date, d As Date, r As Long, i As Long, baseCol As Long, summaryCol As Long
    Dim events() As Collection, maxEvents As Long, eventRow As Long, endRow As Long, totalRow As Long, weekRow As Long
    Dim dayDowntime() As Double, dayCount() As Long, weekDowntime() As Double, weekCount() As Long
    Dim totalDowntime() As Double, totalCount() As Long, startTime As Date, endTime As Date, ev As Variant
    LoadAttractions attrNames, attrDisplay, attrCount
    If attrCount = 0 Then Exit Sub
    sheetName = MonthNameRu(monthNumber) & CStr(yearNumber)
    If SheetExists(sheetName) Then
        If MsgBox(ChrW(1051) & ChrW(1080) & ChrW(1089) & ChrW(1090) & " " & sheetName & " " & ChrW(1091) & ChrW(1081) & ChrW(1077) & " " & ChrW(1089) & ChrW(1091) & ChrW(1099) & ChrW(1077) & ChrW(1089) & ChrW(1090) & ChrW(1074) & ChrW(1091) & ChrW(1077) & ChrW(1090) & ". " & ChrW(1055) & ChrW(1077) & ChrW(1088) & ChrW(1077) & ChrW(1089) & ChrW(1086) & ChrW(1079) & ChrW(1076) & ChrW(1072) & ChrW(1090) & ChrW(1100) & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
        Application.DisplayAlerts = False: ThisWorkbook.Worksheets(sheetName).Delete: Application.DisplayAlerts = True
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName: ActiveWindow.DisplayGridlines = False
    summaryCol = 2 + attrCount * 4: BuildMonthHeader ws, attrDisplay, attrCount, summaryCol
    ReDim weekDowntime(1 To attrCount): ReDim weekCount(1 To attrCount)
    ReDim totalDowntime(1 To attrCount): ReDim totalCount(1 To attrCount)
    firstDay = DateSerial(yearNumber, monthNumber, 1): lastDay = DateSerial(yearNumber, monthNumber + 1, 0): r = 3
    For d = firstDay To lastDay
        ReDim events(1 To attrCount): ReDim dayDowntime(1 To attrCount): ReDim dayCount(1 To attrCount): maxEvents = 1
        For i = 1 To attrCount
            Set events(i) = GetEventsForDay(d, attrNames(i))
            If events(i).Count > maxEvents Then maxEvents = events(i).Count
        Next i
        GetScheduleForDate d, startTime, endTime
        ws.Cells(r, 1).value = d
        For i = 1 To attrCount: ws.Cells(r, 2 + (i - 1) * 4 + 1).value = startTime: Next i
        FormatDayBoundaryRow ws, r, attrCount, summaryCol
        For eventRow = 1 To maxEvents
            ws.Cells(r + eventRow, 1).value = d
            For i = 1 To attrCount
                baseCol = 2 + (i - 1) * 4
                If eventRow <= events(i).Count Then
                    ev = events(i)(eventRow)
                    ws.Cells(r + eventRow, baseCol).value = ev(0): ws.Cells(r + eventRow, baseCol + 1).value = ev(1)
                    ws.Cells(r + eventRow, baseCol + 2).value = ev(2): ws.Cells(r + eventRow, baseCol + 3).value = ev(3)
                    If IsNumeric(ev(2)) Then
                        dayDowntime(i) = dayDowntime(i) + CDbl(ev(2))
                        If CDbl(ev(2)) > 0 Then dayCount(i) = dayCount(i) + 1
                    End If
                End If
            Next i
            FormatDataRow ws, r + eventRow, summaryCol
        Next eventRow
        endRow = r + maxEvents + 1: ws.Cells(endRow, 1).value = d
        For i = 1 To attrCount: ws.Cells(endRow, 2 + (i - 1) * 4).value = endTime: Next i
        FormatDayBoundaryRow ws, endRow, attrCount, summaryCol
        totalRow = endRow + 1: ws.Cells(totalRow, 1).value = d
        For i = 1 To attrCount
            baseCol = 2 + (i - 1) * 4: ws.Cells(totalRow, baseCol).value = TimeSpan(startTime, endTime)
            ws.Cells(totalRow, baseCol + 2).value = dayDowntime(i): ws.Cells(totalRow, baseCol + 3).value = dayCount(i)
            weekDowntime(i) = weekDowntime(i) + dayDowntime(i): weekCount(i) = weekCount(i) + dayCount(i)
            totalDowntime(i) = totalDowntime(i) + dayDowntime(i): totalCount(i) = totalCount(i) + dayCount(i)
        Next i
        ws.Cells(totalRow, summaryCol).value = SumDoubleArray(dayDowntime): ws.Cells(totalRow, summaryCol + 1).value = SumLongArray(dayCount)
        FormatTotalRow ws, totalRow, summaryCol: r = totalRow + 1
        If Weekday(d, vbMonday) = 7 Or d = lastDay Then
            weekRow = r: ws.Cells(weekRow, 1).value = CStr(Application.WorksheetFunction.IsoWeekNum(d)) & " " & TextWeek()
            For i = 1 To attrCount
                baseCol = 2 + (i - 1) * 4: ws.Cells(weekRow, baseCol + 2).value = weekDowntime(i): ws.Cells(weekRow, baseCol + 3).value = weekCount(i)
                weekDowntime(i) = 0: weekCount(i) = 0
            Next i
            ws.Cells(weekRow, summaryCol).value = SumRowAttractionValues(ws, weekRow, attrCount, 2)
            ws.Cells(weekRow, summaryCol + 1).value = SumRowAttractionValues(ws, weekRow, attrCount, 3)
            FormatWeekRow ws, weekRow, summaryCol: r = r + 1
        End If
    Next d
    ws.Cells(r, 1).value = TextMonthTotal(monthNumber)
    For i = 1 To attrCount
        baseCol = 2 + (i - 1) * 4: ws.Cells(r, baseCol + 2).value = totalDowntime(i): ws.Cells(r, baseCol + 3).value = totalCount(i)
    Next i
    ws.Cells(r, summaryCol).value = SumDoubleArray(totalDowntime): ws.Cells(r, summaryCol + 1).value = SumLongArray(totalCount)
    FormatGrandTotalRow ws, r, summaryCol: r = r + 1
    ws.Cells(r, 1).value = TextNormativeWorkTime(): ws.Cells(r, 2).value = NormativeTimeForMonth(yearNumber, monthNumber)
    ws.Cells(r, summaryCol - 1).value = TextDowntimeRatio()
    If CDbl(ws.Cells(r, 2).value) > 0 Then ws.Cells(r, summaryCol).value = CDbl(ws.Cells(r - 1, summaryCol).value) / CDbl(ws.Cells(r, 2).value)
    ApplyMonthFormats ws, r, attrCount, summaryCol: FormatNormativeRow ws, r, summaryCol
End Sub

Private Sub UpdateStatistics(ByVal yearNumber As Long)
    Dim ws As Worksheet, inputWs As Worksheet, attrNames() As String, attrDisplay() As String, attrCount As Long
    Dim monthTotals() As Double, monthCounts() As Long, monthNo As Long, i As Long, r As Long, baseCol As Long, summaryCol As Long
    Dim lastRow As Long, attrIndex As Long, duration As Double, totalDowntime As Double, totalCount As Long, totalNoAviator As Double, countNoAviator As Long, normative As Double
    LoadAttractions attrNames, attrDisplay, attrCount: If attrCount = 0 Then Exit Sub
    ReDim monthTotals(1 To 12, 1 To attrCount): ReDim monthCounts(1 To 12, 1 To attrCount)
    Set inputWs = ThisWorkbook.Worksheets(SHEET_INPUT): lastRow = inputWs.Cells(inputWs.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CanReadDate(inputWs.Cells(r, 1).value) Then
            If Year(ReadDateValue(inputWs.Cells(r, 1).value)) = yearNumber Then
                attrIndex = AttractionIndex(CStr(inputWs.Cells(r, 2).value), attrNames, attrCount)
                If attrIndex > 0 Then
                    If CanReadTime(inputWs.Cells(r, 3).value) And CanReadTime(inputWs.Cells(r, 4).value) Then
                        duration = TimeSpan(ReadTimeValue(inputWs.Cells(r, 3).value), ReadTimeValue(inputWs.Cells(r, 4).value))
                        monthNo = Month(ReadDateValue(inputWs.Cells(r, 1).value))
                        monthTotals(monthNo, attrIndex) = monthTotals(monthNo, attrIndex) + duration: monthCounts(monthNo, attrIndex) = monthCounts(monthNo, attrIndex) + 1
                    End If
                End If
            End If
        End If
    Next r
    Set ws = ThisWorkbook.Worksheets(SHEET_STATS): ws.Cells.Clear: ActiveWindow.DisplayGridlines = False: summaryCol = 2 + attrCount * 2
    ws.Cells(1, 1).value = TextMonth()
    For i = 1 To attrCount
        baseCol = 2 + (i - 1) * 2: ws.Range(ws.Cells(1, baseCol), ws.Cells(1, baseCol + 1)).Merge: ws.Cells(1, baseCol).value = attrDisplay(i)
        ws.Cells(2, baseCol).value = TextDowntime(): ws.Cells(2, baseCol + 1).value = TextStopCount()
    Next i
    ws.Cells(1, summaryCol).value = TextTotalDowntime(): ws.Cells(1, summaryCol + 1).value = TextTotalStops()
    For monthNo = 1 To 12
        r = monthNo + 2: ws.Cells(r, 1).value = MonthNameRu(monthNo)
        For i = 1 To attrCount
            baseCol = 2 + (i - 1) * 2: ws.Cells(r, baseCol).value = monthTotals(monthNo, i): ws.Cells(r, baseCol + 1).value = monthCounts(monthNo, i)
        Next i
        ws.Cells(r, summaryCol).value = SumRowStats(ws, r, attrCount, 0): ws.Cells(r, summaryCol + 1).value = SumRowStats(ws, r, attrCount, 1)
    Next monthNo
    r = 15: ws.Cells(r, 1).value = TextTotal()
    For i = 1 To attrCount
        baseCol = 2 + (i - 1) * 2: ws.Cells(r, baseCol).Formula = "=SUM(" & ws.Cells(3, baseCol).Address(0, 0) & ":" & ws.Cells(14, baseCol).Address(0, 0) & ")": ws.Cells(r, baseCol + 1).Formula = "=SUM(" & ws.Cells(3, baseCol + 1).Address(0, 0) & ":" & ws.Cells(14, baseCol + 1).Address(0, 0) & ")"
    Next i
    ws.Cells(r, summaryCol).Formula = "=SUM(" & ws.Cells(3, summaryCol).Address(0, 0) & ":" & ws.Cells(14, summaryCol).Address(0, 0) & ")": ws.Cells(r, summaryCol + 1).Formula = "=SUM(" & ws.Cells(3, summaryCol + 1).Address(0, 0) & ":" & ws.Cells(14, summaryCol + 1).Address(0, 0) & ")"
    totalDowntime = 0: totalCount = 0: totalNoAviator = 0: countNoAviator = 0
    For monthNo = 1 To 12
        For i = 1 To attrCount
            totalDowntime = totalDowntime + monthTotals(monthNo, i): totalCount = totalCount + monthCounts(monthNo, i)
            If LCase(attrNames(i)) <> LCase(AviatorRu()) Then
                totalNoAviator = totalNoAviator + monthTotals(monthNo, i): countNoAviator = countNoAviator + monthCounts(monthNo, i)
            End If
        Next i
    Next monthNo
    normative = NormativeTimeForYear(yearNumber) * attrCount: r = 18: ws.Cells(r, 1).value = "Summary": ws.Cells(r + 1, 1).value = "All downtime": ws.Cells(r + 1, 2).value = totalDowntime: ws.Cells(r + 2, 1).value = "All stops": ws.Cells(r + 2, 2).value = totalCount: ws.Cells(r + 3, 1).value = "Without Aviator DT": ws.Cells(r + 3, 2).value = totalNoAviator: ws.Cells(r + 4, 1).value = "Without Aviator Stops": ws.Cells(r + 4, 2).value = countNoAviator: ws.Cells(r + 5, 1).value = "Normative time": ws.Cells(r + 5, 2).value = normative: ws.Cells(r + 6, 1).value = "Ratio": If normative > 0 Then ws.Cells(r + 6, 2).value = totalDowntime / normative
    FormatStatsSheet ws, attrCount, summaryCol: CreateStatisticsCharts ws, attrCount, attrDisplay
End Sub

Private Sub CreateStatisticsCharts(ByVal ws As Worksheet, ByVal attrCount As Long, ByRef attrDisplay() As String)
    On Error Resume Next
    Dim chObjDowntime As ChartObject, chObjStops As ChartObject, i As Long, baseCol As Long, summaryCol As Long, cats() As String, vals() As Variant
    summaryCol = 2 + attrCount * 2 + 2: Application.ScreenUpdating = True: ws.Calculate
    Dim ch As ChartObject
    For Each ch In ws.ChartObjects: ch.Delete: Next ch
    Set chObjDowntime = ws.ChartObjects.Add(Left:=ws.Cells(1, summaryCol).Left, Top:=ws.Cells(2, 1).Top, Width:=600, Height:=250)
    With chObjDowntime.Chart
        .ChartType = xlColumnStacked: .HasTitle = True: .ChartTitle.text = TextDowntime()
        For i = 1 To attrCount
            baseCol = 2 + (i - 1) * 2
            With .SeriesCollection.NewSeries: .Name = attrDisplay(i): .XValues = ws.Range(ws.Cells(3, 1), ws.Cells(14, 1)): .values = ws.Range(ws.Cells(3, baseCol), ws.Cells(14, baseCol)): End With
        Next i
    End With
    Set chObjStops = ws.ChartObjects.Add(Left:=chObjDowntime.Left, Top:=chObjDowntime.Top + chObjDowntime.Height + 20, Width:=600, Height:=250)
    With chObjStops.Chart
        .ChartType = xlColumnClustered: .HasTitle = True: .ChartTitle.text = TextStopCount()
        ReDim cats(1 To attrCount): ReDim vals(1 To attrCount)
        For i = 1 To attrCount: cats(i) = attrDisplay(i): vals(i) = ws.Cells(15, 2 + (i - 1) * 2 + 1).value: Next i
        With .SeriesCollection.NewSeries: .XValues = cats: .values = vals: End With
        .HasLegend = False
    End With
    Application.ScreenUpdating = False
End Sub

Private Sub LoadAttractions(ByRef names() As String, ByRef displayNames() As String, ByRef attrCount As Long)
    Dim ws As Worksheet, lastRow As Long, r As Long, includeValue As String
    Set ws = ThisWorkbook.Worksheets(SHEET_ATTRACTIONS): lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row: attrCount = 0
    For r = 2 To lastRow
        If Len(Trim(CStr(ws.Cells(r, 1).value))) > 0 Then
            includeValue = LCase(Trim(CStr(ws.Cells(r, 3).value)))
            If includeValue = "" Or includeValue = "1" Or includeValue = "yes" Or includeValue = "true" Or includeValue = YesRu() Then
                attrCount = attrCount + 1: ReDim Preserve names(1 To attrCount): ReDim Preserve displayNames(1 To attrCount)
                names(attrCount) = CStr(ws.Cells(r, 1).value): displayNames(attrCount) = IIf(Len(Trim(CStr(ws.Cells(r, 2).value))) > 0, CStr(ws.Cells(r, 2).value), names(attrCount))
            End If
        End If
    Next r
End Sub

Private Function GetEventsForDay(ByVal targetDate As Date, ByVal attractionName As String) As Collection
    Dim ws As Worksheet, lastRow As Long, r As Long, d As Variant, ev(0 To 3) As Variant
    Set ws = ThisWorkbook.Worksheets(SHEET_INPUT): Set GetEventsForDay = New Collection: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        d = ws.Cells(r, 1).value
        If CanReadDate(d) Then
            If ReadDateValue(d) = targetDate And Trim(CStr(ws.Cells(r, 2).value)) = attractionName Then
                If CanReadTime(ws.Cells(r, 3).value) Or CanReadTime(ws.Cells(r, 4).value) Then
                    ev(0) = IIf(CanReadTime(ws.Cells(r, 3).value), ReadTimeValue(ws.Cells(r, 3).value), Empty)
                    ev(1) = IIf(CanReadTime(ws.Cells(r, 4).value), ReadTimeValue(ws.Cells(r, 4).value), Empty)
                    ev(2) = IIf(CanReadTime(ws.Cells(r, 3).value) And CanReadTime(ws.Cells(r, 4).value), TimeSpan(ev(0), ev(1)), Empty)
                    ev(3) = CStr(ws.Cells(r, 6).value): AddSortedEvent GetEventsForDay, ev
                End If
            End If
        End If
    Next r
End Function

Private Sub AddSortedEvent(ByRef events As Collection, ByVal ev As Variant)
    Dim i As Long: If events.Count = 0 Then events.Add ev: Exit Sub
    For i = 1 To events.Count: If IIf(IsEmpty(ev(0)), 2, CDbl(ev(0))) < IIf(IsEmpty(events(i)(0)), 2, CDbl(events(i)(0))) Then events.Add ev, Before:=i: Exit Sub
    Next i
    events.Add ev
End Sub

Private Sub GetScheduleForDate(ByVal targetDate As Date, ByRef startTime As Date, ByRef endTime As Date)
    Dim ws As Worksheet, wsS As Worksheet, lastRow As Long, r As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_SCHEDULE): Set wsS = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    startTime = TimeValue(wsS.Range("B4").value): endTime = TimeValue(wsS.Range("B5").value): lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CanReadDate(ws.Cells(r, 1).value) Then
            If ReadDateValue(ws.Cells(r, 1).value) = targetDate Then
                If CanReadTime(ws.Cells(r, 3).value) Then startTime = ReadTimeValue(ws.Cells(r, 3).value)
                If CanReadTime(ws.Cells(r, 4).value) Then endTime = ReadTimeValue(ws.Cells(r, 4).value)
                Exit Sub
            End If
        End If
    Next r
End Sub

Private Function TimeSpan(ByVal s As Date, ByVal e As Date) As Double
    TimeSpan = IIf(CDbl(e) < CDbl(s), CDbl(e) + 1 - CDbl(s), CDbl(e) - CDbl(s))
End Function

Private Function CanReadDate(ByVal v As Variant) As Boolean
    On Error Resume Next: CanReadDate = (IsDate(v) Or IsNumeric(v)) And Len(CStr(v)) > 0
End Function

Private Function ReadDateValue(ByVal v As Variant) As Date: ReadDateValue = DateValue(CDate(v)): End Function
Private Function CanReadTime(ByVal v As Variant) As Boolean: On Error Resume Next: CanReadTime = (IsDate(v) Or IsNumeric(v)) And Len(CStr(v)) > 0: End Function
Private Function ReadTimeValue(ByVal v As Variant) As Date: Dim n As Double: n = CDbl(CDate(v)): ReadTimeValue = CDate(n - Fix(n)): End Function
Private Function SumDoubleArray(ByRef v() As Double) As Double: Dim i As Long: For i = LBound(v) To UBound(v): SumDoubleArray = SumDoubleArray + v(i): Next i: End Function
Private Function SumLongArray(ByRef v() As Long) As Long: Dim i As Long: For i = LBound(v) To UBound(v): SumLongArray = SumLongArray + v(i): Next i: End Function
Private Function SumRowAttractionValues(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal o As Long) As Double: Dim i As Long: For i = 1 To c: If IsNumeric(ws.Cells(r, 2 + (i - 1) * 4 + o).value) Then SumRowAttractionValues = SumRowAttractionValues + CDbl(ws.Cells(r, 2 + (i - 1) * 4 + o).value)
    Next i: End Function
Private Function SumRowStats(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal o As Long) As Double: Dim i As Long: For i = 1 To c: If IsNumeric(ws.Cells(r, 2 + (i - 1) * 2 + o).value) Then SumRowStats = SumRowStats + CDbl(ws.Cells(r, 2 + (i - 1) * 2 + o).value)
    Next i: End Function
Private Function NormativeTimeForYear(ByVal y As Long) As Double: Dim ws As Worksheet, r As Long, d As Variant: Set ws = ThisWorkbook.Worksheets(SHEET_SCHEDULE): For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row: d = ws.Cells(r, 1).value: If CanReadDate(d) Then: If Year(ReadDateValue(d)) = y Then NormativeTimeForYear = NormativeTimeForYear + TimeSpan(ws.Cells(r, 3).value, ws.Cells(r, 4).value)
    End If: Next r: End Function
Private Function NormativeTimeForMonth(ByVal y As Long, ByVal m As Long) As Double: Dim ws As Worksheet, r As Long, d As Variant: Set ws = ThisWorkbook.Worksheets(SHEET_SCHEDULE): For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row: d = ws.Cells(r, 1).value: If CanReadDate(d) Then: If Year(ReadDateValue(d)) = y And Month(ReadDateValue(d)) = m Then NormativeTimeForMonth = NormativeTimeForMonth + TimeSpan(ws.Cells(r, 3).value, ws.Cells(r, 4).value)
    End If: Next r: End Function
Private Function AttractionIndex(ByVal n As String, ByRef names() As String, ByVal c As Long) As Long: Dim i As Long: For i = 1 To c: If LCase(Trim(n)) = LCase(names(i)) Then AttractionIndex = i: Exit Function
    Next i: End Function

Private Sub BuildMonthHeader(ByVal ws As Worksheet, ByRef attr() As String, ByVal c As Long, ByVal sc As Long)
    Dim i As Long, bc As Long: ws.Cells(1, 1).value = TextDate(): ws.Range(ws.Cells(1, 1), ws.Cells(2, 1)).Merge
    For i = 1 To c: bc = 2 + (i - 1) * 4: ws.Range(ws.Cells(1, bc), ws.Cells(1, bc + 3)).Merge: ws.Cells(1, bc).value = attr(i): ws.Cells(2, bc).value = TextStopTime(): ws.Cells(2, bc + 1).value = TextStartTime(): ws.Cells(2, bc + 2).value = TextDowntime(): ws.Cells(2, bc + 3).value = TextReason(): Next i
    ws.Cells(1, sc).value = TextTotalDowntime(): ws.Cells(1, sc + 1).value = TextTotalStops(): ws.Range(ws.Cells(1, sc), ws.Cells(2, sc)).Merge: ws.Range(ws.Cells(1, sc + 1), ws.Cells(2, sc + 1)).Merge
    With ws.Range(ws.Cells(1, 1), ws.Cells(2, sc + 1)): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True: .Borders.LineStyle = xlContinuous: End With
End Sub

Private Sub FormatDataRow(ByVal ws As Worksheet, ByVal r As Long, ByVal sc As Long): With ws.Range(ws.Cells(r, 1), ws.Cells(r, sc + 1)): .Borders.LineStyle = xlContinuous: .VerticalAlignment = xlCenter: .WrapText = True: End With: End Sub
Private Sub FormatDayBoundaryRow(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal sc As Long): Dim i As Long: FormatDataRow ws, r, sc: For i = 1 To c: ws.Range(ws.Cells(r, 2 + (i - 1) * 4), ws.Cells(r, 2 + (i - 1) * 4 + 2)).Interior.Color = RGB(255, 192, 0): Next i: End Sub
Private Sub FormatTotalRow(ByVal ws As Worksheet, ByVal r As Long, ByVal sc As Long): FormatDataRow ws, r, sc: ws.Range(ws.Cells(r, 1), ws.Cells(r, sc + 1)).Interior.Color = RGB(255, 255, 0): FormatCountCells ws, r, sc: End Sub
Private Sub FormatWeekRow(ByVal ws As Worksheet, ByVal r As Long, ByVal sc As Long): FormatDataRow ws, r, sc: ws.Range(ws.Cells(r, 1), ws.Cells(r, sc + 1)).Interior.Color = RGB(112, 173, 71): FormatCountCells ws, r, sc: End Sub
Private Sub FormatGrandTotalRow(ByVal ws As Worksheet, ByVal r As Long, ByVal sc As Long): FormatDataRow ws, r, sc: ws.Range(ws.Cells(r, 1), ws.Cells(r, sc + 1)).Interior.Color = RGB(0, 176, 240): FormatCountCells ws, r, sc: End Sub
Private Sub FormatNormativeRow(ByVal ws As Worksheet, ByVal r As Long, ByVal sc As Long): FormatDataRow ws, r, sc: ws.Cells(r, 2).NumberFormat = "[h]:mm:ss": ws.Cells(r, sc).NumberFormat = "0.00": End Sub
Private Sub FormatCountCells(ByVal ws As Worksheet, ByVal r As Long, ByVal sc As Long): Dim i As Long: For i = 5 To sc - 1 Step 4: ws.Cells(r, i).NumberFormat = "0": ws.Cells(r, i).HorizontalAlignment = xlCenter: Next i: ws.Cells(r, sc + 1).NumberFormat = "0": ws.Cells(r, sc + 1).HorizontalAlignment = xlCenter: End Sub

Private Sub ApplyMonthFormats(ByVal ws As Worksheet, ByVal lr As Long, ByVal c As Long, ByVal sc As Long)
    Dim i As Long, bc As Long: ws.Columns(1).ColumnWidth = 11: ws.Columns(1).NumberFormat = "dd.mm.yy"
    For i = 1 To c: bc = 2 + (i - 1) * 4: ws.Columns(bc).ColumnWidth = 9: ws.Columns(bc + 1).ColumnWidth = 12: ws.Columns(bc + 2).ColumnWidth = 12: ws.Columns(bc + 3).ColumnWidth = 15: ws.Columns(bc).NumberFormat = "h:mm": ws.Columns(bc + 1).NumberFormat = "h:mm": ws.Columns(bc + 2).NumberFormat = "h:mm": ws.Columns(bc + 3).HorizontalAlignment = xlCenter: Next i
    ws.Columns(sc).ColumnWidth = 12: ws.Columns(sc + 1).ColumnWidth = 12: ws.Columns(sc).NumberFormat = "[h]:mm:ss": ws.Columns(sc + 1).HorizontalAlignment = xlCenter: ws.Range(ws.Cells(1, 1), ws.Cells(lr, sc + 1)).Font.Size = 10: ws.Activate: ws.Range("B3").Select: ActiveWindow.FreezePanes = True
End Sub

Private Sub FormatStatsSheet(ByVal ws As Worksheet, ByVal c As Long, ByVal sc As Long)
    Dim i As Long, bc As Long: With ws.Range(ws.Cells(1, 1), ws.Cells(2, sc + 1)): .Font.Bold = True: .Interior.Color = RGB(217, 234, 247): .HorizontalAlignment = xlCenter: .Borders.LineStyle = xlContinuous: End With: ws.Range(ws.Cells(15, 1), ws.Cells(15, sc + 1)).Font.Bold = True: ws.Range(ws.Cells(15, 1), ws.Cells(15, sc + 1)).Interior.Color = RGB(0, 176, 240): ws.Range(ws.Cells(1, 1), ws.Cells(15, sc + 1)).Borders.LineStyle = xlContinuous
    For i = 1 To c: bc = 2 + (i - 1) * 2: ws.Range(ws.Cells(3, bc), ws.Cells(15, bc)).NumberFormat = "[h]:mm:ss": ws.Range(ws.Cells(3, bc + 1), ws.Cells(15, bc + 1)).NumberFormat = "0": Next i: ws.Range(ws.Cells(3, sc), ws.Cells(15, sc)).NumberFormat = "[h]:mm:ss": ws.Range("B19,B21,B23").NumberFormat = "[h]:mm": ws.Range("B24").NumberFormat = "0.00%": ws.Columns.AutoFit
End Sub

Private Function YesRu() As String: YesRu = LCase(ChrW(1044) & ChrW(1072)): End Function
Private Function AviatorRu() As String: AviatorRu = ChrW(1040) & ChrW(1074) & ChrW(1080) & ChrW(1072) & ChrW(1090) & ChrW(1086) & ChrW(1088): End Function
Private Function TextDate() As String: TextDate = ChrW(1044) & ChrW(1072) & ChrW(1090) & ChrW(1072): End Function
Private Function TextMonth() As String: TextMonth = ChrW(1052) & ChrW(1077) & ChrW(1089) & ChrW(1103) & ChrW(1094): End Function
Private Function TextStopTime() As String: TextStopTime = ChrW(1042) & ChrW(1088) & ChrW(1077) & ChrW(1084) & ChrW(1103) & " " & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1072) & ChrW(1085) & ChrW(1086) & ChrW(1074): End Function
Private Function TextStartTime() As String: TextStartTime = ChrW(1042) & ChrW(1088) & ChrW(1077) & ChrW(1084) & ChrW(1103) & " " & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1072) & ChrW(1088) & ChrW(1090) & ChrW(1072): End Function
Private Function TextDowntime() As String: TextDowntime = ChrW(1042) & ChrW(1088) & ChrW(1077) & ChrW(1084) & ChrW(1103) & " " & ChrW(1087) & ChrW(1088) & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1086) & ChrW(1103): End Function
Private Function TextReason() As String: TextReason = ChrW(1055) & ChrW(1088) & ChrW(1080) & ChrW(1095) & ChrW(1080) & ChrW(1085) & ChrW(1072): End Function
Private Function TextTotalDowntime() As String: TextTotalDowntime = ChrW(1042) & ChrW(1089) & ChrW(1077) & ChrW(1075) & ChrW(1086) & " " & TextDowntime(): End Function
Private Function TextTotalStops() As String: TextTotalStops = ChrW(1042) & ChrW(1089) & ChrW(1077) & ChrW(1075) & ChrW(1086) & " " & TextStopCount(): End Function
Private Function TextStopCount() As String: TextStopCount = ChrW(1050) & ChrW(1086) & ChrW(1083) & ChrW(1080) & ChrW(1095) & ChrW(1077) & ChrW(1089) & ChrW(1090) & ChrW(1074) & ChrW(1086): End Function
Private Function TextWeek() As String: TextWeek = ChrW(1085) & ChrW(1077) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1103): End Function
Private Function TextMonthTotal(ByVal m As Long) As String: TextMonthTotal = ChrW(1047) & ChrW(1072) & " " & LCase(MonthNameRu(m)): End Function
Private Function TextTotal() As String: TextTotal = ChrW(1042) & ChrW(1089) & ChrW(1077) & ChrW(1075) & ChrW(1086) & ":": End Function
Private Function TextNormativeWorkTime() As String: TextNormativeWorkTime = ChrW(1053) & ChrW(1086) & ChrW(1088) & ChrW(1084) & ChrW(1072) & ChrW(1090) & ChrW(1080) & ChrW(1074): End Function
Private Function TextDowntimeRatio() As String: TextDowntimeRatio = ChrW(1050) & ChrW(1086) & ChrW(1101) & ChrW(1092) & ChrW(1092) & ".": End Function

Private Function GetMonthNumber(ByVal v As String) As Long: Dim i As Long: v = LCase(Trim(v)): If IsNumeric(v) Then GetMonthNumber = CLng(v): Exit Function
    For i = 1 To 12: If v = LCase(MonthNameRu(i)) Then GetMonthNumber = i: Exit Function
    Next i: End Function

Private Function MonthNameRu(ByVal m As Long) As String
    Select Case m
        Case 1: MonthNameRu = ChrW(1071) & ChrW(1085) & ChrW(1074) & ChrW(1072) & ChrW(1088) & ChrW(1100)
        Case 2: MonthNameRu = ChrW(1060) & ChrW(1077) & ChrW(1074) & ChrW(1088) & ChrW(1072) & ChrW(1083) & ChrW(1100)
        Case 3: MonthNameRu = ChrW(1052) & ChrW(1072) & ChrW(1088) & ChrW(1090)
        Case 4: MonthNameRu = ChrW(1040) & ChrW(1087) & ChrW(1088) & ChrW(1077) & ChrW(1083) & ChrW(1100)
        Case 5: MonthNameRu = ChrW(1052) & ChrW(1072) & ChrW(1081)
        Case 6: MonthNameRu = ChrW(1048) & ChrW(1102) & ChrW(1085) & ChrW(1100)
        Case 7: MonthNameRu = ChrW(1048) & ChrW(1102) & ChrW(1083) & ChrW(1100)
        Case 8: MonthNameRu = ChrW(1040) & ChrW(1074) & ChrW(1075) & ChrW(1091) & ChrW(1089) & ChrW(1090)
        Case 9: MonthNameRu = ChrW(1057) & ChrW(1077) & ChrW(1085) & ChrW(1090) & ChrW(1103) & ChrW(1073) & ChrW(1088) & ChrW(1100)
        Case 10: MonthNameRu = ChrW(1054) & ChrW(1082) & ChrW(1090) & ChrW(1103) & ChrW(1073) & ChrW(1088) & ChrW(1100)
        Case 11: MonthNameRu = ChrW(1053) & ChrW(1086) & ChrW(1103) & ChrW(1073) & ChrW(1088) & ChrW(1100)
        Case 12: MonthNameRu = ChrW(1044) & ChrW(1077) & ChrW(1082) & ChrW(1072) & ChrW(1073) & ChrW(1088) & ChrW(1100)
    End Select
End Function

Private Function SheetExists(ByVal n As String) As Boolean: Dim ws As Worksheet: For Each ws In ThisWorkbook.Worksheets: If ws.Name = n Then SheetExists = True: Exit Function
    Next ws: End Function
