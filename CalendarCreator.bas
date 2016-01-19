Attribute VB_Name = "CalendarCreator"
Sub ColorClear(Table)
    With Table.Rows.Shading
        .BackgroundPatternColor = wdColorAutomatic
        .ForegroundPatternColor = wdColorAutomatic
    End With
End Sub


Sub ColorWeekend(Item, WeekdayID, ColorSun, ColorSat)
    With Item.Shading
        If WeekdayID = vbSunday Then
            .BackgroundPatternColor = ColorSun
        ElseIf WeekdayID = vbSaturday Then
            .BackgroundPatternColor = ColorSat
        End If
    End With
End Sub


Function FindYearID(TXT, Optional YearDefault)
    For Each w In TXT.Range.Words
        y = Val(w)
        If y > 1000 Then FindYearID = y: Exit Function
    Next
    FindYearID = Switch(IsMissing(YearDefault), Year(Now), True, YearDefault)
End Function


Function FindMonthID(MonthTXT)
    For i = 1 To 12
        If InStr(1, MonthTXT, MonthName(i), vbTextCompare) Then FindMonthID = i: Exit Function
    Next
End Function


Function IsHorizontalCalendar(Table)  'check table width (week)
    IsHorizontalCalendar = (Table.Rows(3).Cells.Count = 7)
End Function


Function IsVerticalCalendar(Table)  'check table height (month)
    IsVerticalCalendar = (Table.Rows.Count > 28)
End Function


Sub FillCalendarTables(Optional ColorSun, Optional ColorSat, Optional LeadZero = True, Optional YearAll)
If IsMissing(YearAll) Then YearAll = FindYearID(ActiveDocument.Paragraphs(1))
ColorSunSat = Not (IsMissing(ColorSun) Or IsMissing(ColorSat))
If LeadZero Then lz = "00" Else lz = "0"
For Each Table In ActiveDocument.Tables
    MonthID = FindMonthID(Table.Rows(1))
    If Not IsEmpty(MonthID) Then
        YearID = FindYearID(Table.Rows(1), YearAll)
        DaysCountInMonth = Day(DateSerial(YearID, MonthID + 1, 0))

        If IsVerticalCalendar(Table) Then
            Call ColorClear(Table)
            For r = 2 To Table.Rows.Count: Set Row = Table.Rows(r)  'omit 1 row header
                DayID = Row.Cells(1).RowIndex - 1  'minus header
                WeekdayID = Weekday(DateSerial(YearID, MonthID, DayID), vbSunday)
                If DayID <= DaysCountInMonth Then
                    Row.Cells(1).Range.Text = Format(DayID, lz)
                    Row.Cells(2).Range.Text = WeekdayName(WeekdayID, True, vbSunday)
                    If ColorSunSat Then Call ColorWeekend(Row, WeekdayID, ColorSun, ColorSat)
                Else
                    Row.Cells(1).Range.Text = ""
                    Row.Cells(2).Range.Text = ""
                End If
            Next

        ElseIf IsHorizontalCalendar(Table) Then
            DayID = 1
            For r = 3 To Table.Rows.Count: Set Row = Table.Rows(r)  'omit 2 rows header
                Row.Range.Delete
                WeekdayID = Weekday(DateSerial(YearID, MonthID, DayID), vbUseSystemDayOfWeek)
                For d = WeekdayID To 7
                    If DayID <= DaysCountInMonth Then
                        Row.Cells(d) = Format(DayID, lz)
                        DayID = DayID + 1
                    End If
                Next
            Next
        End If

    End If
Next
End Sub


Function FindTables(HeaderTXT)
    Set FindTables = New Collection
    For Each Table In ActiveDocument.Tables
        If LCase(Table.Rows(1)) Like "*" & LCase(HeaderTXT) & "*" Then FindTables.Add (Table)
    Next
End Function


Sub PasteInRangeEnd(ByVal r)
    r.Collapse (wdCollapseEnd)
    r.MoveEnd Count:=-1
    r.Paste
End Sub


Function HorCal_DayToCell(Table, DayID)
    For r = 3 To Table.Rows.Count  'omit 2 rows header
        For d = 1 To 7  'week
            If Val(Table.Cell(r, d)) = DayID Then
                Set HorCal_DayToCell = Table.Cell(r, d)
                Exit Function
            End If
        Next
    Next
End Function


Sub DeleteInlineShapes(Table)
    For Each Shape In Table.Range.InlineShapes
        Shape.Delete
    Next
End Sub


Sub DeleteLegendIcons()
For Each Table In ActiveDocument.Tables
    If IsVerticalCalendar(Table) Or IsHorizontalCalendar(Table) Then
        Call DeleteInlineShapes(Table)
    End If
Next
End Sub


Sub InsertLegendIcons()
Call DeleteLegendIcons
YearAll = FindYearID(ActiveDocument.Paragraphs(1))
For Each Legend In FindTables("legend")
    For c = 1 To Legend.Rows(2).Cells.Count
        Legend.Rows(2).Cells(c).Range.InlineShapes(1).Range.Copy  'icon
        For Each Line In Split(Legend.Rows(3).Cells(c), Chr$(13))  'dates
            If IsDate(Line) Then
                DateID = CDate(Line)
                DayID = Day(DateID)
                YearID = Year(DateID)
                For Each Table In FindTables(MonthName(Month(DateID)))
                    If FindYearID(Table.Rows(1), YearAll) = YearID Then
                        If IsVerticalCalendar(Table) Then
                            Set r = Table.Rows(DayID + 1).Cells(3).Range  'plus header
                        ElseIf IsHorizontalCalendar(Table) Then
                            Set r = HorCal_DayToCell(Table, DayID).Range
                        End If
                        PasteInRangeEnd (r)
                    End If
                Next
            End If
        Next
    Next
Next
End Sub


Sub FillCalendarTablesGray()
    Call FillCalendarTables(wdColorGray15, wdColorGray10)
End Sub


Sub FillCalendarTablesRed()
    Call FillCalendarTables(&H9B9BFC, wdColorGray10)
End Sub
