Attribute VB_Name = "Module2"
Sub Add_Data()
Dim First, Second, SheetName As String
Dim ArrFirst(), ArrSecond()
Dim lRow, lRow2, lRow3, a, b, c As Integer
Dim ave, d As Double
Application.ScreenUpdating = False

'Build Arrays of names again
    Sheets("Tables").Activate
    lRow = Cells(1, 1).End(xlDown).Row
    lRow2 = Cells(1, 1).End(xlDown).End(xlDown).Row
    lRow3 = Cells(1, 1).End(xlDown).End(xlDown).End(xlDown).Row
    First = Cells(2, 1).Value 'Location
    For a = 0 To (lRow - 3)
        ReDim Preserve ArrFirst(a)
        ArrFirst(a) = Cells(3 + a, 1).Value
    Next a
    Second = Cells(lRow2, 1).Value ' Power
    For a = 0 To (lRow3 - (lRow2 + 1))
        ReDim Preserve ArrSecond(a)
        ArrSecond(a) = Cells(lRow2 + 1 + a, 1).Value
    Next a

Worksheets("Data Display").Activate
If Worksheets("Data Display").ChartObjects.Count = 2 Then
    'New Objects
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
        ActiveChart.Parent.Name = "Third Chart"
        Application.CutCopyMode = False
        ActiveSheet.Shapes("Third Chart").Left = Range(Cells(28, 2), Cells(43, (2 * UBound(ArrFirst)) + 5)).Left
        ActiveSheet.Shapes("Third Chart").Top = Range(Cells(28, 2), Cells(43, (2 * UBound(ArrFirst)) + 5)).Top
        ActiveSheet.Shapes("Third Chart").Width = Range(Cells(28, 2), Cells(43, (2 * UBound(ArrFirst)) + 5)).Width
        ActiveSheet.Shapes("Third Chart").Height = Range(Cells(28, 2), Cells(43, (2 * UBound(ArrFirst)) + 5)).Height

    
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
        ActiveChart.Parent.Name = "Fourth Chart"
        Application.CutCopyMode = False
        ActiveSheet.Shapes("Fourth Chart").Left = Range(Cells(28, (2 * UBound(ArrFirst)) + 11), Cells(43, (2 * UBound(ArrFirst)) + 14 + (2 * UBound(ArrSecond)))).Left
        ActiveSheet.Shapes("Fourth Chart").Top = Range(Cells(28, (2 * UBound(ArrFirst)) + 11), Cells(43, (2 * UBound(ArrFirst)) + 14 + (2 * UBound(ArrSecond)))).Top
        ActiveSheet.Shapes("Fourth Chart").Width = Range(Cells(28, (2 * UBound(ArrFirst)) + 11), Cells(43, (2 * UBound(ArrFirst)) + 14 + (2 * UBound(ArrSecond)))).Width
        ActiveSheet.Shapes("Fourth Chart").Height = Range(Cells(28, (2 * UBound(ArrFirst)) + 11), Cells(43, (2 * UBound(ArrFirst)) + 14 + (2 * UBound(ArrSecond)))).Height
        
    ActiveSheet.GroupBoxes.Add(Cells(15, (2 * UBound(ArrFirst)) + 7).Left, Cells(15, (2 * UBound(ArrFirst)) + 7).Top, Range(Cells(15, (2 * UBound(ArrFirst)) + 7), Cells(15, (2 * UBound(ArrFirst)) + 9)).Width, Range(Cells(15, (2 * UBound(ArrFirst)) + 7), Cells(15, (2 * UBound(ArrFirst)) + 9)).Height - 6).Select
    Selection.Characters.Text = "Additional Data Selection"
    Selection.Name = "Additional"
End If

' Userform to Build Sheet
    UserForm2.Show
    ThisWorkbook.Sheets.Add After:=ActiveSheet
    SheetName = UserForm2.Controls("txtTitle").Text
    SheetName = Replace(SheetName, " ", "_")
    ActiveSheet.Name = SheetName
    For a = 0 To ((UserForm2.Controls("Label4").Caption) - 1)
        Cells(1, ((UBound(ArrFirst) + 5) * a) + 1).Value = UserForm2.Controls("txtSeries" & (a + 1)).Value
        Cells(2, ((UBound(ArrFirst) + 5) * a) + 3).Value = First
        For b = 0 To UBound(ArrFirst)
            Cells(3, ((UBound(ArrFirst) + 5) * a) + 3 + b).Value = ArrFirst(b)
        Next b
        Cells(4, ((UBound(ArrFirst) + 5) * a) + 1).Value = Second
        For b = 0 To UBound(ArrSecond)
            Cells(4 + b, ((UBound(ArrFirst) + 5) * a) + 2).Value = ArrSecond(b)
        Next b
  
    'Manually Type Data (Cell by cell)
        Application.ScreenUpdating = True
        For b = 0 To UBound(ArrFirst)
            For c = 0 To UBound(ArrSecond)
                Cells(c + 4, b + 3 + ((UBound(ArrFirst) + 5) * a)).Select
                Selection.Value = InputBox("Average Value For " & UserForm2.Controls("txtSeries" & (a + 1)).Text & _
                    vbCrLf & " when " & vbCrLf & First & " = " & ArrFirst(b) & vbCrLf & " and " & vbCrLf & _
                    Second & " = " & ArrSecond(c) & ":", "1", "1")
            Next c
        Next b
        Application.ScreenUpdating = False
    
    'Calculate Averges
        For c = 0 To UBound(ArrFirst)
            Cells(5 + UBound(ArrSecond), ((UBound(ArrFirst) + 5) * a) + 3 + c).FormulaR1C1 = "=AVERAGE(R[" & (-1) * (1 + UBound(ArrSecond)) & "]C:R[-1]C)"
        Next c
        For c = 0 To UBound(ArrSecond)
            Cells(c + 4, ((UBound(ArrFirst) + 5) * a) + UBound(ArrFirst) + 4).FormulaR1C1 = "=AVERAGE(R[0]C[" & ((-1) * (1 + UBound(ArrFirst))) & "]:R[0]C[-1])"
        Next c
    Next a
'Plot
    'Third Chart
        Worksheets(SheetName).Activate
           c = 5
           For b = 0 To (UBound(ArrSecond) + 1)
           If Worksheets(SheetName).Cells(5, 4).Value = "" Then
                lRow = 1
                Else
                    lRow = (Worksheets(SheetName).Cells(5, 4).End(xlToRight).Column) - 2
           End If
           Worksheets("Data Display").Activate
           ActiveSheet.ChartObjects("Third Chart").Activate
           For a = 0 To (UserForm2.Controls("Label4").Caption - 1)

                d = ActiveChart.FullSeriesCollection.Count
                    ActiveChart.SeriesCollection.NewSeries
                    If b = (UBound(ArrSecond) + 1) Then
                    ActiveChart.FullSeriesCollection(d + 1).Name = "Ave." & "_" & Worksheets(SheetName).Cells(1, ((UBound(ArrFirst) + 5) * a) + 1).Value
                    Else
                    ActiveChart.FullSeriesCollection(d + 1).Name = Worksheets(SheetName).Cells(1, ((UBound(ArrFirst) + 5) * a) + 1).Value & "_" & Worksheets(SheetName).Cells(4 + b, ((UBound(ArrFirst) + 5) * a) + 2)
                    End If
                    ActiveChart.FullSeriesCollection(d + 1).Values = "='" & SheetName & "'!" & Range(Cells(4 + b, ((UBound(ArrFirst) + 5) * a) + 3), Cells(4 + b, ((UBound(ArrFirst) + 5) * a) + 1 + lRow)).Address
                    ActiveChart.FullSeriesCollection(d + 1).XValues = "='" & SheetName & "'!" & Range(Cells(3, 3), Cells(3, 1 + lRow)).Address
                ActiveChart.FullSeriesCollection(d + 1).Format.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
                    ActiveChart.FullSeriesCollection(d + 1).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
                    c = c + 1
                    If c = 12 Then
                     c = 5
                    End If
           Next a
        Next b
        ActiveChart.SetElement (msoElementLegendTop)
        ActiveChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
        ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
        ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
        With ActiveChart.Axes(xlValue)
            .HasTitle = True
            With .AxisTitle
                .Caption = "Average"
            End With
        End With
        ActiveChart.Axes(xlCategory).AxisTitle.Select
        ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = First
        Selection.Format.TextFrame2.TextRange.Characters.Text = First
        ActiveChart.SetElement (msoElementChartTitleNone)
    'Fourth chart
        c = 5
        Worksheets(SheetName).Activate
        For b = 0 To (UBound(ArrFirst) + 1)
           If Worksheets(SheetName).Cells(5, 2).Value = "" Then
                lRow = 1
                Else
                    lRow = (Worksheets(SheetName).Cells(4, 3).End(xlDown).Row) - 5
           End If
           Worksheets("Data Display").Activate
           ActiveSheet.ChartObjects("Fourth Chart").Activate
           For a = 0 To (UserForm2.Controls("Label4").Caption - 1)
                d = (ActiveChart.FullSeriesCollection.Count)
                     ActiveChart.SeriesCollection.NewSeries
                         If b = (UBound(ArrFirst) + 1) Then
                         ActiveChart.FullSeriesCollection(d + 1).Name = "Ave." & "_" & Worksheets(SheetName).Cells(1, ((UBound(ArrFirst) + 5) * a) + 1).Value
                         Else
                         ActiveChart.FullSeriesCollection(d + 1).Name = Worksheets(SheetName).Cells(1, ((UBound(ArrFirst) + 5) * a) + 1).Value & "_" & Worksheets(SheetName).Cells(3, ((UBound(ArrFirst) + 5) * a) + 3 + b)
                         End If
                         ActiveChart.FullSeriesCollection(d + 1).Values = "='" & SheetName & "'!" & Range(Cells(4, ((UBound(ArrFirst) + 5) * a) + 3 + b), Cells(4 + lRow, ((UBound(ArrFirst) + 5) * a) + 3 + b)).Address
                         ActiveChart.FullSeriesCollection(d + 1).XValues = "='" & SheetName & "'!" & Range(Cells(4, ((UBound(ArrFirst) + 5) * a) + 2), Cells(4 + lRow, ((UBound(ArrFirst) + 5) * a) + 2)).Address
                     ActiveChart.FullSeriesCollection(d + 1).Format.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
                         ActiveChart.FullSeriesCollection(d + 1).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
                         c = c + 1
                         If c = 12 Then
                          c = 5
                         End If
           Next a
        Next b
        ActiveChart.SetElement (msoElementLegendTop)
        ActiveChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
        ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
        ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
        With ActiveChart.Axes(xlValue)
            .HasTitle = True
            With .AxisTitle
                .Caption = "Average"
            End With
        End With
        ActiveChart.Axes(xlCategory).AxisTitle.Select
        ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = Second
        Selection.Format.TextFrame2.TextRange.Characters.Text = Second
        ActiveChart.SetElement (msoElementChartTitleNone)
    'Extend Frame, add radio buttons and asign controls
        ActiveSheet.Shapes.Range(Array("Additional")).Height = ActiveSheet.Shapes.Range(Array("Additional")).Height + Cells(1, 1).Height
        If Cells(10, 1).Value = "" Then
        Cells(10, 1).Value = UserForm2.Controls("Label4").Caption
        Cells(10, 1).Locked = True
        ActiveSheet.OptionButtons.Add(ActiveSheet.Shapes.Range(Array("Additional")).Left + 3, ActiveSheet.Shapes.Range(Array("Additional")).Top + (Cells(2, 2).Top * 1) - 4, ActiveSheet.Shapes.Range(Array("Additional")).Width - 6, Cells(2, 2).Height).Select
        Selection.Characters.Text = SheetName
        Selection.Name = "Additional0"
        Selection.OnAction = "Additional0_Click"
            Else
            If Cells(11, 1).Value = "" Then
            Cells(11, 1).Value = UserForm2.Controls("Label4").Caption
            Cells(11, 1).Locked = True
            ActiveSheet.OptionButtons.Add(ActiveSheet.Shapes.Range(Array("Additional")).Left + 3, ActiveSheet.Shapes.Range(Array("Additional")).Top + (Cells(2, 2).Top * 2) - 4, ActiveSheet.Shapes.Range(Array("Additional")).Width - 6, ActiveSheet.Shapes.Range(Array("Display10")).Height).Select
            Selection.Characters.Text = SheetName
            Selection.Name = "Additional1"
            Selection.OnAction = "Additional1_Click"
                Else
                If Cells(12, 1).Value = "" Then
                Cells(12, 1).Value = UserForm2.Controls("Label4").Caption
                Cells(12, 1).Locked = True
                ActiveSheet.OptionButtons.Add(ActiveSheet.Shapes.Range(Array("Additional")).Left + 3, ActiveSheet.Shapes.Range(Array("Additional")).Top + (Cells(2, 2).Top * 3) - 4, ActiveSheet.Shapes.Range(Array("Additional")).Width - 6, ActiveSheet.Shapes.Range(Array("Display10")).Height).Select
                Selection.Characters.Text = SheetName
                Selection.Name = "Additional2"
                Selection.OnAction = "Additional2_Click"
                    Else
                    Cells(13, 1).Value = UserForm2.Controls("Label4").Caption
                    Cells(13, 1).Locked = True
                    ActiveSheet.OptionButtons.Add(ActiveSheet.Shapes.Range(Array("Additional")).Left + 3, ActiveSheet.Shapes.Range(Array("Additional")).Top + (Cells(2, 2).Top * 4) - 4, ActiveSheet.Shapes.Range(Array("Additional")).Width - 6, ActiveSheet.Shapes.Range(Array("Display10")).Height).Select
                    Selection.Characters.Text = SheetName
                    Selection.Name = "Additional3"
                    Selection.OnAction = "Additional3_Click"
                End If
            End If
        End If
        ' can then use value to inform how many series ther is per added set
        Cells(8, 1).Value = UBound(ArrFirst) + 1
        Cells(8, 1).Locked = True
        Cells(9, 1).Value = UBound(ArrSecond) + 1
        Cells(9, 1).Locked = True
    Range("A1:A13").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
'Unload (UserForm2)
Cells(2, 2).Select
Application.ScreenUpdating = True
End Sub
