Attribute VB_Name = "Module11"
Sub NewWB_Data_Table()
    Dim fPath As String, fName As String, Location As String
    Dim lColumn, lRow, lRow2 As Integer
    Dim FileSeporator As String
       Dim k, k1, k2, a, b, c, d, m, n, o, p, q As Integer
       Dim ZCount(), Area(), MaxArea(), MinArea(), Circ(), Feret(), Angle(), GlobalArea(), Crust_Thick(), Hole_Area(), Hole_Circ()
       Dim SDCount(), SDArea(), SDMaxArea(), SDMinArea(), SDCirc(), SDFeret(), SDAngle(), SDGlobalArea(), SDCrust_Thick(), SDHole_Area(), SDHole_Circ()
    
    FileSeporator = "\"
    lColumn = ActiveSheet.Cells(3, Columns.Count).End(xlToLeft).Column
    fPath = Application.ThisWorkbook.Path & FileSeporator
    fName = ActiveWorkbook.Name
    
    Dim ModulePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select code for Button Control 'Image Select Buttons'"
        .Filters.Add "VB Code", "*.bas", 1
        .AllowMultiSelect = False
        .Show
        ModulePath = .SelectedItems.Item(1)
    End With
    
    ' Find Titles
       Dim Row8 As String, Row9 As String
       Row8 = Cells(10, 1).Value
       Row9 = Cells(11, 1).Value
    
    ' Find Array of unique values in Power/Location Rows
       Dim Repeat As String, ArrRow8(), ArrRow9()
       ReDim ArrRow8(0), ArrRow9(0)
       ArrRow8(0) = 17
       ArrRow9(0) = 17
       a = 0
       
       Range(Cells(10, 2), Cells(10, lColumn)).Select
       For Each Cell In Selection
            Repeat = ""
            If (Cell <> "") Then
                For q = 0 To UBound(ArrRow8)
                    If Cell = ArrRow8(q) Then
                        Repeat = Repeat & "X"
                    Else
                        Repeat = Repeat & "_"
                    End If
                Next q
                If InStr(Repeat, "X") = 0 Then
                    If ArrRow8(0) = 17 And a = 0 Then
                    ArrRow8(0) = Cell
                    a = 1
                    Else
                    ReDim Preserve ArrRow8(UBound(ArrRow8) + 1)
                    ArrRow8(UBound(ArrRow8)) = Cell
                    End If
                End If
            End If
        Next Cell
       a = 0
       Range(Cells(11, 2), Cells(11, lColumn)).Select
       For Each Cell In Selection
            Repeat = ""
            If (Cell <> "") Then
                For q = 0 To UBound(ArrRow9)
                    If Cell = ArrRow9(q) Then
                        Repeat = Repeat & "X"
                    Else
                        Repeat = Repeat & "_"
                    End If
                Next q
                If InStr(Repeat, "X") = 0 Then
                    If ArrRow9(0) = 17 And a = 0 Then
                    ArrRow9(0) = Cell
                    a = 1
                    Else
                    ReDim Preserve ArrRow9(UBound(ArrRow9) + 1)
                    ArrRow9(UBound(ArrRow9)) = Cell
                    End If
                End If
            End If
        Next Cell
       
       If ArrRow8(0) = 17 Then
        ReDim ArrRow8(0)
        Row8 = "(blank)"
       End If
       If ArrRow9(0) = 17 Then
        ReDim ArrRow9(0)
        Row9 = "(blank)"
       End If
       k1 = UBound(ArrRow8)
       k2 = UBound(ArrRow9)
       p = 0
       
   Application.ScreenUpdating = False
       Workbooks.Add.SaveAs fPath & "Results_Summary" & ".xlsm", FileFormat:=52
       Cells(1, 1).Value = Row8
       Cells(1, 2).Value = Row9
       For m = 0 To k1
        For n = 0 To k2
            Cells(p + 2, 1).Value = ArrRow8(m)
            Cells(p + 2, 2).Value = ArrRow9(n)
            p = p + 1
        Next n
       Next m
       
       ' Find all Values for vaiables
       Workbooks(fName).Activate
       p = 0
       q = 0
       For m = 0 To k1
            For n = 0 To k2
             ReDim ZCount(0), Area(0), MaxArea(0), MinArea(0), Circ(0), Feret(0), Angle(0), GlobalArea(0), Crust_Thick(0)
             ReDim Hole_Area(0), Hole_Circ(0), SDArea(0), SDMaxArea(0), SDMinArea(0), SDCirc(0), SDFeret(0), SDAngle(0)
             ReDim SDCount(0), SDGlobalArea(0), SDCrust_Thick(0), SDHole_Area(0), SDHole_Circ(0)
             p = 0
                For o = 2 To lColumn Step 19
                    If Cells(10, o).Value = ArrRow8(m) And Cells(11, o).Value = ArrRow9(n) Then
                        ReDim Preserve ZCount(p), Area(p), MaxArea(p), MinArea(p), Circ(p), Feret(p), Angle(p), GlobalArea(p), Crust_Thick(p)
                        ReDim Preserve Hole_Area(p), Hole_Circ(p), SDArea(p), SDMaxArea(p), SDMinArea(p), SDCirc(p), SDFeret(p), SDAngle(p)
                        ReDim Preserve SDCount(p), SDGlobalArea(p), SDCrust_Thick(p), SDHole_Area(p), SDHole_Circ(p)

                        If Cells(4, o).Value = "NaN" Then
                            ZCount(p) = 0
                            Else
                            ZCount(p) = Cells(4, o).Value
                        End If
                        If Cells(4, o + 1).Value = "NaN" Then
                            Area(p) = 0
                            Else
                            Area(p) = Cells(4, o + 1).Value
                        End If
                        If Cells(7, o + 1).Value = "NaN" Then
                            MaxArea(p) = 0
                            Else
                            MaxArea(p) = Cells(7, o + 1).Value
                        End If
                        If Cells(8, o + 1).Value = "NaN" Then
                            MinArea(p) = 0
                            Else
                            MinArea(p) = Cells(8, o + 1).Value
                        End If
                        If Cells(4, o + 9).Value = "NaN" Then
                            Circ(p) = 0
                            Else
                            Circ(p) = Cells(4, o + 9).Value
                        End If
                        If Cells(4, o + 10).Value = "NaN" Then
                            Feret(p) = 0
                            Else
                            Feret(p) = Cells(4, o + 10).Value
                        End If
                        If Cells(4, o + 13).Value = "NaN" Then
                            Angle(p) = 0
                            Else
                            Angle(p) = Cells(4, o + 13).Value
                        End If
                        If Cells(15, o + 7).Value = "NaN" Then
                            GlobalArea(p) = 0
                            Else
                            GlobalArea(p) = Cells(15, o + 7).Value
                        End If
                        If Cells(18, o + 4).Value = "NaN" Then
                            Crust_Thick(p) = 0
                            Else
                            Crust_Thick(p) = Cells(18, o + 4).Value
                        End If
                        If Cells(21, o).Value = "NaN" Then
                            Hole_Area(p) = 0
                            Else
                            Hole_Area(p) = Cells(21, o).Value
                        End If
                        If Cells(21, o + 10).Value = "NaN" Then
                            Hole_Circ(p) = 0
                            Else
                            Hole_Circ(p) = Cells(21, o + 10).Value
                        End If
                        If Cells(5, o + 1).Value = "NaN" Then
                            SDArea(p) = 0
                            Else
                            SDArea(p) = Cells(5, o + 1).Value
                        End If
                        If Cells(5, o + 9).Value = "NaN" Then
                            SDCirc(p) = 0
                            Else
                            SDCirc(p) = Cells(5, o + 9).Value
                        End If
                        If Cells(5, o + 10).Value = "NaN" Then
                            SDFeret(p) = 0
                            Else
                            SDFeret(p) = Cells(5, o + 10).Value
                        End If
                        If Cells(5, o + 13).Value = "NaN" Then
                            SDAngle(p) = 0
                            Else
                            SDAngle(p) = Cells(5, o + 13).Value
                        End If
                        
                        SDCount(p) = 0
                        SDMaxArea(p) = 0
                        SDMinArea(p) = 0
                        SDGlobalArea(p) = 0
                        SDCrust_Thick(p) = 0
                        SDHole_Area(p) = 0
                        SDHole_Circ(p) = 0

                        p = p + 1
                    End If
                Next o
                    Workbooks("Results_Summary").Activate
                                        
                  ' Count
                    Res = Calc_Stats(ZCount, ZCount, SDCount)
                    Cells(1, 3).Value = "Ave Count"
                    Cells(1, 14).Value = "SD. Count"
                    Cells(1, 25).Value = "+/-. Count"
                    Cells(q + 2, 3).Value = Res(0)  'Mean
                    Cells(q + 2, 14).Value = Res(1)  'StDev
                    Cells(q + 2, 25).Value = Res(2)  'Uncirt
                  ' Area
                    Res = Calc_Stats(ZCount, Area, SDArea)
                    Cells(1, 4).Value = "Ave Area (sq.mm)"
                    Cells(1, 15).Value = "SD. Area (sq.mm)"
                    Cells(1, 26).Value = "+/-. Area (sq.mm)"
                    Cells(q + 2, 4).Value = Res(0)  'Mean
                    Cells(q + 2, 15).Value = Res(1)  'StDev
                    Cells(q + 2, 26).Value = Res(2)  'Uncirt
                  ' Max Area
                    Res = Calc_Stats(ZCount, MaxArea, SDMaxArea)
                    Cells(1, 5).Value = "Ave Max. Area (sq.mm)"
                    Cells(1, 16).Value = "SD. Max. Area (sq.mm)"
                    Cells(1, 27).Value = "+/-. Max. Area (sq.mm)"
                    Cells(q + 2, 5).Value = Res(0)  'Mean
                    Cells(q + 2, 16).Value = Res(1)  'StDev
                    Cells(q + 2, 27).Value = Res(2)  'Uncirt
                  ' Min Area
                    Res = Calc_Stats(ZCount, MinArea, SDMinArea)
                    Cells(1, 6).Value = "Ave Min. Area (sq.mm)"
                    Cells(1, 17).Value = "SD. Min. Area (sq.mm)"
                    Cells(1, 28).Value = "+/-. Min. Area (sq.mm)"
                    Cells(q + 2, 6).Value = Res(0)  'Mean
                    Cells(q + 2, 17).Value = Res(1)  'StDev
                    Cells(q + 2, 28).Value = Res(2)  'Uncirt
                  ' Circ
                    Res = Calc_Stats(ZCount, Circ, SDCirc)
                    Cells(1, 7).Value = "Ave Circularity"
                    Cells(1, 18).Value = "SD. Circularity"
                    Cells(1, 29).Value = "+/-. Circularity"
                    Cells(q + 2, 7).Value = Res(0)  'Mean
                    Cells(q + 2, 18).Value = Res(1)  'StDev
                    Cells(q + 2, 29).Value = Res(2)  'Uncirt
                  ' Feret Length
                    Res = Calc_Stats(ZCount, Feret, SDFeret)
                    Cells(1, 8).Value = "Ave Feret Length (mm)"
                    Cells(1, 19).Value = "SD. Feret Length (mm)"
                    Cells(1, 30).Value = "+/-. Feret Length (mm)"
                    Cells(q + 2, 8).Value = Res(0)  'Mean
                    Cells(q + 2, 19).Value = Res(1)  'StDev
                    Cells(q + 2, 30).Value = Res(2)  'Uncirt
                  ' Angle
                    Res = Calc_Stats(ZCount, Angle, SDAngle)
                    Cells(1, 9).Value = "Ave Feret Angle (deg)"
                    Cells(1, 20).Value = "SD. Feret Angle (deg)"
                    Cells(1, 31).Value = "+/-. Feret Angle (deg)"
                    Cells(q + 2, 9).Value = Res(0)  'Mean
                    Cells(q + 2, 20).Value = Res(1)  'StDev
                    Cells(q + 2, 31).Value = Res(2)  'Uncirt
                  ' GlobalArea
                    Res = Calc_Stats(ZCount, GlobalArea, SDGlobalArea)
                    Cells(1, 10).Value = "Ave Global Porosity (%)"
                    Cells(1, 21).Value = "SD. Global Porosity (%)"
                    Cells(1, 32).Value = "+/-. Global Porosity (%)"
                    Cells(q + 2, 10).Value = Res(0)  'Mean
                    Cells(q + 2, 21).Value = Res(1)  'StDev
                    Cells(q + 2, 32).Value = Res(2)  'Uncirt
                  ' Crust_Thick
                    Res = Calc_Stats(ZCount, Crust_Thick, SDCrust_Thick)
                    Cells(1, 11).Value = "Ave Approx. Crust Thickness (mm)"
                    Cells(1, 22).Value = "SD. Approx. Crust Thickness (mm)"
                    Cells(1, 33).Value = "+/-. Approx. Crust Thickness (mm)"
                    Cells(q + 2, 11).Value = Res(0)  'Mean
                    Cells(q + 2, 22).Value = Res(1)  'StDev
                    Cells(q + 2, 33).Value = Res(2)  'Uncirt
                  ' Hole_Area
                    Res = Calc_Stats(ZCount, Hole_Area, SDHole_Area)
                    Cells(1, 12).Value = "Ave Hole Area (sq.mm)"
                    Cells(1, 23).Value = "SD. Hole Area (sq.mm)"
                    Cells(1, 34).Value = "+/-. Hole Area (sq.mm)"
                    Cells(q + 2, 12).Value = Res(0)  'Mean
                    Cells(q + 2, 23).Value = Res(1)  'StDev
                    Cells(q + 2, 34).Value = Res(2)  'Uncirt
                  ' Hole_Circ
                    Res = Calc_Stats(ZCount, Hole_Circ, SDHole_Circ)
                    Cells(1, 13).Value = "Ave Hole Circularity"
                    Cells(1, 24).Value = "SD. Hole Circularity"
                    Cells(1, 35).Value = "+/-. Hole Circularity"
                    Cells(q + 2, 13).Value = Res(0)  'Mean
                    Cells(q + 2, 24).Value = Res(1)  'StDev
                    Cells(q + 2, 35).Value = Res(2)  'Uncirt
                    
                    Workbooks(fName).Activate
                    q = q + 1
            Next n
       Next m
    
    Workbooks("Results_Summary").Activate
    lRow = (ActiveSheet.Cells(1, 1).End(xlDown).Row)
    For q = lRow To 2 Step -1
        If Cells(q, 3) = 0 Then
            Rows(q).EntireRow.Delete
        End If
    Next q
    lRow = (ActiveSheet.Cells(1, 1).End(xlDown).Row) + 1
    
    Rows(1).Insert
    
    Range(Cells(3, 3), Cells(lRow, 3)).NumberFormat = "0"
    Range(Cells(3, 4), Cells(lRow, 4)).NumberFormat = "0.00"
    Range(Cells(3, 5), Cells(lRow, 5)).NumberFormat = "0.00"
    Range(Cells(3, 6), Cells(lRow, 6)).NumberFormat = "0.00"
    Range(Cells(3, 7), Cells(lRow, 7)).NumberFormat = "0.00"
    Range(Cells(3, 8), Cells(lRow, 8)).NumberFormat = "0.00"
    Range(Cells(3, 9), Cells(lRow, 9)).NumberFormat = "0"
    Range(Cells(3, 10), Cells(lRow, 10)).NumberFormat = "0.0"
    Range(Cells(3, 11), Cells(lRow, 11)).NumberFormat = "0.00"
    Range(Cells(3, 12), Cells(lRow, 12)).NumberFormat = "0"
    Range(Cells(3, 13), Cells(lRow, 13)).NumberFormat = "0.00"
    
    Range(Cells(3, 14), Cells(lRow, 14)).NumberFormat = "0"
    Range(Cells(3, 15), Cells(lRow, 15)).NumberFormat = "0.00"
    Range(Cells(3, 16), Cells(lRow, 16)).NumberFormat = "0.00"
    Range(Cells(3, 17), Cells(lRow, 17)).NumberFormat = "0.00"
    Range(Cells(3, 18), Cells(lRow, 18)).NumberFormat = "0.00"
    Range(Cells(3, 19), Cells(lRow, 19)).NumberFormat = "0.00"
    Range(Cells(3, 20), Cells(lRow, 20)).NumberFormat = "0"
    Range(Cells(3, 21), Cells(lRow, 21)).NumberFormat = "0.0"
    Range(Cells(3, 22), Cells(lRow, 22)).NumberFormat = "0.00"
    Range(Cells(3, 23), Cells(lRow, 23)).NumberFormat = "0"
    Range(Cells(3, 24), Cells(lRow, 24)).NumberFormat = "0.00"
    
    Range(Cells(3, 25), Cells(lRow, 25)).NumberFormat = "0"
    Range(Cells(3, 26), Cells(lRow, 26)).NumberFormat = "0.00"
    Range(Cells(3, 27), Cells(lRow, 27)).NumberFormat = "0.00"
    Range(Cells(3, 28), Cells(lRow, 28)).NumberFormat = "0.00"
    Range(Cells(3, 29), Cells(lRow, 29)).NumberFormat = "0.00"
    Range(Cells(3, 30), Cells(lRow, 30)).NumberFormat = "0.00"
    Range(Cells(3, 31), Cells(lRow, 31)).NumberFormat = "0"
    Range(Cells(3, 32), Cells(lRow, 32)).NumberFormat = "0.0"
    Range(Cells(3, 33), Cells(lRow, 33)).NumberFormat = "0.00"
    Range(Cells(3, 34), Cells(lRow, 34)).NumberFormat = "0"
    Range(Cells(3, 35), Cells(lRow, 35)).NumberFormat = "0.00"
    
    Range(Cells(2, 1), Cells(2, 35)).EntireColumn.AutoFit
    
    Range(Cells(1, 3), Cells(1, 13)).Merge
    Range(Cells(1, 3), Cells(1, 13)).Value = "Mean Average"
    Range(Cells(1, 14), Cells(1, 24)).Merge
    Range(Cells(1, 14), Cells(1, 24)).Value = "Standard Deviation"
    Range(Cells(1, 25), Cells(1, 35)).Merge
    Range(Cells(1, 25), Cells(1, 35)).Value = "Uncirtanty (+/-)"
    
        ' sort assending order
    Dim tmp
    For a = LBound(ArrRow8) To UBound(ArrRow8)
        For b = a + 1 To UBound(ArrRow8)
            If ArrRow8(a) > ArrRow8(b) Then
                tmp = ArrRow8(b)
                ArrRow8(b) = ArrRow8(a)
                ArrRow8(a) = tmp
            End If
        Next b
    Next a
    For a = LBound(ArrRow9) To UBound(ArrRow9)
        For b = a + 1 To UBound(ArrRow9)
            If ArrRow9(a) > ArrRow9(b) Then
                tmp = ArrRow9(b)
                ArrRow9(b) = ArrRow9(a)
                ArrRow9(a) = tmp
            End If
        Next b
    Next a
    
    ' Make Image Pages
    Dim ArrNames(), ArrPicRow(), ArrPicCol(), ArrTypes()
    ArrNames = Array("Picture_Colour", "Picture_Area", "Picture_Area_Dis", "Picture_Circ", "Picture_Circ_Dis", _
                "Picture_Feret_Len", "Picture_Feret_len_dis", "Picture_Feret_Angle", "Picture_Porosity_Map")
    ArrPicRow = Array(23, 24, 24, 25, 25, 26, 26, 26, 23)
    ArrPicCol = Array(0, 4, 0, 4, 0, 4, 0, 11, 10)
    ArrTypes = Array("Count", "Area (sq.mm)", "Max. Area (sq.mm)", "Min. Area (sq.mm)", "Circularity", "Feret Length(mm)", "Feret Angle (deg)", "Global Area (sq.mm)", "Approx. Crust Thickness (mm)", "Hole Area (sq.mm)", "Hole Circularity")
    For d = 0 To UBound(ArrNames)
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = ArrNames(d)
        
        For a = 0 To UBound(ArrRow8)
            Cells(a + 2, 1).Value = ArrRow8(a)
        Next a
        For b = 0 To UBound(ArrRow9)
            Cells(1, (7 * b) + 2).Value = ArrRow9(b)
        Next b
        
        Workbooks(fName).Worksheets("Front Page").Activate
        For a = 0 To UBound(ArrRow8)
            For b = 0 To UBound(ArrRow9)
                For c = 2 To lColumn Step 19
                    If Cells(10, c).Value = ArrRow8(a) And Cells(11, c).Value = ArrRow9(b) Then
                        Workbooks(fName).Worksheets("Front Page").Activate
                        Cells(ArrPicRow(d), ArrPicCol(d) + c + 2).Select
                        Selection.Copy
                        Workbooks("Results_Summary").Worksheets(ArrNames(d)).Activate
                        Cells(a + 2, (7 * b) + 2).Select
                        Selection.RowHeight = 330
                        ActiveSheet.Paste
                        Application.CutCopyMode = False
                        Workbooks(fName).Worksheets("Front Page").Activate
                        GoTo NextIteration
                    End If
                Next c
NextIteration:
            Next b
        Next a
        Workbooks("Results_Summary").Worksheets(ArrNames(d)).Activate
    Next d
    
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Pivot"
                            
                            Dim sht As Worksheet
                            Dim pvtCache As PivotCache
                            Dim pvt As PivotTable
                            Dim StartPvt As String
                            Dim SrcData As String
                            
                            'Determine the data range you want to pivot
                              SrcData = "Sheet1!" & Range(Cells(2, 1), Cells(lRow, 35)).Address(ReferenceStyle:=xlR1C1)
                            
                            'Create a new worksheet
                              Set sht = Sheets("Pivot")
                            
                            'Where do you want Pivot Table to start?
                              StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)
                            
                            'Create Pivot Cache from Source Data
                              Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
                                SourceType:=xlDatabase, _
                                SourceData:=SrcData)
                            
                            'Create Pivot table from Pivot Cache
                              Set pvt = pvtCache.CreatePivotTable( _
                                TableDestination:=StartPvt, _
                                TableName:="PivotTable")

        Sheets("Pivot").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable")
        .ColumnGrand = False
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable").RepeatAllLabels xlRepeatLabels
    
    With ActiveSheet.PivotTables("PivotTable").PivotFields(Row8)
        .Orientation = xlRowField
        .Position = 1
        .Caption = Row8
    End With
    With ActiveSheet.PivotTables("PivotTable").PivotFields(Row9)
        .Orientation = xlPageField
        .Position = 1
    End With
    
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Ave Count"), "Ave. Count", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Ave Area (sq.mm)"), "Ave. Area (sq.mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Ave Max. Area (sq.mm)"), "Ave. Max. Area (sq.mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Ave Min. Area (sq.mm)"), "Ave. Min. Area (sq.mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Ave Circularity"), "Ave. Circularity", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Ave Feret Length (mm)"), "Ave. Feret Length (mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Ave Feret Angle (deg)"), "Ave. Feret Angle (deg)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Ave Global Porosity (%)"), "Ave. Global Porosity (%)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Ave Approx. Crust Thickness (mm)"), "Ave. Approx. Crust Thickness (mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Ave Hole Area (sq.mm)"), "Ave. Hole Area (sq.mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("Ave Hole Circularity"), "Ave. Hole Circularity", xlAverage
        
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("SD. Count"), "SD Count", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("SD. Area (sq.mm)"), "SD Area (sq.mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("SD. Max. Area (sq.mm)"), "SD Max. Area (sq.mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("SD. Min. Area (sq.mm)"), "SD Min. Area (sq.mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("SD. Circularity"), "SD Circularity", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("SD. Feret Length (mm)"), "SD Feret Length (mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("SD. Feret Angle (deg)"), "SD Feret Angle (deg)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("SD. Global Porosity (%)"), "SD Global Porosity (%)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("SD. Approx. Crust Thickness (mm)"), "SD Approx. Crust Thickness (mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("SD. Hole Area (sq.mm)"), "SD Hole Area (sq.mm)", xlAverage
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables( _
        "PivotTable").PivotFields("SD. Hole Circularity"), "SD Hole Circularity", xlAverage
    
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Tables"
    Cells(1, 1).Value = "Average"
    Cells(2, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    lRow = (ActiveSheet.Cells(3, 1).End(xlDown).Row)
    If Cells(lRow, 1).Value = "(blank)" Then
        Cells(lRow, 1).Value = ""
    End If
    Cells(2, 1).Value = Row8
    
    If IsEmpty(ArrRow8(0)) Then
    ArrRow8(0) = "(blank)"
    End If
    If IsEmpty(ArrRow9(0)) Then
    ArrRow9(0) = "(blank)"
    End If
    
    For a = 0 To UBound(ArrRow9)
        Worksheets("Pivot").Activate
        ActiveSheet.PivotTables("PivotTable").PivotFields(Row9).CurrentPage = ArrRow9(a)
        Range("A3").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Worksheets("Tables").Activate
        Cells(1, (25 * a) + 26).Value = ArrRow9(a)
        Cells(2, (25 * a) + 26).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        If Cells(lRow, (25 * a) + 26).Value = "(blank)" Then
            Cells(lRow, (25 * a) + 26).Value = ""
        End If
        Cells(2, (25 * a) + 26).Value = Row8
    Next a
    
    Worksheets("Pivot").Activate
    ActiveSheet.PivotTables("PivotTable").PivotFields(Row8).Orientation = xlHidden
    ActiveSheet.PivotTables("PivotTable").PivotFields(Row9).Orientation = xlHidden
    With ActiveSheet.PivotTables("PivotTable").PivotFields(Row9)
        .Orientation = xlRowField
        .Position = 1
        .Caption = Row9
    End With
    With ActiveSheet.PivotTables("PivotTable").PivotFields(Row8)
        .Orientation = xlPageField
        .Position = 1
    End With
    
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Worksheets("Tables").Activate
    Cells(lRow + 2, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells((lRow + 2), 1).Value = Row9
    
    For a = 0 To UBound(ArrRow8)
        Worksheets("Pivot").Activate
        ActiveSheet.PivotTables("PivotTable").PivotFields(Row8).CurrentPage = ArrRow8(a)
        Range("A3").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Worksheets("Tables").Activate
        Cells(lRow + 1, (25 * a) + 26).Value = ArrRow8(a)
        Cells(lRow + 2, (25 * a) + 26).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Cells((lRow + 2), (25 * a) + 26).Value = Row9
    Next a
    
For a = 0 To (UBound(ArrRow8) + 1)
    lRow = ActiveSheet.Cells(lRow, 1).End(xlDown).End(xlDown).Row
    If Cells(lRow, (25 * a) + 1).Value = "(blank)" Then
        Cells(lRow, (25 * a) + 1).Value = ""
    End If
Next a

Application.CutCopyMode = False

lRow = Cells(1, 1).End(xlDown).Row
If ArrRow9(0) = "(blank)" Then
    lRow2 = (Cells(1, 1).End(xlDown).End(xlDown).Row) + 1
Else
    lRow2 = Cells(lRow + 3, 1).End(xlDown).Row
End If
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Data Display"
    Range("B8").Select
    
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
    ActiveChart.Parent.Name = "First Chart"
    Application.CutCopyMode = False
    
    c = 5
    For a = 0 To (UBound(ArrRow9) + 1)
        If Worksheets("Tables").Cells(4, (25 * a) + 2).Value = "" Then
            lRow = 3
            Else
                lRow = Worksheets("Tables").Cells(3, (25 * a) + 2).End(xlDown).Row
        End If
        
        For d = 0 To UBound(ArrRow8)
            If Not (Worksheets("Tables").Cells(3 + d, (25 * a) + 1).Value = ArrRow8(d)) Then
                Worksheets("Tables").Activate
                If (d + 3) > lRow Then
                    ActiveSheet.Range(Cells(d + 3, (25 * a) + 1), Cells(d + 3, (25 * a) + 24)).Select
                    Selection.Cut
                    Selection.Offset(1, 0).Select
                    ActiveSheet.Paste
                Else
                    ActiveSheet.Range(Cells(d + 3, (25 * a) + 1), Cells(lRow, (25 * a) + 24)).Select
                    Selection.Cut
                    Selection.Offset(1, 0).Select
                    ActiveSheet.Paste
                End If
                If d = 0 Then
                    If ArrRow9(0) = 0 Then
                        ActiveSheet.Range(Cells(3, 26), Cells(3, 49)).Select
                        Selection.Copy
                        Cells(3, (25 * a) + 1).Select
                        ActiveSheet.Paste
                    End If
                End If
                Cells(3 + d, (25 * a) + 1).Value = ArrRow8(d)
                Worksheets("Data Display").Activate
                lRow = lRow + 1
            End If
        Next d

        For b = 1 To 11
            ActiveChart.SeriesCollection.NewSeries
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b) - 1).Name = Worksheets("Tables").Cells(1, (25 * a) + 1).Value & "_" & Worksheets("Tables").Cells(2, (25 * a) + b + 1).Value
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b) - 1).Values = "=Tables!" & Range(Cells(3, (25 * a) + b + 1), Cells(lRow, (25 * a) + b + 1)).Address
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b) - 1).XValues = "=Tables!" & Range(Cells(3, (25 * a) + 1), Cells(lRow, (25 * a) + 1)).Address
                ActiveChart.SeriesCollection.NewSeries
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).Name = Worksheets("Tables").Cells(1, (25 * a) + 1).Value & "_" & Worksheets("Tables").Cells(2, (25 * a) + b + 12).Value
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).Values = "=Tables!" & Range(Cells(3, (25 * a) + b + 12), Cells(lRow, (25 * a) + b + 12)).Address
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).XValues = "=Tables!" & Range(Cells(3, (25 * a) + 1), Cells(lRow, (25 * a) + 1)).Address

            ActiveChart.FullSeriesCollection((22 * a) + (2 * b) - 1).Format.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b) - 1).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).Format.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).Format.Line.Transparency = 1
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
            
                c = c + 1
                If c = 13 Then
                 c = 5
                End If
                If c = 11 Then
                 c = 12
                End If
            
            ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).AxisGroup = 2
        Next b
    Next a
    
    ActiveChart.SetElement (msoElementLegendTop)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    
    ActiveChart.SetElement (msoElementSecondaryValueAxisTitleAdjacentToAxis)
    ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Text = "Standard Deviation"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Standard Deviation"
    
    With ActiveChart.Axes(xlValue)
        .HasTitle = True
        With .AxisTitle
            .Caption = "Average"
        End With
    End With
    
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = Row8
    Selection.Format.TextFrame2.TextRange.Characters.Text = Row8
    ActiveChart.SetElement (msoElementChartTitleNone)
    ActiveSheet.Shapes("First Chart").Left = Range(Cells(2, 2), Cells(17, (2 * UBound(ArrRow8)) + 5)).Left
    ActiveSheet.Shapes("First Chart").Top = Range(Cells(2, 2), Cells(17, (2 * UBound(ArrRow8)) + 5)).Top
    ActiveSheet.Shapes("First Chart").Width = Range(Cells(2, 2), Cells(17, (2 * UBound(ArrRow8)) + 5)).Width
    ActiveSheet.Shapes("First Chart").Height = Range(Cells(2, 2), Cells(17, (2 * UBound(ArrRow8)) + 5)).Height
    
    'Second Chart
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
    ActiveChart.Parent.Name = "Second Chart"
    Application.CutCopyMode = False
    lRow = lRow + 3
    c = 5
    For a = 0 To (UBound(ArrRow8) + 1)
        If Worksheets("Tables").Cells(8 + UBound(ArrRow8), (25 * a) + 1).Value = "" Then
            lRow = 7 + UBound(ArrRow8)
            Else
                lRow = Worksheets("Tables").Cells(7 + UBound(ArrRow8), (25 * a) + 1).End(xlDown).Row
        End If
        
        If ArrRow9(0) = "(blank)" Then
        n = Worksheets("Tables").Cells(1, 1).End(xlDown).End(xlDown).Row
        Else
        n = Worksheets("Tables").Cells(1, 1).End(xlDown).End(xlDown).Row
        End If
        For d = 0 To UBound(ArrRow9)
            If Not (Worksheets("Tables").Cells(1 + n + d, (25 * a) + 1).Value = ArrRow9(d)) Then
                Worksheets("Tables").Activate
                If (d + 3 + n) > lRow Then
                    ActiveSheet.Range(Cells(d + 1 + n, (25 * a) + 1), Cells(d + 1 + n, (25 * a) + 24)).Select
                    Selection.Cut
                    Selection.Offset(1, 0).Select
                    ActiveSheet.Paste
                Else
                    ActiveSheet.Range(Cells(d + 1 + n, (25 * a) + 1), Cells(lRow, (25 * a) + 24)).Select
                    Selection.Cut
                    Selection.Offset(1, 0).Select
                    ActiveSheet.Paste
                End If
                If d = 0 Then
                    If ArrRow9(0) = 0 Then
                        ActiveSheet.Range(Cells(1 + n, 26), Cells(1 + n, 49)).Select
                        Selection.Copy
                        Cells(1 + n, (25 * a) + 1).Select
                        ActiveSheet.Paste
                    End If
                End If
                Cells(1 + n + d, (25 * a) + 1).Value = ArrRow9(d)
                Worksheets("Data Display").Activate
                lRow = lRow + 1
            End If
        Next d
       
       For b = 1 To 11
            ActiveChart.SeriesCollection.NewSeries
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b) - 1).Name = Worksheets("Tables").Cells(n - 1, (25 * a) + 1).Value & "_" & Worksheets("Tables").Cells(n, (25 * a) + b + 1)
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b) - 1).Values = "=Tables!" & Range(Cells(1 + n, (25 * a) + b + 1), Cells(lRow, (25 * a) + b + 1)).Address
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b) - 1).XValues = "=Tables!" & Range(Cells(1 + n, (25 * a) + 1), Cells(lRow, (25 * a) + 1)).Address
                ActiveChart.SeriesCollection.NewSeries
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).Name = Worksheets("Tables").Cells(n - 1, (25 * a) + 1).Value & "_" & Worksheets("Tables").Cells(n, (25 * a) + b + 12)
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).Values = "=Tables!" & Range(Cells(1 + n, (25 * a) + b + 12), Cells(lRow, (25 * a) + b + 12)).Address
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).XValues = "=Tables!" & Range(Cells(1 + n, (25 * a) + 1), Cells(lRow, (25 * a) + 1)).Address
            ActiveChart.FullSeriesCollection((22 * a) + (2 * b) - 1).Format.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b) - 1).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).Format.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).Format.Line.Transparency = 1
                ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent & c
                c = c + 1
                If c = 13 Then
                 c = 5
                End If
                If c = 11 Then
                 c = 12
                End If
            
            ActiveChart.FullSeriesCollection((22 * a) + (2 * b)).AxisGroup = 2
        Next b
    Next a
    
    ActiveChart.SetElement (msoElementLegendTop)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesMinorMajor)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    
    ActiveChart.SetElement (msoElementSecondaryValueAxisTitleAdjacentToAxis)
    ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Text = "Standard Deviation"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Standard Deviation"
    
    With ActiveChart.Axes(xlValue)
        .HasTitle = True
        With .AxisTitle
            .Caption = "Average"
        End With
    End With
    
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = Row9
    Selection.Format.TextFrame2.TextRange.Characters.Text = Row9
    
    ActiveChart.SetElement (msoElementChartTitleNone)

    ActiveSheet.Shapes("Second Chart").Left = Range(Cells(2, (2 * UBound(ArrRow8)) + 11), Cells(17, (2 * UBound(ArrRow8)) + 14 + (2 * UBound(ArrRow9)))).Left
    ActiveSheet.Shapes("Second Chart").Top = Range(Cells(2, (2 * UBound(ArrRow8)) + 11), Cells(17, (2 * UBound(ArrRow8)) + 14 + (2 * UBound(ArrRow9)))).Top
    ActiveSheet.Shapes("Second Chart").Width = Range(Cells(2, (2 * UBound(ArrRow8)) + 11), Cells(17, (2 * UBound(ArrRow8)) + 14 + (2 * UBound(ArrRow9)))).Width
    ActiveSheet.Shapes("Second Chart").Height = Range(Cells(2, (2 * UBound(ArrRow8)) + 11), Cells(17, (2 * UBound(ArrRow8)) + 14 + (2 * UBound(ArrRow9)))).Height
    
    Workbooks("Results_Summary").VBProject.VBComponents.Import ModulePath
    
    'Add Radio Buttons
    ActiveSheet.GroupBoxes.Add(Cells(2, (2 * UBound(ArrRow8)) + 7).Left, Cells(2, (2 * UBound(ArrRow8)) + 7).Top, Range(Cells(2, (2 * UBound(ArrRow8)) + 7), Cells(UBound(ArrTypes) + 4, (2 * UBound(ArrRow8)) + 9)).Width, Range(Cells(2, (2 * UBound(ArrRow8)) + 7), Cells(UBound(ArrTypes) + 4, (2 * UBound(ArrRow8)) + 9)).Height - 6).Select
        Selection.Characters.Text = "Select Charictoristic to Display"
        Selection.Name = "Displays"
    For a = 0 To UBound(ArrTypes)
        ActiveSheet.OptionButtons.Add(Cells(a + 3, (2 * UBound(ArrRow8)) + 7).Left + 3, Cells(a + 3, (2 * UBound(ArrRow8)) + 7).Top, Range(Cells(a + 3, (2 * UBound(ArrRow8)) + 7), Cells(a + 3, (2 * UBound(ArrRow8)) + 9)).Width - 6, Cells(a + 3, (2 * UBound(ArrRow8)) + 7).Height).Select
            Selection.Characters.Text = ArrTypes(a)
            Selection.Name = "Display" & a
            Selection.OnAction = "'Results_Summary.xlsm'!Display" & a & "_Click"
    Next a

    ActiveSheet.GroupBoxes.Add(Cells(20, (2 * UBound(ArrRow8)) + 7).Left, Cells(20, (2 * UBound(ArrRow8)) + 7).Top, Range(Cells(20, (2 * UBound(ArrRow8)) + 7), Cells(UBound(ArrRow9) + 24, (2 * UBound(ArrRow8)) + 7)).Width, Range(Cells(20, (2 * UBound(ArrRow8)) + 7), Cells(UBound(ArrRow9) + 24, (2 * UBound(ArrRow8)) + 7)).Height - 6).Select
        Selection.Characters.Text = Row9
        Selection.Name = Row9
        ActiveSheet.OptionButtons.Add(Cells(21, (2 * UBound(ArrRow8)) + 7).Left + 3, Cells(21, (2 * UBound(ArrRow8)) + 7).Top, Range(Cells(21, (2 * UBound(ArrRow8)) + 7), Cells(21, (2 * UBound(ArrRow8)) + 7)).Width - 6, Cells(21, (2 * UBound(ArrRow8)) + 7).Height).Select
            Selection.Characters.Text = "Ave."
            Selection.Name = "FirstAve"
            Selection.OnAction = "'Results_Summary.xlsm'!FirstAve_Click"
        ActiveSheet.OptionButtons.Add(Cells(22, (2 * UBound(ArrRow8)) + 7).Left + 3, Cells(22, (2 * UBound(ArrRow8)) + 7).Top, Range(Cells(22, (2 * UBound(ArrRow8)) + 7), Cells(22, (2 * UBound(ArrRow8)) + 7)).Width - 6, Cells(22, (2 * UBound(ArrRow8)) + 7).Height).Select
            Selection.Characters.Text = "All"
            Selection.Name = "FirstAll"
            Selection.OnAction = "'Results_Summary.xlsm'!FirstAll_Click"
    For a = 0 To UBound(ArrRow9)
        ActiveSheet.OptionButtons.Add(Cells(a + 23, (2 * UBound(ArrRow8)) + 7).Left + 3, Cells(a + 23, (2 * UBound(ArrRow8)) + 7).Top, Range(Cells(a + 23, (2 * UBound(ArrRow8)) + 7), Cells(a + 23, (2 * UBound(ArrRow8)) + 7)).Width - 6, Cells(a + 23, (2 * UBound(ArrRow8)) + 7).Height).Select
            Selection.Characters.Text = ArrRow9(a)
            Selection.Name = "First" & a
            Selection.OnAction = "'Results_Summary.xlsm'!First" & a & "_Click"
    Next a
    ActiveSheet.GroupBoxes.Add(Cells(20, (2 * UBound(ArrRow8)) + 9).Left, Cells(20, (2 * UBound(ArrRow8)) + 9).Top, Range(Cells(20, (2 * UBound(ArrRow8)) + 9), Cells(UBound(ArrRow8) + 24, (2 * UBound(ArrRow8)) + 9)).Width, Range(Cells(20, (2 * UBound(ArrRow8)) + 9), Cells(UBound(ArrRow8) + 24, (2 * UBound(ArrRow8)) + 9)).Height - 6).Select
        Selection.Characters.Text = Row8
        Selection.Name = Row8
        ActiveSheet.OptionButtons.Add(Cells(21, (2 * UBound(ArrRow8)) + 9).Left + 3, Cells(21, (2 * UBound(ArrRow8)) + 9).Top, Range(Cells(21, (2 * UBound(ArrRow8)) + 9), Cells(21, (2 * UBound(ArrRow8)) + 9)).Width - 6, Cells(21, (2 * UBound(ArrRow8)) + 9).Height).Select
            Selection.Characters.Text = "Ave."
            Selection.Name = "SecondAve"
            Selection.OnAction = "'Results_Summary.xlsm'!SecondAve_Click"
        ActiveSheet.OptionButtons.Add(Cells(22, (2 * UBound(ArrRow8)) + 9).Left + 3, Cells(22, (2 * UBound(ArrRow8)) + 9).Top, Range(Cells(22, (2 * UBound(ArrRow8)) + 9), Cells(22, (2 * UBound(ArrRow8)) + 9)).Width - 6, Cells(22, (2 * UBound(ArrRow8)) + 9).Height).Select
            Selection.Characters.Text = "All"
            Selection.Name = "SecondAll"
            Selection.OnAction = "'Results_Summary.xlsm'!SecondAll_Click"
    For a = 0 To UBound(ArrRow8)
        ActiveSheet.OptionButtons.Add(Cells(a + 23, (2 * UBound(ArrRow8)) + 9).Left + 3, Cells(a + 23, (2 * UBound(ArrRow8)) + 9).Top, Range(Cells(a + 23, (2 * UBound(ArrRow8)) + 9), Cells(a + 23, (2 * UBound(ArrRow8)) + 9)).Width - 6, Cells(a + 23, (2 * UBound(ArrRow8)) + 9).Height).Select
            Selection.Characters.Text = ArrRow8(a)
            Selection.Name = "Second" & a
            Selection.OnAction = "'Results_Summary.xlsm'!Second" & a & "_Click"
    Next a

    Sheets("Data Display").Move Before:=Sheets(1)
    Cells(1, 1).Value = (-2)
    Cells(2, 1).Value = (-2)
    Run "Results_Summary.xlsm!Display0_Click"
    
    Range("A1:A2").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    
    Application.ScreenUpdating = True

End Sub
Private Function Calc_Stats(ByRef ZCount(), Unit(), SD())
    Dim k As Integer, q As Integer, Mean As Double, StDev As Double, Uncirt As Double, Tot_Count As Double
    Dim Res(2)
    k = UBound(Unit)
    Mean = 0
    StDev = 0
    If Unit(0) = Empty Then
        Mean = 0
        StDev = 0
        Uncirt = 0
    Else
        For q = 0 To k
            Mean = Mean + (Unit(q) * ZCount(q))
            Tot_Count = Tot_Count + ZCount(q)
        Next q
       Mean = Mean / Tot_Count
                        
        For q = 0 To k
            StDev = StDev + (ZCount(q) * (SD(q) + (Unit(q) - Mean) ^ 2))
        Next q
        StDev = Sqr(StDev / Tot_Count)
        Uncirt = StDev / Sqr(Tot_Count)
    End If
    
    Res(0) = Mean
    Res(1) = StDev
    Res(2) = Uncirt
    
    Calc_Stats = Res
End Function


Sub Populate_Front_page()
Attribute Populate_Front_page.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Populate_Front_page Macro
'

    Dim i As Integer, fPath As String, fName As String, Location As String, ImageName As String
    Dim lColumn As Integer, lRow As Integer, ThisWS As Worksheet, ThisWB As Worksheet
    Dim ImageType As String, Title1 As String, Title2 As String, Title3 As String
    Dim Prefix1 As String, Prefix2 As String, Prefix3 As String
    Dim Sufix1 As String, Sufix2 As String, Sufix3 As String
    Dim ROW1 As Boolean, ROW2 As Boolean, ROW3 As Boolean
    Dim FileSeporator As String
    
            UserForm1.Show
            
            FileSeporator = UserForm1.txtFileSeporator.Text
            ROW1 = UserForm1.Row_1.Value
            ROW2 = UserForm1.Row_2.Value
            ROW3 = UserForm1.Row_3.Value
            
            If UserForm1.Filled.Value = True Then
               ImageType = "Filled"
            ElseIf UserForm1.Ring.Value = True Then
               ImageType = "Ring"
            Else
                MsgBox "Please Select Shape Type", vbOKOnly
            End If
            
            If ROW1 = True Then
                Title1 = UserForm1.txtTitle1.Text
                Prefix1 = UserForm1.txtPrefix1.Text
                Sufix1 = UserForm1.txtSufix1.Text
            Else
                Title1 = ""
                Prefix1 = ""
                Sufix1 = ""
            End If
            
            If ROW2 = True Then
                Title2 = UserForm1.txtTitle2.Text
                Prefix2 = UserForm1.txtPrefix2.Text
                Sufix2 = UserForm1.txtSufix2.Text
            Else
                Title2 = ""
                Prefix2 = ""
                Sufix2 = ""
            End If
            
            If ROW3 = True Then
                Title3 = UserForm1.txtTitle3.Text
                Prefix3 = UserForm1.txtPrefix3.Text
                Sufix3 = UserForm1.txtSufix3.Text
            Else
                Title3 = ""
                Prefix3 = ""
                Sufix3 = ""
            End If
            
            Unload UserForm1


    Set ThisWS = ActiveSheet
    
    lColumn = ActiveSheet.Cells(2, Columns.Count).End(xlToLeft).Column
      
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Front Page"
    Range("B2").FormulaR1C1 = "=A!R[-1]C[-1]"
    Range("B2").AutoFill Destination:=Range("B2:B3"), Type:=xlFillDefault
    Range("B3").AutoFill Destination:=Range("B3:T3"), Type:=xlFillDefault
    Range("B4").FormulaR1C1 = "=MAX(A!R[-1]C[-1]:R[10000]C[-1])"
    Range("C4").FormulaR1C1 = "=AVERAGE(A!R[-1]C[-1]:R[10000]C[-1])"
    Range("C5").FormulaR1C1 = "=STDEV.S(A!R[-2]C[-1]:R[10000]C[-1])"
    Range("C6").FormulaR1C1 = "=R[-2]C/SQRT(COUNTA(A!R[-3]C[-1]:R[10000]C[-1]))"
    Range("C7").FormulaR1C1 = "=MAX(A!R[-5]C[-1]:R[10000]C[-1])"
    Range("C8").FormulaR1C1 = "=MIN(A!R[-6]C[-1]:R[10000]C[-1])"
    Range("C4:C8").NumberFormat = "0.00"
    Range("C4:C8").AutoFill Destination:=Range("C4:S8"), Type:=xlFillDefault
    Range("A4").Value = "Average"
    Range("A5").Value = "st.Dev"
    Range("A6").Value = "+/-"
    Range("A7").Value = "Max"
    Range("A8").Value = "Min"
        
    Range("A10").FormulaR1C1 = Title1
    Range("A11").FormulaR1C1 = Title2
    Range("A12").FormulaR1C1 = Title3
    Range("A14").FormulaR1C1 = "Global Data"
    Range("A17").FormulaR1C1 = "Crust Data"
    If ImageType = "Ring" Then
        Range("A20").FormulaR1C1 = "Hole Data"
    End If
    
    ActiveWindow.FreezePanes = False
    Cells(3, 2).Select
    ActiveWindow.FreezePanes = True
    
    Range("B2:T8").AutoFill Destination:=Range(Cells(2, 2), Cells(8, lColumn)), Type:=xlFillDefault
    
    Range("A23").RowHeight = 409
    Range("A24").RowHeight = 409
    Range("A25").RowHeight = 409
    Range("A26").RowHeight = 409
    
    i = 2
    fPath = Application.ThisWorkbook.Path & FileSeporator
    fName = ThisWorkbook.Name
    
    Do Until i > lColumn
            ImageName = Worksheets(2).Cells(2, i).Value
            ImageName = Left(ImageName, Len(ImageName) - 4)
            Cells(2, i).Select
            
            If InStr(1, Cells(2, i), "control", vbTextCompare) > 0 Then
                If ROW1 = True Then
                    Cells(10, i).Value = "0"
                End If
                If ROW2 = True Then
                    Cells(11, i).Value = "0"
                End If
                If ROW3 = True Then
                    Cells(12, i).Value = "0"
                End If
            Else
                Dim OpenPos As Long, ClosePos As Long
                OpenPos = 1
                ClosePos = 1
                If ROW1 = True Then
                    OpenPos = InStr(ClosePos, Cells(2, i).Value, Prefix1, vbTextCompare)
                    ClosePos = InStr(OpenPos + 1, Cells(2, i).Value, Sufix1, vbTextCompare)
                    Cells(10, i).Value = Mid(Cells(2, i).Value, OpenPos + 1, ClosePos - OpenPos - 1)
                    If InStr(1, Cells(10, i).Value, 0) > 0 And InStr(1, Cells(10, i).Value, 5) > 0 Then
                    Cells(10, i).Value = 0.5
                    End If
                End If
                If ROW2 = True Then
                    OpenPos = InStr(ClosePos, Cells(2, i).Value, Prefix2, vbTextCompare)
                    ClosePos = InStr(OpenPos + 1, Cells(2, i).Value, Sufix2, vbTextCompare)
                    Cells(11, i).Value = Mid(Cells(2, i).Value, OpenPos + 1, ClosePos - OpenPos - 1)
                    If InStr(1, Cells(11, i).Value, "0") > 0 And InStr(1, Cells(11, i).Value, "5") > 0 Then
                        Cells(11, i).Value = 0.5
                    End If
                End If
                If ROW3 = True Then
                    OpenPos = InStr(ClosePos, Cells(2, i).Value, Prefix3, vbTextCompare)
                    ClosePos = InStr(OpenPos + 1, Cells(2, i).Value, Sufix3, vbTextCompare)
                    Cells(12, i).Value = Mid(Cells(2, i).Value, OpenPos + 1, ClosePos - OpenPos - 1)
                    If InStr(1, Cells(12, i).Value, 0) > 0 And InStr(1, Cells(12, i).Value, 5) > 0 Then
                        Cells(12, i).Value = 0.5
                    End If
                End If
            End If

        
        Application.ScreenUpdating = False
        
        Location = fPath & ImageName & FileSeporator & ImageName & "_Global-Measurements"
        Workbooks.Open (Location & ".csv")
            ActiveSheet.Range("B1:P2").Select
            Selection.Copy
            Windows(fName).Activate
            Cells(14, i).Select
            ActiveSheet.Paste
            Windows(ImageName & "_Global-Measurements.csv").Close
        
        Location = fPath & ImageName & FileSeporator & ImageName & "_Crust-Summary"
        Workbooks.Open (Location & ".csv")
            Range("C1", "H1").Select
            Selection.Copy
            Windows(fName).Activate
            Cells(17, i).Select
            ActiveSheet.Paste
            Windows(ImageName & "_Crust-Summary.csv").Activate
                lRow = (ActiveSheet.Cells(20, 2).SpecialCells(xlCellTypeLastCell).Row) - 3
            If lRow < 0 Then
                lRow = 2
            End If
            Range(Cells(lRow, 3), Cells(lRow, 8)).Select
            Selection.Copy
            Windows(fName).Activate
            Cells(18, i).Select
            ActiveSheet.Paste
            Windows(ImageName & "_Crust-Summary.csv").Close
        
        If ImageType = "Ring" Then
            Location = fPath & ImageName & FileSeporator & ImageName & "_Ring-Hole-Measurments"
            Workbooks.Open (Location & ".csv")
                Range("B1", "R1").Select
                Selection.Copy
                Windows(fName).Activate
                Cells(20, i).Select
                ActiveSheet.Paste
                Windows(ImageName & "_Ring-Hole-Measurments.csv").Activate
                Range("B3", "R3").Select
                Selection.Copy
                Windows(fName).Activate
                Cells(21, i).Select
                ActiveSheet.Paste
                Windows(ImageName & "_Ring-Hole-Measurments.csv").Close
        End If
        
            Range(Cells(23, i), Cells(23, (i + 4))).Merge
            Range(Cells(23, (i + 5)), Cells(23, (i + 9))).Merge
            Range(Cells(23, (i + 10)), Cells(23, (i + 14))).Merge
            Range(Cells(24, i), Cells(26, (i + 3))).Merge Across:=True
            Range(Cells(24, (i + 4)), Cells(26, (i + 10))).Merge Across:=True
            Range(Cells(24, (i + 11)), Cells(26, (i + 17))).Merge Across:=True
            
            Location = fPath & ImageName & FileSeporator & ImageName & ".tif"
            ActiveSheet.Shapes.AddPicture(Location, _
            linktofile:=False, savewithdocument:=True, Left:=Cells(23, i).Left + 5, Top:=Cells(23, i).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 40 * 5
            Selection.Placement = xlMoveAndSize

            ActiveSheet.Shapes.AddPicture(fPath & ImageName & FileSeporator & ImageName & "_Thresholded.tif", _
            linktofile:=False, savewithdocument:=True, Left:=Cells(23, i + 5).Left + 5, Top:=Cells(23, i + 9).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 40 * 5
            Selection.Placement = xlMoveAndSize
            
            ActiveSheet.Shapes.AddPicture(fPath & ImageName & FileSeporator & ImageName & "_Porosity-Map.tif", _
            linktofile:=False, savewithdocument:=True, Left:=Cells(23, i + 10).Left + 5, Top:=Cells(23, i + 14).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 40 * 5
            Selection.Placement = xlMoveAndSize
            
            ActiveSheet.Shapes.AddPicture(fPath & ImageName & FileSeporator & ImageName & "_Area-Distribution.jpg", _
            linktofile:=False, savewithdocument:=True, Left:=Cells(24, i).Left + 5, Top:=Cells(24, i + 3).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 40 * 4
            Selection.Placement = xlMoveAndSize
            
            ActiveSheet.Shapes.AddPicture(fPath & ImageName & FileSeporator & ImageName & "_Area-Map.tif", _
            linktofile:=False, savewithdocument:=True, Left:=Cells(24, i + 4).Left + 5, Top:=Cells(24, i + 10).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 35 * 7
            Selection.Placement = xlMoveAndSize
            
            ActiveSheet.Shapes.AddPicture(fPath & ImageName & FileSeporator & ImageName & "_Porosity-Map-Contours.tif", _
            linktofile:=False, savewithdocument:=True, Left:=Cells(24, i + 11).Left + 5, Top:=Cells(24, i + 17).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 35 * 7
            Selection.Placement = xlMoveAndSize
                
            ActiveSheet.Shapes.AddPicture(fPath & ImageName & FileSeporator & ImageName & "_Circularity-Distribution.jpg", _
            linktofile:=False, savewithdocument:=True, Left:=Cells(25, i).Left + 5, Top:=Cells(25, i + 3).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 40 * 4
            Selection.Placement = xlMoveAndSize
            
            ActiveSheet.Shapes.AddPicture(fPath & ImageName & FileSeporator & ImageName & "_Circularity-Map.tif", _
            linktofile:=False, savewithdocument:=True, Left:=Cells(25, i + 4).Left + 5, Top:=Cells(25, i + 10).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 35 * 7
            Selection.Placement = xlMoveAndSize
            
            ActiveSheet.Shapes.AddPicture(fPath & ImageName & FileSeporator & ImageName & "_Granulometry.jpg", _
            linktofile:=False, savewithdocument:=True, Left:=Cells(25, i + 11).Left + 5, Top:=Cells(25, i + 17).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 35 * 7
            Selection.Placement = xlMoveAndSize
            
            ActiveSheet.Shapes.AddPicture(fPath & ImageName & FileSeporator & ImageName & "_Feret-Length-Distribution.jpg", _
            linktofile:=False, savewithdocument:=True, Left:=Cells(26, i).Left + 5, Top:=Cells(26, i + 3).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 40 * 4
            Selection.Placement = xlMoveAndSize
            
            ActiveSheet.Shapes.AddPicture(fPath & ImageName & FileSeporator & ImageName & "_Feret-Length-Map.tif", _
            linktofile:=False, savewithdocument:=True, Left:=Cells(26, i + 4).Left + 5, Top:=Cells(26, i + 10).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 35 * 7
            Selection.Placement = xlMoveAndSize
            
            ActiveSheet.Shapes.AddPicture(fPath & ImageName & FileSeporator & ImageName & "_Feret-Angle-Map.tif", _
            linktofile:=False, savewithdocument:=True, Left:=Cells(26, i + 11).Left + 5, Top:=Cells(26, i + 17).Top + 5, Width:=-1, Height:=-1).Select
            Selection.ShapeRange.Width = 35 * 7
            Selection.Placement = xlMoveAndSize
            
        Application.ScreenUpdating = True
        
            Application.CutCopyMode = False
            ActiveWorkbook.Save
            i = i + 19
     Loop
    
    
    Range("B1").Select

End Sub
