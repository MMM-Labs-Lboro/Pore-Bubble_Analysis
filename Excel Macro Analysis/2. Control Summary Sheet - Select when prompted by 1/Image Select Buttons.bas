Attribute VB_Name = "Module21"
Sub Change_Display(PicLoc, PicWidth, x)
Dim Repeat As String
Dim a, b, c As Integer
Dim First() 'Location Array
Dim Second() 'Power Array

' Build Axis arrays
ReDim First(0)
First(0) = 17
Worksheets(PicLoc).Activate
Range(Cells(1, 1), Cells(30, 1)).Select
For Each cell In Selection
     Repeat = ""
     If (cell <> "") Then
         For a = 0 To UBound(First)
             If cell = First(a) Then
                 Repeat = Repeat & "X"
             Else
                 Repeat = Repeat & "_"
             End If
         Next a
         If InStr(Repeat, "X") = 0 Then
            If First(0) = 17 Then
                First(0) = cell
            Else
                ReDim Preserve First(UBound(First) + 1)
                First(UBound(First)) = cell
            End If
         End If
     End If
 Next cell
ReDim Second(0)
Second(0) = 17
Range(Cells(1, 1), Cells(1, 40)).Select
For Each cell In Selection
     Repeat = ""
     If (cell <> "") Then
         For a = 0 To UBound(Second)
             If cell = Second(a) Then
                 Repeat = Repeat & "X"
             Else
                 Repeat = Repeat & "_"
             End If
         Next a
         If InStr(Repeat, "X") = 0 Then
            If Second(0) = 17 Then
                Second(0) = cell
            Else
                ReDim Preserve Second(UBound(Second) + 1)
                Second(UBound(Second)) = cell
            End If
         End If
     End If
 Next cell
Worksheets("Data Display").Activate

' Delete old info
ActiveSheet.ChartObjects("First Chart").Activate
c = ActiveChart.FullSeriesCollection.Count
For a = 1 To c
    ActiveChart.FullSeriesCollection(a).IsFiltered = True
Next a
ActiveSheet.ChartObjects("Second Chart").Activate
c = ActiveChart.FullSeriesCollection.Count
For a = 1 To c
    ActiveChart.FullSeriesCollection(a).IsFiltered = True
Next a
For Each Picture In ActiveSheet.Pictures
    If Not Intersect(Picture.TopLeftCell, Range(Cells(18, 1), Cells(40, (UBound(First) * 2) + 5))) Is Nothing Then
        Picture.Delete
    End If
Next Picture
For Each Picture In ActiveSheet.Pictures
    If Not Intersect(Picture.TopLeftCell, Range(Cells(18, (UBound(First) * 2) + 11), Cells(40, 100))) Is Nothing Then
        Picture.Delete
    End If
Next Picture

' Add New Data
' Control Info (in all)
If Second(0) = 0 Then
    Worksheets(PicLoc).Activate
    For Each Picture In ActiveSheet.Pictures
        Picture.Select
        If Not Intersect(Picture.TopLeftCell, Cells(2, 2)) Is Nothing Then
            Picture.Select
            Selection.Copy
            Worksheets("Data Display").Activate
            Cells(19, 3).Select
            ActiveSheet.Paste
            Selection.ShapeRange.LockAspectRatio = msoTrue
            Selection.ShapeRange.Width = PicWidth
            Cells(19, (2 * UBound(First)) + 12).Select
            ActiveSheet.Paste
            Selection.ShapeRange.LockAspectRatio = msoTrue
            Selection.ShapeRange.Width = PicWidth
            Worksheets(PicLoc).Activate
        End If
    Next Picture
End If

' FIRST
Worksheets("Data Display").Activate
ActiveSheet.ChartObjects("First Chart").Activate
If Cells(1, 1).Value < 1 Then
    If Cells(1, 1).Value = (-3) Then 'Compare Salts
        ActiveChart.FullSeriesCollection(220 + x).IsFiltered = False
        ActiveChart.FullSeriesCollection(231 + x).IsFiltered = False
        If x = 2 Then
            ActiveChart.FullSeriesCollection(243).IsFiltered = False
            ActiveChart.FullSeriesCollection(247).IsFiltered = False
        End If
        If x = 5 Then
            ActiveChart.FullSeriesCollection(244).IsFiltered = False
            ActiveChart.FullSeriesCollection(248).IsFiltered = False
        End If
        If x = 6 Then
            ActiveChart.FullSeriesCollection(245).IsFiltered = False
            ActiveChart.FullSeriesCollection(249).IsFiltered = False
        End If
        If x = 7 Then
            ActiveChart.FullSeriesCollection(246).IsFiltered = False
            ActiveChart.FullSeriesCollection(250).IsFiltered = False
        End If
    End If
    If Cells(1, 1).Value = (-2) Then 'Ave
        ActiveChart.FullSeriesCollection((2 * x) - 1).IsFiltered = False
        ActiveChart.FullSeriesCollection(2 * x).IsFiltered = False
    End If
    If Cells(1, 1).Value = (-1) Then 'All
        For a = 1 To (UBound(Second) + 1)
            ActiveChart.FullSeriesCollection((22 * a) + (2 * x) - 1).IsFiltered = False
            ActiveChart.FullSeriesCollection((22 * a) + (2 * x)).IsFiltered = False
        Next a
    End If
    If Cells(1, 1).Value = 0 Then 'Control
        ActiveChart.FullSeriesCollection(21 + (2 * x)).IsFiltered = False
        ActiveChart.FullSeriesCollection(22 + (2 * x)).IsFiltered = False
    End If
Else
    ' inc. Control
    ActiveChart.FullSeriesCollection(21 + (2 * x)).IsFiltered = False
    ActiveChart.FullSeriesCollection(22 + (2 * x)).IsFiltered = False
    ' Rest
    a = Cells(1, 1).Value
    ActiveChart.FullSeriesCollection((22 * (a + 1)) + (2 * x) - 1).IsFiltered = False
    ActiveChart.FullSeriesCollection((22 * (a + 1)) + (2 * x)).IsFiltered = False
End If

Worksheets(PicLoc).Activate
For a = 0 To UBound(First)
If Worksheets("Data Display").Cells(1, 1).Value < 0 Then
    If Second(0) = 0 Then
        b = 1
    Else
        b = 0
    End If
    Else: b = Worksheets("Data Display").Cells(1, 1).Value
End If
For Each Picture In ActiveSheet.Pictures
    If Not Intersect(Picture.TopLeftCell, Cells(a + 2, (7 * b) + 2)) Is Nothing Then
        Picture.Select
        Selection.Copy
        Worksheets("Data Display").Activate
        Cells(19, 3 + (2 * a)).Select
        ActiveSheet.Paste
        Selection.ShapeRange.LockAspectRatio = msoTrue
        Selection.ShapeRange.Width = PicWidth
        Worksheets(PicLoc).Activate
    End If
Next Picture
Next a

' SECOND
Worksheets("Data Display").Activate
ActiveSheet.ChartObjects("Second Chart").Activate
If Cells(2, 1).Value < 1 Then
    If Cells(2, 1).Value = (-2) Then 'Ave
        ActiveChart.FullSeriesCollection((2 * x) - 1).IsFiltered = False
        ActiveChart.FullSeriesCollection(2 * x).IsFiltered = False
    End If
    If Cells(2, 1).Value = (-1) Then 'All
        For a = 1 To (UBound(First) + 1)
            ActiveChart.FullSeriesCollection((22 * a) + (2 * x) - 1).IsFiltered = False
            ActiveChart.FullSeriesCollection((22 * a) + (2 * x)).IsFiltered = False
        Next a
    End If
    If Cells(2, 1).Value = 0 Then 'Control
        ActiveChart.FullSeriesCollection(21 + (2 * x)).IsFiltered = False
        ActiveChart.FullSeriesCollection(22 + (2 * x)).IsFiltered = False
    End If
Else
    ' inc. Control
    ActiveChart.FullSeriesCollection(21 + (2 * x)).IsFiltered = False
    ActiveChart.FullSeriesCollection(22 + (2 * x)).IsFiltered = False
    ' Rest
    a = Cells(2, 1).Value
    ActiveChart.FullSeriesCollection((22 * (a + 1)) + (2 * x) - 1).IsFiltered = False
    ActiveChart.FullSeriesCollection((22 * (a + 1)) + (2 * x)).IsFiltered = False
End If

Worksheets(PicLoc).Activate
For b = 0 To UBound(Second)
If Worksheets("Data Display").Cells(2, 1).Value < 0 Then
    If Second(0) = 0 Then
        a = 1
    Else
        a = 0
    End If
    Else: a = Worksheets("Data Display").Cells(2, 1).Value
End If
For Each Picture In ActiveSheet.Pictures
    If Not Intersect(Picture.TopLeftCell, Cells(a + 2, (7 * b) + 2)) Is Nothing Then
        Picture.Select
        Selection.Copy
        Worksheets("Data Display").Activate
        Cells(19, (2 * UBound(First)) + 12 + (b * 2)).Select
        ActiveSheet.Paste
        Selection.ShapeRange.LockAspectRatio = msoTrue
        Selection.ShapeRange.Width = PicWidth
        Worksheets(PicLoc).Activate
    End If
Next Picture
Next b

Worksheets("Data Display").Activate
Cells(15, 15).Select
End Sub
Sub Display0_Click()
Application.ScreenUpdating = False

Dim PicLoc As String
Dim PicWidth, x As Integer

PicLoc = "Picture_Colour"
PicWidth = 90
x = 1

Call Change_Display(PicLoc, PicWidth, x)
Application.ScreenUpdating = True
End Sub
Sub Display1_Click()
Application.ScreenUpdating = False

Dim PicLoc As String
Dim PicWidth, x As Integer

PicLoc = "Picture_Area"
PicWidth = 90
x = 2

Call Change_Display(PicLoc, PicWidth, x)
Application.ScreenUpdating = True
End Sub
Sub Display2_Click()
Application.ScreenUpdating = False

Dim PicLoc As String
Dim PicWidth, x As Integer

PicLoc = "Picture_Area_Dis"
PicWidth = 90
x = 3

Call Change_Display(PicLoc, PicWidth, x)
Application.ScreenUpdating = True
End Sub
Sub Display3_Click()
Application.ScreenUpdating = False

Dim PicLoc As String
Dim PicWidth, x As Integer

PicLoc = "Picture_Area_Dis"
PicWidth = 90
x = 4

Call Change_Display(PicLoc, PicWidth, x)
Application.ScreenUpdating = True
End Sub
Sub Display4_Click()
Application.ScreenUpdating = False

Dim PicLoc As String
Dim PicWidth, x As Integer

PicLoc = "Picture_Circ"
PicWidth = 90
x = 5

Call Change_Display(PicLoc, PicWidth, x)
Application.ScreenUpdating = True
End Sub
Sub Display5_Click()
Application.ScreenUpdating = False

Dim PicLoc As String
Dim PicWidth, x As Integer

PicLoc = "Picture_Feret_Len"
PicWidth = 90
x = 6

Call Change_Display(PicLoc, PicWidth, x)
Application.ScreenUpdating = True
End Sub
Sub Display6_Click()
Application.ScreenUpdating = False

Dim PicLoc As String
Dim PicWidth, x As Integer
PicLoc = "Picture_Feret_Angle"
PicWidth = 90
x = 7

Call Change_Display(PicLoc, PicWidth, x)
Application.ScreenUpdating = True
End Sub
Sub Display7_Click()
Application.ScreenUpdating = False

Dim PicLoc As String
Dim PicWidth, x As Integer

PicLoc = "Picture_Porosity_Map"
PicWidth = 90
x = 8

Call Change_Display(PicLoc, PicWidth, x)
Application.ScreenUpdating = True
End Sub
Sub Display8_Click()
Application.ScreenUpdating = False

Dim PicLoc As String
Dim PicWidth, x As Integer
PicLoc = "Picture_Colour"
PicWidth = 90
x = 9

Call Change_Display(PicLoc, PicWidth, x)
Application.ScreenUpdating = True
End Sub
Sub Display9_Click()
Application.ScreenUpdating = False

Dim PicLoc As String
Dim PicWidth, x As Integer

PicLoc = "Picture_Colour"
PicWidth = 90
x = 10

Call Change_Display(PicLoc, PicWidth, x)
Application.ScreenUpdating = True
End Sub
Sub Display10_Click()
Application.ScreenUpdating = False

Dim PicLoc As String
Dim PicWidth, x As Integer
PicLoc = "Picture_Colour"
PicWidth = 90
x = 11

Call Change_Display(PicLoc, PicWidth, x)
Application.ScreenUpdating = True
End Sub
Sub Change_Display1(x)
Dim a, b, c, d, minseries, maxseries As Integer

' Delete old info
ActiveSheet.ChartObjects("Third Chart").Activate
c = ActiveChart.FullSeriesCollection.Count
For a = 1 To c
    ActiveChart.FullSeriesCollection(a).IsFiltered = True
Next a
ActiveSheet.ChartObjects("Fourth Chart").Activate
c = ActiveChart.FullSeriesCollection.Count
For a = 1 To c
    ActiveChart.FullSeriesCollection(a).IsFiltered = True
Next a

d = Cells(9 + x, 1).Value

' Third
b = Cells(9, 1).Value 'No in Second
maxseries = 0
For c = 1 To x
maxseries = maxseries + Cells(9 + c, 1).Value
Next c
maxseries = maxseries * (b + 1)
minseries = maxseries + 1 - (d * (b + 1))
ActiveSheet.ChartObjects("Third Chart").Activate
If Cells(1, 1).Value < 1 Then
    If Cells(1, 1).Value = (-2) Then
        For c = (maxseries - d + 1) To maxseries
            ActiveChart.FullSeriesCollection(c).IsFiltered = False
        Next c
    End If
    If Cells(1, 1).Value = (-1) Then
        For c = minseries To (maxseries - d)
            ActiveChart.FullSeriesCollection(c).IsFiltered = False
        Next c
    End If
    If Cells(1, 1).Value = 0 Then
        For c = minseries To (minseries + d - 1)
            ActiveChart.FullSeriesCollection(c).IsFiltered = False
        Next c
    End If
Else
    ' inc. Conrtol
    For c = minseries To (minseries + d - 1)
        ActiveChart.FullSeriesCollection(c).IsFiltered = False
    Next c
    ' Rest
    a = Cells(1, 1).Value
    For c = minseries + (a * d) To minseries + ((a + 1) * d) - 1
        ActiveChart.FullSeriesCollection(c).IsFiltered = False
    Next c
End If

' Fourth
b = Cells(8, 1).Value 'No in First
maxseries = 0
For c = 1 To x
maxseries = maxseries + Cells(9 + c, 1).Value
Next c
maxseries = maxseries * (b + 1)
minseries = maxseries + 1 - (d * (b + 1))
ActiveSheet.ChartObjects("Fourth Chart").Activate
If Cells(2, 1).Value < 1 Then
    If Cells(2, 1).Value = (-2) Then
        For c = (maxseries - d + 1) To maxseries
            ActiveChart.FullSeriesCollection(c).IsFiltered = False
        Next c
    End If
    If Cells(2, 1).Value = (-1) Then
        For c = minseries To (maxseries - d)
            ActiveChart.FullSeriesCollection(c).IsFiltered = False
        Next c
    End If
    If Cells(2, 1).Value = 0 Then
        For c = minseries To (minseries + d - 1)
            ActiveChart.FullSeriesCollection(c).IsFiltered = False
        Next c
    End If
Else
    ' inc. Conrtol
    For c = minseries To (minseries + d - 1)
        ActiveChart.FullSeriesCollection(c).IsFiltered = False
    Next c
    ' Rest
    a = Cells(2, 1).Value
    For c = minseries + (a * d) To minseries + ((a + 1) * d) - 1
        ActiveChart.FullSeriesCollection(c).IsFiltered = False
    Next c
End If
End Sub
Sub Additional0_Click()
Application.ScreenUpdating = False
Dim x As Integer
x = 1
Call Change_Display1(x)
Application.ScreenUpdating = True
End Sub
Sub Additional1_Click()
Application.ScreenUpdating = False
Dim x As Integer
x = 2
Call Change_Display1(x)
Application.ScreenUpdating = True
End Sub
Sub Additional2_Click()
Application.ScreenUpdating = False
Dim x As Integer
x = 3
Call Change_Display1(x)
Application.ScreenUpdating = True
End Sub
Sub Additional3_Click()
Application.ScreenUpdating = False
Dim x As Integer
x = 4
Call Change_Display1(x)
Application.ScreenUpdating = True
End Sub
Sub FirstAve_Click()
    Cells(1, 1).Value = (-2)
End Sub
Sub FirstAll_Click()
    Cells(1, 1).Value = (-1)
End Sub
Sub First0_Click()
    Cells(1, 1).Value = 0
End Sub
Sub First1_Click()
    Cells(1, 1).Value = 1
End Sub
Sub First2_Click()
    Cells(1, 1).Value = 2
End Sub
Sub First3_Click()
    Cells(1, 1).Value = 3
End Sub
Sub First4_Click()
    Cells(1, 1).Value = 4
End Sub
Sub First5_Click()
    Cells(1, 1).Value = 5
End Sub
Sub First6_Click()
    Cells(1, 1).Value = 6
End Sub
Sub First7_Click()
    Cells(1, 1).Value = 7
End Sub
Sub First8_Click()
    Cells(1, 1).Value = 8
End Sub
Sub First9_Click()
    Cells(1, 1).Value = 9
End Sub
Sub FirstSalt_Click()
    Cells(1, 1).Value = (-3)
End Sub
Sub SecondAve_Click()
    Cells(2, 1).Value = (-2)
End Sub
Sub SecondAll_Click()
    Cells(2, 1).Value = (-1)
End Sub
Sub Second0_Click()
    Cells(2, 1).Value = 0
End Sub
Sub Second1_Click()
    Cells(2, 1).Value = 1
End Sub
Sub Second2_Click()
    Cells(2, 1).Value = 2
End Sub
Sub Second3_Click()
    Cells(2, 1).Value = 3
End Sub
Sub Second4_Click()
    Cells(2, 1).Value = 4
End Sub
Sub Second5_Click()
    Cells(2, 1).Value = 5
End Sub
Sub Second6_Click()
    Cells(2, 1).Value = 6
End Sub
Sub Second7_Click()
    Cells(2, 1).Value = 7
End Sub
Sub Second8_Click()
    Cells(2, 1).Value = 8
End Sub
Sub Second9_Click()
    Cells(2, 1).Value = 9
End Sub
