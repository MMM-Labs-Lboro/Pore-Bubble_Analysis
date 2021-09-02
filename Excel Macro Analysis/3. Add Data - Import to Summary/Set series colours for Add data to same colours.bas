Attribute VB_Name = "Module1"
Sub set_colours()
Dim a, b, c, d, f, g, h, j

b = ActiveChart.FullSeriesCollection.Count
c = Application.InputBox("Number to start with? (5 - 12)", "Set Start", 6, Type:=1)
Do While c > 12
c = Application.InputBox("uh. Please, less than 12.", "Set Start", 6, Type:=1)
Loop
d = Application.InputBox("How many Series? (1 - " & (12 - c) & ")", "Set No Colours", 4, Type:=1)
Do While (c + d) > 12
d = Application.InputBox("uh. Please, renge must be between 1 and " & (12 - c), "Set No Colours", 4, Type:=1)
Loop
g = Application.InputBox("Avoid a colour? (between " & c & " and 12)", "Avoid Colour", 7, Type:=1)
Do While g > 12 Or g < c
g = Application.InputBox("uh. Please, avoid must be between " & c & " and 12)", "Avoid Colour", 7, Type:=1)
Loop
h = Application.InputBox("Avoid another colour? (between " & c & " and 12)", "Avoid Colour", 11, Type:=1)
Do While h > 12 Or h < c
h = Application.InputBox("uh. Please, avoid must be between " & c & " and 12)", "Avoid Colour", 11, Type:=1)
Loop

f = c
j = 0
If (g < (c + d)) And (g > c) Then
    j = j + 1
End If
If (h <> g) And (g < (c + d)) And (g > c) Then
    j = j + 1
End If

For a = 1 To b
    ActiveChart.FullSeriesCollection(a).Format.Line.ForeColor.ObjectThemeColor = msoThemeColorAccent & f
    ActiveChart.FullSeriesCollection(a).Format.Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent & f
    f = f + 1
    If f = g Then
     f = g + 1
    End If
    If f = h Then
     f = h + 1
    End If
    If f = c + d + j Then
     f = c
    End If
Next a

End Sub
