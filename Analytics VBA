Analytics with Charts and MsgBoxes (Code)

Sub CheckCalendar()

    Dim Days As Integer
    Dim myCell As Range
    Dim myRange As Range
    Dim allCellsEmpty As Boolean
    Dim dayNames() As Variant
    Dim dayRanges() As Variant
    
    dayNames = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    dayRanges = Array("B2:B26", "C2:C26", "D2:D26", "E2:E26", "F2:F26", "G2:G26", "H2:H26")
    
    For Days = 0 To 6
        Set myRange = ThisWorkbook.Sheets("Weekly Calendar").Range(dayRanges(i))
        allCellsEmpty = True
        
        ' Check if all cells in the range are empty
        For Each myCell In myRange
            If Not IsEmpty(myCell.Value) Then
                allCellsEmpty = False
                Exit For
            End If
        Next myCell
        
        If allCellsEmpty = True Then
            If Days = 0 Or Days = 6 Then ' Sunday or Saturday
                MsgBox ("Your " & dayNames(Days) & " is pretty free. Try to stay productive, or rest and recover for this weekend")
            Else ' Weekdays
                MsgBox ("Your " & dayNames(Days) & " is pretty free. Try to add some tasks to stay productive!")
            End If
        End If
    Next Days
End Sub

Sub OpenLogReflectionForm()
    LogReflection.Show
End Sub

Sub OpenLogTaskForm()
    LogTaskTime.Show
End Sub

Sub CreateCategoryPieChart()

    Dim ws As Worksheet
    Dim analytics As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim countQuiz As Long
    Dim countTest As Long
    Dim countAssignment As Long
    Dim countExam As Long
    Dim countHomework As Long

    Set ws = ThisWorkbook.Sheets("Weekly Calendar")
    Set analytics = ThisWorkbook.Sheets("Analytics")

    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    For Each cell In ws.Range("P2:P" & lastRow)
        If cell.Value = "Quiz" Then countQuiz = countQuiz + 1
        If cell.Value = "Test" Then countTest = countTest + 1
        If cell.Value = "Assignment" Then countAssignment = countAssignment + 1
        If cell.Value = "Exam" Then countExam = countExam + 1
        If cell.Value = "Homework" Then countHomework = countHomework + 1
    Next cell

    analytics.Range("AA1:AB10").Clear
    analytics.Range("AA1").Value = "Category"
    analytics.Range("AB1").Value = "Count"
    
    analytics.Range("AA2").Value = "Quiz"
    analytics.Range("AB2").Value = countQuiz
    analytics.Range("AA3").Value = "Test"
    analytics.Range("AB3").Value = countTest
    analytics.Range("AA4").Value = "Assignment"
    analytics.Range("AB4").Value = countAssignment
    analytics.Range("AA5").Value = "Exam"
    analytics.Range("AB5").Value = countExam
    analytics.Range("AA6").Value = "Homework"
    analytics.Range("AB6").Value = countHomework

    Dim ch As ChartObject
    For Each ch In analytics.ChartObjects
        If ch.Name = "CategoryPieChart" Then
            ch.Delete
            Exit For
        End If
    Next ch

    Set ch = analytics.ChartObjects.Add(Left:=50, Top:=analytics.Rows("4").Top, Width:=400, Height:=300)

    ch.Name = "CategoryPieChart"
    
    With ch.Chart
        .ChartType = xlPie
        .SetSourceData analytics.Range("AA1:AB6")
        .HasTitle = True
        .ChartTitle.Text = "Task Category Breakdown"
        .ApplyDataLabels
    End With
    
    ws.PivotTables("PvWeeklyCount").RefreshTable
    
    analytics.Activate
    MsgBox ("Pie chart with breakdown of task categories updated.")
    
End Sub

Sub ClearWeeklyTasks()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Weekly Calendar")
    ws.Range("J3:Q100").ClearContents
    ws.Range("J3:Q26").Interior.Color = RGB(218, 233, 248)
End Sub

Sub ClearCalendar()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Weekly Calendar")
    ws.Range("B2:H26").ClearContents
    ws.Range("B2:H26").Interior.Color = RGB(218, 233, 248)
End Sub

Sub CreateSchedulePieChart()

    Dim ws As Worksheet
    Dim analytics As Worksheet
    Dim cell As Range
    Dim rng As Range

    Dim countStudy As Long
    Dim countSocial As Long
    Dim countPersonal As Long
    Dim countOther As Long

    Dim colSocial As Long
    Dim colStudy As Long
    Dim colPersonal As Long

    colSocial = RGB(255, 223, 186)
    colStudy = RGB(186, 255, 186)
    colPersonal = RGB(186, 186, 255)
    colOther = RGB(166, 201, 238)

    Set ws = ThisWorkbook.Sheets("Weekly Calendar")
    Set analytics = ThisWorkbook.Sheets("Analytics")

    Set rng = ws.Range("B2:H26")

    For Each cell In rng
        Select Case cell.Interior.Color
            Case colStudy: countStudy = countStudy + 1
            Case colSocial: countSocial = countSocial + 1
            Case colPersonal: countPersonal = countPersonal + 1
            Case colOther: countOther = countOther + 1
        End Select
    Next cell


    analytics.Range("AG1:AH10").Clear

    analytics.Range("AG1").Value = "Category"
    analytics.Range("AH1").Value = "Hours"

    analytics.Range("AG2").Value = "Study": analytics.Range("AH2").Value = countStudy
    analytics.Range("AG3").Value = "Social": analytics.Range("AH3").Value = countSocial
    analytics.Range("AG4").Value = "Personal": analytics.Range("AH4").Value = countPersonal
    analytics.Range("AG5").Value = "Other": analytics.Range("AH5").Value = countOther

    Dim ch As ChartObject
    For Each ch In analytics.ChartObjects
        If ch.Name = "SchedulePieChart" Then
            ch.Delete
            Exit For
        End If
    Next ch

    Set ch = analytics.ChartObjects.Add( _
        Left:=analytics.Columns("K").Left, _
        Top:=analytics.Rows("4").Top, _
        Width:=400, Height:=300)

    ch.Name = "SchedulePieChart"

    With ch.Chart
        .ChartType = xlPie
        .SetSourceData analytics.Range("AG1:AH5")
        .HasTitle = True
        .ChartTitle.Text = "Weekly Time Category Breakdown"
        .ApplyDataLabels
    End With

    analytics.Activate
    MsgBox "Schedule Pie Chart Updated!"


End Sub

Sub CheckInterleaving()

Dim ws As Worksheet
Dim analytics As Worksheet
Dim Row As Long
Dim cell As Range
Dim i As Long

For i = 3 To 24
    If Cells(1 + 1, 17).Value <> "" Then
        'Check Low High and Medium Interleaving
        If Cells(i, 17).Value = "Easy" Then
            If Cells(i + 1, 17).Value = "Easy" Then
                MsgBox "Please Split up Easy Tasks with Hard or Medium Tasks"
                End
            End If
        
            
        ElseIf Cells(i, 17).Value = "Medium" Then
            If Cells(i + 1, 17).Value = "Medium" Then
                
                If Cells(i + 2, 17).Value = "" Then
                    MsgBox " Your Interleaving looks good!"
                    End
                ElseIf Cells(i + 2, 17).Value = "Medium" Then
                    MsgBox "No more than two medium tasks should be put together in a row"
                    End
                ElseIf Cells(i + 2, 17).Value = "Hard" Then
                    MsgBox ("Please have an easy task after two medium tasks!")
                    End
                End If
                    
            End If
        ElseIf Cells(i, 17).Value = "Hard" Then
            If Cells(i + 1, 17).Value = "Hard" Then
                MsgBox "Please Split up hard tasks with medium and easy ones"
                End
            End If
        End If
        
    End If
Next i
MsgBox "You have great interleaving, keep it up!"


End Sub
Sub ClearTaskTime()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Journal")
    ws.Range("A3:C24").ClearContents
    ws.Range("A3:C24").Interior.Color = RGB(218, 233, 248)
End Sub
Sub ClearProductivity()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Journal")
    ws.Range("H3:J24").ClearContents
    ws.Range("H3:J24").Interior.Color = RGB(242, 206, 239)
End Sub



