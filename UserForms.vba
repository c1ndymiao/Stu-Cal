' UserForms (Code)
Private Sub MainPageButtonInstructions_Click()
    MainPage.Show
End Sub

Private Sub AddATaskButton_Click()
    ToDoForm.Show
End Sub

Private Sub ScheduleTimeBlocksButton_Click()
    ScheduleForm.Show
End Sub
' Schedule Form:
Private Sub UserForm_Initialize()

    With ScheduleForm.DateDropDown
        .AddItem "Sunday"
        .AddItem "Monday"
        .AddItem "Tuesday"
        .AddItem "Wednesday"
        .AddItem "Thursday"
        .AddItem "Friday"
        .AddItem "Saturday"
    End With
    
    With ScheduleForm.CategoryDropDown
        .AddItem "Social"
        .AddItem "Study"
        .AddItem "Personal"
        .AddItem "Other"
    End With


    With ScheduleForm.TimeDropDown
        .AddItem "12:00 AM"
        .AddItem "1:00 AM"
        .AddItem "2:00 AM"
        .AddItem "3:00 AM"
        .AddItem "4:00 AM"
        .AddItem "5:00 AM"
        .AddItem "6:00 AM"
        .AddItem "7:00 AM"
        .AddItem "8:00 AM"
        .AddItem "9:00 AM"
        .AddItem "10:00 AM"
        .AddItem "11:00 AM"
        .AddItem "12:00 PM"
        .AddItem "1:00 PM"
        .AddItem "2:00 PM"
        .AddItem "3:00 PM"
        .AddItem "4:00 PM"
        .AddItem "5:00 PM"
        .AddItem "6:00 PM"
        .AddItem "7:00 PM"
        .AddItem "8:00 PM"
        .AddItem "9:00 PM"
        .AddItem "10:00 PM"
        .AddItem "11:00 PM"
    End With
End Sub

Private Sub AddScheduleButton_Click()
    Dim selectedDate As String
    Dim selectedTime As String
    Dim eventName As String
    Dim category As String
    Dim startCol As Integer
    Dim timeRow As Integer
    Dim duration As Integer
    Dim dayNames As Variant
    Dim timeBlocks As Variant
    Dim combinedEvent As String

    eventName = EventTextBox.Value
    selectedTime = TimeDropDown.Value
    selectedDate = DateDropDown.Value
    category = CategoryDropDown.Value

    If eventName = "" Or selectedTime = "" Or selectedDate = "" Then
        MsgBox "Please fill in all fields."
        Exit Sub
    End If

    dayNames = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    timeBlocks = Array("12:00 AM", "1:00 AM", "2:00 AM", "3:00 AM", "4:00 AM", "5:00 AM", "6:00 AM", "7:00 AM", "8:00 AM", _
                        "9:00 AM", "10:00 AM", "11:00 AM", "12:00 PM", "1:00 PM", "2:00 PM", "3:00 PM", "4:00 PM", _
                        "5:00 PM", "6:00 PM", "7:00 PM", "8:00 PM", "9:00 PM", "10:00 PM", "11:00 PM")

    startCol = Application.Match(selectedDate, dayNames, 0) + 1
    timeRow = Application.Match(selectedTime, timeBlocks, 0)

    combinedEvent = eventName
    
    Set cell = Sheets("Weekly Calendar").Cells(timeRow + 1, startCol)
    cell.Value = combinedEvent
    cell.WrapText = True
    
    
    Select Case category
        Case "Social"
            cell.Interior.Color = RGB(255, 223, 186) ' Light orange
        Case "Study"
            cell.Interior.Color = RGB(186, 255, 186) ' Light green
        Case "Personal"
            cell.Interior.Color = RGB(186, 186, 255) ' Light blue
        Case "Other"
            cell.Interior.Color = RGB(166, 201, 238)
    End Select

    duration = 1

    If duration > 1 Then
        Sheets("Weekly Calendar").Range(Cells(timeRow + 1, startCol), Cells(timeRow + duration - 1, startCol)).Merge
        Sheets("Weekly Calendar").Range(Cells(timeRow + 1, startCol), Cells(timeRow + duration - 1, startCol)).Value = combinedEvent
        Sheets("Weekly Calendar").Range(Cells(timeRow + 1, startCol), Cells(timeRow + duration - 1, startCol)).WrapText = True
    End If

    EventTextBox.Value = ""
    TimeDropDown.Value = ""
    DateDropDown.Value = ""
    CategoryDropDown.Value = ""

    MsgBox "Event successfully added to schedule!"
End Sub

' To Do Form:
Private Sub UserForm_Initialize()

    With ToDoForm.PriorityLevelComboBox
        .AddItem "Low"
        .AddItem "Medium"
        .AddItem "High"
    End With
    
    With ToDoForm.cbDifficulty
        .AddItem "Easy"
        .AddItem "Medium"
        .AddItem "Hard"
    End With
    
    With ToDoForm.SubjectComboBox
    .AddItem "Math"
    .AddItem "English"
    .AddItem "Chemistry"
    .AddItem "Physics"
    .AddItem "Computer Science"
    .AddItem "Engineering"
    End With
    
    With ToDoForm.AssessmentCategoryComboBox
    .AddItem "Quiz"
    .AddItem "Test"
    .AddItem "Exam"
    .AddItem "Assignment"
    .AddItem "Homework"
    End With


End Sub

Private Sub AddTaskButton_Click()
    Dim taskName As String
    Dim dueDate As String
    Dim timeEstimation As String
    Dim priorityLevel As String
    Dim Difficulty As String
    Dim subject As String
    Dim assessmentCategory As String
    Dim weeklyCalendarSheet As Worksheet
    Dim lastRow As Integer
    Dim taskRow As Range

    Set weeklyCalendarSheet = ThisWorkbook.Sheets("Weekly Calendar")
    
    taskName = TaskTextbox.Value
    dueDate = DueDateTextBox.Value
    timeEstimation = TimeEstimationTextbox.Value
    priorityLevel = PriorityLevelComboBox.Value
    Difficulty = cbDifficulty.Value
    subject = SubjectComboBox.Value
    assessmentCategory = AssessmentCategoryComboBox.Value
    
    If taskName = "" Or dueDate = "" Or timeEstimation = "" Or priorityLevel = "" Or subject = "" Or assessmentCategory = "" Then
        MsgBox "Please fill in all fields."
        Exit Sub
    End If
    
    lastRow = weeklyCalendarSheet.Cells(weeklyCalendarSheet.Rows.Count, "J").End(xlUp).Row + 1
    If lastRow > 26 Then
        MsgBox "The weekly tasks table is full."
        Exit Sub
    End If
    
    If lastRow < 3 Then lastRow = 3
    
    Set taskRow = weeklyCalendarSheet.Range("J" & lastRow)
    
    weeklyCalendarSheet.Range("J" & lastRow & ":K" & lastRow).Merge
    taskRow.Value = taskName
    taskRow.Offset(0, 1).Value = dueDate
    taskRow.Offset(0, 2).Value = timeEstimation
    taskRow.Offset(0, 3).Value = priorityLevel
    taskRow.Offset(0, 4).Value = subject
    taskRow.Offset(0, 5).Value = assessmentCategory
    taskRow.Offset(0, 6).Value = Difficulty
    
    Select Case priorityLevel
        Case "High"
            taskRow.Offset(0, 3).Interior.Color = RGB(255, 0, 0)
        Case "Medium"
            taskRow.Offset(0, 3).Interior.Color = RGB(255, 255, 0)
        Case "Low"
            taskRow.Offset(0, 3).Interior.Color = RGB(0, 255, 0)
        Case Else
            taskRow.Offset(0, 3).Interior.ColorIndex = xlNone
    End Select
    
    TaskTextbox.Value = ""
    DueDateTextBox.Value = ""
    TimeEstimationTextbox.Value = ""
    PriorityLevelComboBox.Value = ""
    SubjectComboBox.Value = ""
    AssessmentCategoryComboBox.Value = ""
    cbDifficulty.Value = ""
    
    MsgBox "Task successfully added to the Weekly Calendar!"
End Sub

' Log Reflection:
Private Sub UserForm_Initialize()
    With Me.ProdRating
        .AddItem "Highly Focused"
        .AddItem "Made Good Progress"
        .AddItem "Okay/Mixed"
        .AddItem "Distracted/Unproductive"
    End With
End Sub

Private Sub LogReflectionButton_Click()
    Dim refdate As String
    Dim rating As String
    Dim reflectionText As String
    Dim ws As Worksheet
    Dim lastRow As Integer
    Dim entryRow As Range

    Set ws = ThisWorkbook.Sheets("Journal")

    refdate = logDate.Value
    rating = ProdRating.Value
    reflectionText = reflection.Value

    If refdate = "" Or rating = "" Or reflectionText = "" Then
        MsgBox "Please fill in all fields."
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row + 1
    If lastRow < 4 Then lastRow = 4

    Set entryRow = ws.Range("H" & lastRow)

    entryRow.Value = refdate
    entryRow.Offset(0, 1).Value = rating
    entryRow.Offset(0, 2).Value = reflectionText
    
    If rating = "Highly Focused" Then
        entryRow.Offset(0, 1).Interior.Color = RGB(135, 206, 250)

    ElseIf rating = "Made Good Progress" Then
        entryRow.Offset(0, 1).Interior.Color = RGB(144, 238, 144)

    ElseIf rating = "Okay/Mixed" Then
        entryRow.Offset(0, 1).Interior.Color = RGB(255, 255, 153)

    ElseIf rating = "Distracted/Unproductive" Then
        entryRow.Offset(0, 1).Interior.Color = RGB(255, 165, 0)
    End If
    
    logDate.Value = ""
    ProdRating.Value = ""
    reflection.Value = ""

    MsgBox "Reflection successfully added!"
    Unload Me
End Sub


' Log Task Time:
Private Sub LogEntryButton_Click()
    Dim taskName As String
    Dim timeEstimated As String
    Dim actualTime As String
    Dim taskSheet As Worksheet
    Dim lastRow As Integer
    Dim entryRow As Range

    Set taskSheet = ThisWorkbook.Sheets("Journal")

    taskName = Task.Value
    timeEstimated = TimeEst.Value
    actualTime = CompTime.Value

    If taskName = "" Or timeEstimated = "" Or actualTime = "" Then
        MsgBox "Please fill in all fields."
        Exit Sub
    End If

    lastRow = taskSheet.Cells(taskSheet.Rows.Count, "A").End(xlUp).Row + 1
    If lastRow < 4 Then lastRow = 4

    Set entryRow = taskSheet.Range("A" & lastRow)

    entryRow.Value = taskName
    entryRow.Offset(0, 1).Value = timeEstimated
    entryRow.Offset(0, 2).Value = actualTime
    
    If actualTime > timeEstimated Then
        entryRow.Offset(0, 2).Interior.Color = RGB(255, 102, 102)
    Else
        entryRow.Offset(0, 2).Interior.Color = RGB(144, 238, 144)
    End If

    Task.Value = ""
    TimeEst.Value = ""
    CompTime.Value = ""

    MsgBox "Log successfully added to the Task Time Tracker!"
    Unload Me
End Sub
