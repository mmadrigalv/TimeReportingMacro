Sub Report_Generator()
    'CONSTANTS
    Const wdLineStyleSingle As Integer = 1
    
    Const wdAlignParagraphLeft As Integer = 0
    Const wdAlignParagraphCenter As Integer = 1
    Const wdAlignParagraphRight As Integer = 2
    
    Const wdCellAlignVerticalCenter As Integer = 1
    
    Const wdLineWidth075pt As Integer = 6
    Const wdLineWidth100pt As Integer = 8
    Const wdLineWidth225pt As Integer = 18
    
    Const wdBorderTop As Integer = -1
    Const wdBorderLeft As Integer = -2
    Const wdBorderBottom As Integer = -3
    Const wdBorderRight As Integer = -4
    
    const AtmoseraClient = "Atmosera"

    ' GET THE WORKBOOK
    Dim currentWorkbook As Workbook
    Set currentWorkbook = ActiveWorkbook

    Dim path As String
    path = currentWorkbook.path

    'GET THE SHEET
    Dim ws As Worksheet
    Set ws = currentWorkbook.Sheets("Detailed Report")
    ws.AutoFilterMode = False
    
    'GET COLUMN INDEX
    Dim headerRange As Range
    Dim header As Range
    Set headerRange = ws.Rows(1) 'First row
    
    Dim headersToFind() As String
    headersToFind = Split("Project,Client,Description,Task,User,Start Date,Duration (decimal)", ",")
    Dim headerText As Variant
    
    Set columnIndexName = CreateObject("Scripting.Dictionary")
    Dim counter As Integer
    For Each headerText In headersToFind
        counter = 1
        For Each header In headerRange.Cells
            If CStr(header.Value) = headerText Then
                columnIndexName.Add headerText, counter
                Exit For
            End If
            counter = counter + 1
        Next header
    Next headerText
    
    Dim projectColumnIndex As Integer
    projectColumnIndex = columnIndexName.Item("Project")
    
    Dim clientColumnIndex As Integer
    clientColumnIndex = columnIndexName.Item("Client")
    
    Dim descriptionColumnIndex As Integer
    descriptionColumnIndex = columnIndexName.Item("Description")
    
    Dim taskColumnIndex As Integer
    taskColumnIndex = columnIndexName.Item("Task")
    
    Dim userColumnIndex As Integer
    userColumnIndex = columnIndexName.Item("User")
    
    Dim startDateColumnIndex As Integer
    startDateColumnIndex = columnIndexName.Item("Start Date")
    
     Dim durationColumnIndex As Integer
     durationColumnIndex = columnIndexName.Item("Duration (decimal)")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    Dim sortRange As Range
    Set sortRange = ws.Range("A1:Q" & lastRow)
    
    Dim sortColumnDate As Range
    Set sortColumnDate = sortRange.Columns(startDateColumnIndex)
    
    Dim sortColumnUser As Range
    Set sortColumnUser = sortRange.Columns(userColumnIndex)
    
    Dim sortColumnClient As Range
    Set sortColumnClient = sortRange.Columns(clientColumnIndex)
    
    sortRange.Sort Key1:=sortColumnDate, Key2:=sortColumnUser, Key3:=sortColumnClient, Order1:=xlAscending, Order2:=xlAscending, Order3:=xlAscending, header:=xlYes
 
    Set usersName = New Collection
    Set clientsName = New Collection
    Set projectsName = New Collection
    Set weeks = New Collection
    
    Dim key As String
    Dim rowIndex As Long
    Dim lasRow As Long
    
    Set userNameMapping = CreateObject("Scripting.Dictionary")
    Set clientNameMapping = CreateObject("Scripting.Dictionary")
    Set projectNameMapping = CreateObject("Scripting.Dictionary")
    Set dateMapping = CreateObject("Scripting.Dictionary")
    
    Set descriptionValueMapping = CreateObject("Scripting.Dictionary")
    Set durationValueMapping = CreateObject("Scripting.Dictionary")
    
    Dim duplicatedRow As Object
    Set duplicatedRow = CreateObject("Scripting.Dictionary")

   Dim minDate As Date
   minDate = GetFormatedDate(ws.Cells(2, startDateColumnIndex).Value)
   
   Dim maxDate As Date
   maxDate = minDate
   
    For i = 2 To lastRow
        Dim dateProcessed As Date
        dateProcessed = GetFormatedDate(ws.Cells(i, startDateColumnIndex).Value)
        If dateProcessed < minDate Then
            minDate = dateProcessed
        End If
        
        If dateProcessed > maxDate Then
            maxDate = dateProcessed
        End If
    Next i

    Dim startingDate As String
    startingDate = "Saturday"
    
    For rowIndex = 2 To lastRow

        user = ws.Cells(rowIndex, userColumnIndex).Value
        
        Dim client As String
        client = ws.Cells(rowIndex, clientColumnIndex).Value
        
        Dim project As String
        project = ws.Cells(rowIndex, projectColumnIndex).Value
        
        Dim startDate As String
        startDate = GetFormatedDate(ws.Cells(rowIndex, startDateColumnIndex).Value)
        
        Dim durationInHours As Double
        durationInHours = ws.Cells(rowIndex, durationColumnIndex).Value
        
        Dim taskDescription As String
        taskDescription = ws.Cells(rowIndex, taskColumnIndex).Value
        If taskDescription <> "" Then
            taskDescription = taskDescription & " " & ws.Cells(rowIndex, descriptionColumnIndex).Value
        Else
         taskDescription = ws.Cells(rowIndex, descriptionColumnIndex).Value
        End If
        
        
        key = user & "_" & client & "_" & project & "_" & startDate
        
        userNameMapping.Add rowIndex, user
       
        On Error Resume Next
            usersName.Add user, CStr(user)
        On Error GoTo 0
        
        clientNameMapping.Add rowIndex, client
        
        On Error Resume Next
            clientsName.Add client, CStr(client)

            If CStr(client) = AtmoseraClient And startingDate = "Saturday" Then
                DownloadAtmosEraLogo
            End If
        On Error GoTo 0
        
        projectNameMapping.Add rowIndex, project
        
        On Error Resume Next
            projectsName.Add project, CStr(project)
        On Error GoTo 0
        
        dateMapping.Add rowIndex, startDate
        
         If descriptionValueMapping.Exists(key) Then
            duplicatedRow.Add rowIndex, rowIndex
            Dim line As Variant
            Dim descriptionLines() As String
            descriptionLines = Split(Trim(descriptionValueMapping(key)), vbCr)
            descriptionValueMapping(key) = ""
            
            Dim hasFound As Boolean
            hasFound = False
            For Each line In descriptionLines
                If Not line = "" Then
                    If StringEndsWith(line, taskDescription) Then
                        hasFound = True
                        
                        Dim startPosition As Integer
                        startPosition = InStr(1, line, "- ") + 2
                        
                        Dim endPosition As Integer
                        endPosition = InStr(1, line, " h ")
                        
                        Dim duration As Double
                        duration = CDbl(Mid(line, startPosition, endPosition - startPosition))
                        duration = duration + durationInHours
                        
                        If descriptionValueMapping(key) = "" Then
                             descriptionValueMapping(key) = descriptionValueMapping(key) & "- " & Format(duration, "0.00") & " h " & taskDescription
                        Else
                             descriptionValueMapping(key) = descriptionValueMapping(key) & vbCr & "- " & Format(duration, "0.00") & " h " & taskDescription
                        End If
                    Else
                        If descriptionValueMapping(key) = "" Then
                            descriptionValueMapping(key) = descriptionValueMapping(key) & line
                        Else
                            descriptionValueMapping(key) = descriptionValueMapping(key) & vbCr & line
                        End If
                    End If
                End If
            Next line
            
            If Not hasFound Then
                descriptionValueMapping(key) = descriptionValueMapping(key) & vbCr & "- " & Format(durationInHours, "0.00") & " h " & taskDescription
             End If
         Else
            descriptionValueMapping.Add key, "- " & Format(durationInHours, "0.00") & " h " & taskDescription
         End If
        
         If durationValueMapping.Exists(key) Then
            durationValueMapping(key) = Format(durationValueMapping(key) + durationInHours, "0.00")
         Else
            durationValueMapping.Add key, Format(durationInHours, "0.00")
         End If
        
    Next rowIndex

    Dim daysOfWeek() As String

    Dim userItem As Variant
    For Each userItem In usersName
    
        Dim clientItem As Variant
        For Each clientItem In clientsName
            If clientItem = AtmoseraClient Then
                startingDate = "Monday"
            Else
                startingDate = "Saturday"
            End If
            
            Dim projectItem As Variant
            For Each projectItem In projectsName
            
            
            If startingDate = "Saturday" Then
                If Weekday(minDate) <> vbSaturday Then
                    minDate = minDate + 1 - Weekday(minDate, vbSaturday) 'previous Saturday
                End If
                
                If Weekday(maxDate) <> vbFriday Then
                    maxDate = maxDate + 8 - Weekday(maxDate, vbFriday) 'next friday
                End If
            Else
                If Weekday(minDate) <> vbMonday Then
                    minDate = minDate + 1 - Weekday(minDate, vbMonday) 'previous monday
                End If
                
                If Weekday(maxDate) <> vbSunday Then
                    maxDate = maxDate + 8 - Weekday(maxDate, vbSunday) 'next Sunday
                End If
            End If
        
            Dim initialDate As Date
            initialDate = minDate
            Dim weekId As Integer
            weekId = 1
            weeks.Add 1
            
            Dim weekInitialDay As Date
            weekInitialDay = vbNull
            
            Dim weekLastDay As Date
            weekLastDay = vbNull
        
            'GET WEEKS STARTING DAY SATURDAY
            Dim daysCount As Integer
            daysCount = 1
            While initialDate <= maxDate
                If daysCount = 7 Then
                    weekId = weekId + 1
                    weeks.Add weekId
                    daysCount = 1
                Else
                    daysCount = daysCount + 1
                End If
                
                initialDate = initialDate + 1
            Wend
            
             
             weekLastDay = vbNull
             Dim weekItem As Variant
             For Each weekItem In weeks
                
                 If weekLastDay >= maxDate Then
                    GoTo skipWeek
                 End If

                 If weekLastDay = vbNull Then
                    weekInitialDay = minDate
                Else
                    weekInitialDay = weekLastDay + 1
                End If
                weekLastDay = weekInitialDay + 6
                
                ws.AutoFilterMode = False
                
                ws.Range(GetColumnLetter(projectColumnIndex) & "1").AutoFilter Field:=projectColumnIndex, Criteria1:=projectItem
                ws.Range(GetColumnLetter(clientColumnIndex) & "1").AutoFilter Field:=clientColumnIndex, Criteria1:=clientItem
                ws.Range(GetColumnLetter(userColumnIndex) & "1").AutoFilter Field:=userColumnIndex, Criteria1:=userItem
                
                Dim criteriaArray() As Variant
                Dim iterationDate As Date
                ReDim criteriaArray(0)
                iterationDate = weekInitialDay
                criteriaArray(0) = Format(iterationDate, "dd/MM/yyyy")
                Do While iterationDate < weekLastDay
                    iterationDate = iterationDate + 1
                    ReDim Preserve criteriaArray(UBound(criteriaArray) + 1)
                    criteriaArray(UBound(criteriaArray)) = Format(iterationDate, "dd/MM/yyyy")
                Loop
                ws.Range(GetColumnLetter(startDateColumnIndex) & "1").AutoFilter Field:=startDateColumnIndex, Criteria1:=criteriaArray, Operator:=xlFilterValues
                
                'Filter Rows Count
                Dim filterRange As Range
                Set filterRange = ws.Range(GetColumnLetter(startDateColumnIndex) & "1:" & GetColumnLetter(startDateColumnIndex) & lastRow)
            
                Dim visibleRowCount As Long
                Dim cell As Range
            
                ' Initialize the visible row count
                visibleRowCount = 0
            
                ' Loop through each cell in the filtered range
                For Each cell In filterRange
                    ' Check if the cell is visible
                    If cell.EntireRow.Hidden = False Then
                        visibleRowCount = visibleRowCount + 1
                    End If
                Next cell
                'Filter Rows Count
                
            
                If visibleRowCount <= 1 Then
                    ws.AutoFilterMode = False
                    GoTo skipWeek
                End If
                ws.AutoFilterMode = False
                 
                 Dim wordApp As Object
                 Set wordApp = CreateObject("Word.Application")
                 wordApp.Visible = False
                 
                 Dim wordDoc As Object
                 Set wordDoc = wordApp.Documents.Add

                 If clientItem = AtmoseraClient Then
                    Dim pic As Variant
                    Set rng = wordDoc.Content
                    rng.Collapse Direction:=0 'wdCollapseEnd
                    
                    ' Insert the image with specified properties
                    Set pic = wordDoc.InlineShapes.AddPicture( _
                        Filename:=ActiveWorkbook.path & "\atmosera_logo.png", _
                        LinkToFile:=False, _
                        SaveWithDocument:=True, _
                        Range:=rng)

                    pic.Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                End If
                 
                 With wordDoc
                     .Content.Font.Size = 12
                     
                     .Content.Paragraphs.Add
                     With .Content.Paragraphs.Last.Range
                         .Text = "Developer Status Report"
                         .Font.Name = "Arial"
                         .Font.Size = 16
                         .Font.Bold = True
                         .ParagraphFormat.Alignment = wdAlignParagraphLeft
                     End With
                     
                     Dim userNameIndex As Integer
                     userNameIndex = userNameMapping.Item(CStr(userItem)) + 2
                     
                     Dim clientNameIndex As Integer
                     clientNameIndex = clientNameMapping.Item(clientItem)

                    
                     .Content.Paragraphs.Add
                     With .Content.Paragraphs.Last.Range
                         .Text = "Client Name: " & vbTab & vbTab & CStr(clientItem)
                         .Font.Size = 10
                         .ParagraphFormat.Alignment = wdAlignParagraphLeft
                         .Font.Bold = True
            
                        ' Set the specific portions of text to bold
                         .Words(1).Font.Bold = False ' Client
                         .Words(2).Font.Bold = False ' Name
                     End With
                     
                     .Content.Paragraphs.Add
                     With .Content.Paragraphs.Last.Range
                         .Text = "Project Name: " & vbTab & vbTab & CStr(projectItem)
                         .Font.Size = 10
                         .ParagraphFormat.Alignment = wdAlignParagraphLeft
                         .Font.Bold = True
            
                        ' Set the specific portions of text to bold
                         .Words(1).Font.Bold = False ' Project
                         .Words(2).Font.Bold = False ' Name
                     End With
                     
                     .Content.Paragraphs.Add
                     With .Content.Paragraphs.Last.Range
                         .Text = "Developer Name: " & vbTab & CStr(userItem)
                         .Font.Size = 10
                         .ParagraphFormat.Alignment = wdAlignParagraphLeft
                         .Font.Bold = True
            
                        ' Set the specific portions of text to bold
                         .Words(1).Font.Bold = False ' Developer
                         .Words(2).Font.Bold = False ' Name
                     End With
                     
                     .Content.Paragraphs.Add
                     With .Content.Paragraphs.Last.Range
                         .Text = "Week Ending: " & vbTab & vbTab & Format(weekLastDay, "MM/dd/yyyy") & vbCr
                         .Font.Size = 10
                         .ParagraphFormat.Alignment = wdAlignParagraphLeft
                         .Font.Bold = True
            
                        ' Set the specific portions of text to bold
                         .Words(1).Font.Bold = False ' Week
                         .Words(2).Font.Bold = False ' Ending
                                     
                        .InsertParagraphAfter
                     End With
                
                     .Tables.Add Range:=.Content.Paragraphs.Add.Range, NumRows:=2, NumColumns:=4
                     With .Tables(1)
                         .Borders.Enable = True
                        
                         .cell(1, 1).Range.Text = "Weekly Activity Summary (Required)"
                         .cell(1, 1).Merge MergeTo:=.cell(1, 4)
                         .cell(1, 1).SetWidth ColumnWidth:=500, RulerStyle:=wdAdjustFirstColumn
                         .cell(1, 1).Range.Paragraphs.Alignment = wdAlignParagraphCenter
                         .cell(1, 1).Range.Paragraphs(1).Range.Font.Smallcaps = True
                         
                         .cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0
                         .cell(1, 1).VerticalAlignment = wdCellAlignVerticalCenter
                         
                          .cell(1, 1).Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                          .cell(1, 1).Borders(wdBorderTop).LineWidth = wdLineWidth225pt
                          .cell(1, 1).Borders(wdBorderTop).Color = RGB(0, 0, 0)
                                                    
                          .cell(1, 1).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                          .cell(1, 1).Borders(wdBorderLeft).LineWidth = wdLineWidth225pt
                          .cell(1, 1).Borders(wdBorderLeft).Color = RGB(0, 0, 0)
                                                   
                          
                          .cell(1, 1).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                          .cell(1, 1).Borders(wdBorderRight).LineWidth = wdLineWidth225pt
                          .cell(1, 1).Borders(wdBorderRight).Color = RGB(0, 0, 0)
            
                         
                         .cell(2, 1).Range.Text = "Day"
                         .cell(2, 1).SetWidth ColumnWidth:=40, RulerStyle:=wdAdjustFirstColumn
                         .cell(2, 1).Range.Paragraphs(1).Range.Font.Smallcaps = True
                         .cell(2, 1).Range.ParagraphFormat.SpaceAfter = 0
                         .cell(2, 1).VerticalAlignment = wdCellAlignVerticalCenter
                         .cell(2, 1).Range.Paragraphs.Alignment = wdAlignParagraphCenter
                         
                         .cell(2, 1).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                         .cell(2, 1).Borders(wdBorderLeft).LineWidth = wdLineWidth225pt
                         
                         .cell(2, 2).Range.Text = "Date"
                         .cell(2, 2).SetWidth ColumnWidth:=80, RulerStyle:=wdAdjustFirstColumn
                         .cell(2, 2).Range.Paragraphs(1).Range.Font.Smallcaps = True
                         .cell(2, 2).Range.ParagraphFormat.SpaceAfter = 0
                         .cell(2, 2).VerticalAlignment = wdCellAlignVerticalCenter
                         .cell(2, 2).Range.Paragraphs.Alignment = wdAlignParagraphCenter
                         
                         .cell(2, 3).Range.Text = "Hours"
                         .cell(2, 3).SetWidth ColumnWidth:=50, RulerStyle:=wdAdjustFirstColumn
                         .cell(2, 3).Range.Paragraphs(1).Range.Font.Smallcaps = True
                         .cell(2, 3).Range.ParagraphFormat.SpaceAfter = 0
                         .cell(2, 3).VerticalAlignment = wdCellAlignVerticalCenter
                         .cell(2, 3).Range.Paragraphs.Alignment = wdAlignParagraphCenter
                         
                         .cell(2, 4).Range.Text = "Activity"
                         .cell(2, 4).SetWidth ColumnWidth:=330, RulerStyle:=wdAdjustFirstColumn
                         .cell(2, 4).Range.Paragraphs(1).Range.Font.Smallcaps = True
                         .cell(2, 4).Range.ParagraphFormat.SpaceAfter = 0
                         .cell(2, 4).VerticalAlignment = wdCellAlignVerticalCenter
                         .cell(2, 4).Range.Paragraphs.Alignment = wdAlignParagraphCenter
                         .cell(2, 4).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                         .cell(2, 4).Borders(wdBorderRight).LineWidth = wdLineWidth225pt
                         
                          .cell(1, 1).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                          .cell(1, 1).Borders(wdBorderBottom).LineWidth = wdLineWidth225pt
                          .cell(1, 1).Borders(wdBorderBottom).Color = RGB(0, 0, 0)
                         
                         .Rows(1).Range.Font.Name = "Times New Roman"
                         .Rows(1).Range.Font.Size = 10
                         .Rows(1).Range.Font.Bold = True
                         .Rows(1).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
                         .Rows(1).Height = 20
                         
                         .Rows(2).Range.Font.Name = "Times New Roman"
                         .Rows(2).Range.Font.Size = 10
                         .Rows(2).Range.Font.Bold = True
                         .Rows(2).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
                         .Rows(2).Height = 20
                     End With
                     
                                
                     Dim dayOfWeek As String
                     Dim activityDate As Date
                     Dim totalHours As Double
                     totalHours = 0
                     
                     Dim wordRowIndex As Integer
                     wordRowIndex = 2
                                          
                     weekLastDay = weekInitialDay + 6
                     
                    For rowIndex = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).Row
                        If duplicatedRow.Exists(rowIndex) Or _
                           userNameMapping(rowIndex) <> CStr(userItem) Or _
                           clientNameMapping(rowIndex) <> clientItem Or _
                           projectNameMapping(rowIndex) <> projectItem Or _
                           CDate(dateMapping(rowIndex)) < weekInitialDay Or _
                           CDate(dateMapping(rowIndex)) > weekLastDay Then
                            GoTo skipIteration
                        End If
                                            
                        user = ws.Cells(rowIndex, userColumnIndex).Value
                        
                        client = ws.Cells(rowIndex, clientColumnIndex).Value
                        
                        project = ws.Cells(rowIndex, projectColumnIndex).Value
                        
                        startDate = GetFormatedDate(ws.Cells(rowIndex, startDateColumnIndex).Value) 'Format(ws.Cells(rowIndex, 10).Value, "dd/MM/yyyy")
                        
                        key = user & "_" & client & "_" & project & "_" & startDate
                    
                        activityDate = startDate
                        
                        dayOfWeek = Format(activityDate, "ddd")
                        
                        Select Case dayOfWeek
                           Case "sáb.", "Sat"
                              dayOfWeek = "Sat."
                           Case "dom.", "Sun"
                              dayOfWeek = "Sun."
                           Case "lun.", "Mon"
                               dayOfWeek = "Mon."
                           Case "mar.", "Tue"
                               dayOfWeek = "Tue."
                            Case "mié.", "Wed"
                               dayOfWeek = "Wed."
                            Case "jue.", "Thu"
                               dayOfWeek = "Thu."
                            Case "vie.", "Fri"
                               dayOfWeek = "Fri."
                        End Select
                        
                        
                        .Tables(1).Rows.Add
                        With .Tables(1)
                            .cell(wordRowIndex + 2, 1).Range.Text = dayOfWeek
                            .cell(wordRowIndex + 2, 1).Range.Paragraphs(1).Range.Font.Smallcaps = False
                            
                            .cell(wordRowIndex + 2, 2).Range.Text = Format(activityDate, "MM/dd/yyyy")
                            .cell(wordRowIndex + 2, 2).Range.Paragraphs(1).Range.Font.Smallcaps = False
                            
                            .cell(wordRowIndex + 2, 3).Range.Text = durationValueMapping(key)
                            .cell(wordRowIndex + 2, 3).Range.Paragraphs.Alignment = wdAlignParagraphRight
                            .cell(wordRowIndex + 2, 3).Range.Paragraphs(1).Range.Font.Smallcaps = False
                            
                            .cell(wordRowIndex + 2, 4).Range.Text = descriptionValueMapping(key)
                            .cell(wordRowIndex + 2, 4).Range.Paragraphs.Alignment = wdAlignParagraphLeft
                            
                            Dim cellParagraphs As Variant
                            For Each cellParagraphs In .cell(wordRowIndex + 2, 4).Range.Paragraphs
                                cellParagraphs.Range.Font.Smallcaps = False
                            Next cellParagraphs

                            .Rows(wordRowIndex + 1).Range.Font.Bold = False
                            .Rows(wordRowIndex + 1).Range.Shading.BackgroundPatternColor = RGB(255, 255, 255)
                            .Rows(wordRowIndex).Borders(wdBorderBottom).LineWidth = wdLineWidth075pt
                            .cell(wordRowIndex + 1, 1).Range.Font.Bold = True
                        End With
                        
                        ' Accumulate total hours
                        totalHours = totalHours + durationValueMapping(key)
                        
                        wordRowIndex = wordRowIndex + 1
skipIteration:
                    Next rowIndex
                        
                        ' Add Totalization row
                        .Tables(1).Rows.Add
                        With .Tables(1)
                            .cell(.Rows.count, 1).Range.Text = "Total"
                            .cell(.Rows.count, 1).Range.Font.Bold = True
                            .cell(.Rows.count, 3).Range.Text = Format(totalHours, "0.00") ' Display the accumulated total hours
                            .cell(.Rows.count, 3).Range.Font.Bold = True
                            .cell(.Rows.count, 3).Range.Paragraphs.Alignment = wdAlignParagraphRight
                            
                            .Rows(.Rows.count).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                            .Rows(.Rows.count).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                            .Rows(.Rows.count).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                            .Rows(.Rows.count).Borders(wdBorderLeft).LineWidth = wdLineWidth225pt
                            .Rows(.Rows.count).Borders(wdBorderBottom).LineWidth = wdLineWidth225pt
                            .Rows(.Rows.count).Borders(wdBorderRight).LineWidth = wdLineWidth225pt
                        End With

                        If wordRowIndex > 2 Then
                            If wordRowIndex < 9 Then
                                ' Loop through each day of the week
                                Dim dateCol As Date
                                dateCol = weekInitialDay
                                
                                If (startingDate = "Saturday") Then
                                    daysOfWeek = Split("Sat.,Sun.,Mon.,Tue.,Wed.,Thu.,Fri.", ",")
                                Else
                                    daysOfWeek = Split("Mon.,Tue.,Wed.,Thu.,Fri.,Sat.,Sun.", ",")
                                End If
                                
                                For i = 0 To UBound(daysOfWeek)
                                    found = False
                                    
                                    Dim dayCell As Variant
                                    dayCell = Split(.Tables(1).Rows(i + 3).Cells(1), ".")(0) & "."
                                    
                                    If UCase(dayCell) = UCase(daysOfWeek(i)) Or UCase(dayCell) = "Total" Then
                                        found = True
                                    End If
                                   
                                    If Not found Then
                                        .Tables(1).Rows.Add BeforeRow:=.Tables(1).Rows(i + 3)
                                        
                                        .Tables(1).cell(i + 3, 1).Range.Text = CStr(daysOfWeek(i))
                                        .Tables(1).Rows(i + 3).Borders(wdBorderBottom).LineWidth = wdLineWidth075pt
                                        .Tables(1).cell(i + 3, 2).Range.Text = CStr(Format(dateCol, "MM/dd/yyyy"))
                                        
                                        .Tables(1).cell(i + 3, 3).Range.Text = "0.00"
                                        .Tables(1).cell(i + 3, 3).Range.Paragraphs.Alignment = wdAlignParagraphRight
                                        .Tables(1).cell(i + 3, 3).Range.Font.Bold = False
                                    End If
                                    
                                    dateCol = dateCol + 1
                                Next i
                                
                            End If

                            .Tables(1).cell(10, 1).Merge MergeTo:=.Tables(1).cell(10, 2)
                            .Tables(1).cell(10, 1).Range.Paragraphs.Alignment = wdAlignParagraphRight

                            Dim otherTables() As String
                            otherTables = Split("ACCOMPLISHMENTS (REQUIRED),UNPLANNED TASKS,PROGRESS NOT ACHIEVED,PROGRESS PLANNED NEXT WEEK (REQUIRED),OPEN ISSUES OR CONCERNS,MISCELLANEOUS SCHEDULING", ",")
                            
                            For i = 0 To UBound(otherTables)
                                .Content.InsertParagraphAfter
                                
                                .Tables.Add Range:=.Content.Paragraphs.Add.Range, NumRows:=2, NumColumns:=1
                                With .Tables.Item(i + 2)
                                    .Borders.Enable = True
                                    .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                                    .Borders(wdBorderTop).LineWidth = wdLineWidth225pt
                                    
                                    .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                                    .Borders(wdBorderLeft).LineWidth = wdLineWidth225pt
                                    
                                    .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                                    .Borders(wdBorderBottom).LineWidth = wdLineWidth225pt
                                    
                                    .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                                    .Borders(wdBorderRight).LineWidth = wdLineWidth225pt
                            
                                    .cell(1, 1).Range.Text = UCase(otherTables(i))
                                    .cell(1, 1).SetWidth ColumnWidth:=500, RulerStyle:=wdAdjustFirstColumn
                                    .cell(1, 1).Range.Paragraphs.Alignment = wdAlignParagraphLeft
                                    .cell(1, 1).Range.Paragraphs(1).Range.Font.Smallcaps = True
                                    .cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0
                                    .cell(1, 1).VerticalAlignment = wdCellAlignVerticalCenter
                                                                        
                                    .cell(2, 1).SetWidth ColumnWidth:=500, RulerStyle:=wdAdjustFirstColumn
                                    .cell(2, 1).Range.Paragraphs(1).Range.Font.Smallcaps = False
                                    
                                    .Rows(1).Range.Font.Name = "Times New Roman"
                                    .Rows(1).Range.Font.Size = 10
                                    .Rows(1).Range.Font.Bold = True
                                    .Rows(1).Range.Shading.BackgroundPatternColor = RGB(230, 230, 230)
                                    .Rows(1).Height = 20
                                    
                                    .Rows(2).Range.Font.Name = "Times New Roman"
                                    .Rows(2).Range.Font.Size = 10
                                    .Rows(2).Range.Font.Bold = False
                                    .Rows(2).Range.Shading.BackgroundPatternColor = RGB(255, 255, 255)
                                    .Rows(2).Range.ListFormat.ApplyBulletDefault
                                End With
                            Next i
                            
                            .SaveAs path & "\Status Report - " & Format(weekLastDay, "MM-dd-yyyy") & " - " & CStr(clientItem) & " - " & CStr(projectItem) & " - " & CStr(userItem) & ".docx"
                            .Close
                        Else
                            wordApp.DisplayAlerts = False
            
                            ' Close all documents without saving changes
                            .Close SaveChanges:=wdDoNotSaveChanges
                           
                        End If
                    End With

                    wordApp.Quit
                    Set wordDoc = Nothing
                    Set wordApp = Nothing
skipWeek:
                Next weekItem
            Next projectItem
        Next clientItem
    Next userItem
    
    ws.AutoFilterMode = False
    MsgBox "Generación de reportes concluida.", vbOKOnly, "Final del proceso", vbNull, vbNull
End Sub

Private Function GetFormatedDate(dateToCast As String) As Date
     GetFormatedDate = DateValue(Format(dateToCast, "dd/MM/yyyy"))  'Cast String to Date Specifing that format on excel string is dd/MM/yyyy
End Function

Private Function GetColumnLetter(columnIndex As Integer) As String
    GetColumnLetter = Split(Cells(1, columnIndex).Address, "$")(1)
End Function

Private Function StringEndsWith(ByVal mainString As String, ByVal searchString As String) As Boolean
    Dim lengthMain As Long
    Dim lengthSearch As Long
    
    ' Get the lengths of the main string and the search string
    lengthMain = Len(mainString)
    lengthSearch = Len(searchString)
    
    ' Check if the main string ends with the search string
    If lengthMain >= lengthSearch Then
        StringEndsWith = (Right(mainString, lengthSearch) = searchString)
    Else
        StringEndsWith = False
    End If
End Function

Private Sub DownloadAtmosEraLogo()
    Dim url As String
    url = "https://lh3.googleusercontent.com/u/0/drive-viewer/AK7aPaBl62TAv4OpqReKAeRbCbYKlmx7HJEcNKlXSEky7uC_M7FwlH-RrXf6b42gmsfR3Sb7ImIhenwdEZ99munpeUObZcj_LA=w1860-h895" ' Replace with the direct URL of the image you obtained from the browser

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Send an HTTP GET request to the URL
    http.Open "GET", url, False
    http.send

    ' Check if the request was successful (status code 200)
    If http.Status = 200 Then
        Dim imageBytes() As Byte
        imageBytes = http.responseBody

        ' Specify the local file path where you want to save the image
        Dim filePath As String
        filePath = ActiveWorkbook.path & "\atmosera_logo.png"
        ' Save the image to the specified file path
        Open filePath For Binary As #1
        Put #1, , imageBytes
        Close #1

        'MsgBox "Image downloaded successfully."
    Else
        MsgBox "Failed to download the image. HTTP Status: " & http.Status
    End If
End Sub



