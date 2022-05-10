Option Explicit

Public Sub CreateClient(Authorization As String, UserAgent As String, ClientName As String)

    ' Purpose:
    ' Create a new Client on Float
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' ClientName - the name of the client
    
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request

        .Open "POST", "https://api.float.com/v3/clients", False

        .setRequestHeader "Authorization", "Bearer " & Authorization
        .setRequestHeader "User-Agent", UserAgent
        .setRequestHeader "Content-Type", "application/json"
        
    End With
    
    Dim Body As String, Message As String
    Body = "{" & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & ClientName & Chr(34)
    
    Body = Body & "}"
    Request.send Body
    
    ' Invalid field
    If Request.Status = 422 Then
    
        Message = "No client created." & vbNewLine & vbNewLine
        Message = Message & "Float API responded with: 422 Unprocessable Entity - The data supplied has failed validation." & vbNewLine & vbNewLine
        Message = Message & Request.responseText
        
        MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
        
    End If

End Sub
    
    
Public Sub CreateDepartment(Authorization As String, UserAgent As String, Department As String)

    ' Purpose:
    ' Create a new Department on Float
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' Department - the name of the department
    
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request

        .Open "POST", "https://api.float.com/v3/departments", False

        .setRequestHeader "Authorization", "Bearer " & Authorization
        .setRequestHeader "User-Agent", UserAgent
        .setRequestHeader "Content-Type", "application/json"
        
    End With
    
    Dim Body As String, Message As String
    Body = "{" & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & Department & Chr(34)
    
    Body = Body & "}"
    Request.send Body
    
    ' Invalid field
    If Request.Status = 422 Then
    
        Message = "No department created." & vbNewLine & vbNewLine
        Message = Message & "Float API responded with: 422 Unprocessable Entity - The data supplied has failed validation." & vbNewLine & vbNewLine
        Message = Message & Request.responseText
        
        MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
        
    End If

End Sub
        
        
Public Sub CreateHoliday(Authorization As String, UserAgent As String, Name As String, StartDate As String, Optional EndDate As String)

    ' Purpose:
    ' Create a new Holiday on Float
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' Name - the name of the holiday
    ' StartDate - the starting date of the holiday in the form YYYY-MM-DD
    ' EndDate - the ending date of the holiyda in the form YYYY-MM-DD
    
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request

        .Open "POST", "https://api.float.com/v3/holidays", False

        .setRequestHeader "Authorization", "Bearer " & Authorization
        .setRequestHeader "User-Agent", UserAgent
        .setRequestHeader "Content-Type", "application/json"
        
    End With
    
    Dim Body As String, Message As String
    Body = "{" & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & Name & Chr(34)
    Body = Body & "," & Chr(34) & "date" & Chr(34) & ":" & Chr(34) & StartDate & Chr(34)
    Body = Body & "," & Chr(34) & "end_date" & Chr(34) & ":" & Chr(34) & EndDate & Chr(34)
    
    Body = Body & "}"
    Request.send Body
    
    ' Invalid field
    If Request.Status = 422 Then
    
        Message = "No department created." & vbNewLine & vbNewLine
        Message = Message & "Float API responded with: 422 Unprocessable Entity - The data supplied has failed validation." & vbNewLine & vbNewLine
        Message = Message & Request.responseText
        
        MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
        
    End If

End Sub
        
        
Public Sub CreateMilestone(Authorization As String, UserAgent As String, Name As String, ProjectID As Long, StartDate As String, _
    Optional EndDate As String)

    ' Purpose:
    ' Create a new Milestone on Float
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' Name - the name of the milestone
    ' ProjectID - the project_id on Float of the Project this milestone belongs to
    ' StartDate - the date the milestone starts in the form YYYY-MM-DD
    ' EndDate - the date the milestone ends in the form YYYY-MM-DD
    
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request

        .Open "POST", "https://api.float.com/v3/milestones", False

        .setRequestHeader "Authorization", "Bearer " & Authorization
        .setRequestHeader "User-Agent", UserAgent
        .setRequestHeader "Content-Type", "application/json"
        
    End With
    
    Dim Body As String, Message As String
    Body = "{" & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & Name & Chr(34)
    Body = Body & "," & Chr(34) & "project_id" & Chr(34) & ":" & Chr(34) & ProjectID & Chr(34)
    Body = Body & "," & Chr(34) & "date" & Chr(34) & ":" & Chr(34) & StartDate & Chr(34)
    Body = Body & "," & Chr(34) & "end_date" & Chr(34) & ":" & Chr(34) & EndDate & Chr(34)
    
    Body = Body & "}"
    Request.send Body
    
    ' Invalid field
    If Request.Status = 422 Then
    
        Message = "No department created." & vbNewLine & vbNewLine
        Message = Message & "Float API responded with: 422 Unprocessable Entity - The data supplied has failed validation." & vbNewLine & vbNewLine
        Message = Message & Request.responseText
        
        MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
        
    End If

End Sub
    

Public Sub CreatePerson(Authorization As String, UserAgent As String, Name As String, Optional Email As String, Optional JobTitle As String, _
    Optional DepartmentID As Long, Optional Notes As String, Optional AutoEmail As Boolean = False, Optional FullTime As Boolean = True, _
    Optional WorkDaysHours As Collection, Optional Active As Boolean = True, Optional Contractor As Boolean = False, _
    Optional Tags As Collection, Optional StartDate As String, Optional EndDate As String, Optional DefaultHourlyRate As Double)

    ' Purpose:
    ' Create a new Person on Float
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' Name - person's full name
    ' Email - person's email address
    ' JobTitle - person's job title
    ' DepartmentID - the department_id of the department on Float
    ' Notes - notes on the person
    ' AutoEmail - whether or not the person's schedule should be emailed to them at the beginning of each week
    ' FullTime - whether or not the employee works full-time
    ' WorkDaysHours - the number of hours for each of the 7 days of the week the person works starting with Sunday
    ' Active - whether the person is current or an ex-employee
    ' Contractor - whether the person is a contractor or not
    ' Tags - any tags related to the person
    ' StartDate - the date the person started working in one of the following forms: YYYY-MM-DD, YY-MM-DD, DD-MM-YYYY
    ' EndDate - the date the person will stop working in one of the following forms: YYYY-MM-DD, YY-MM-DD, DD-MM-YYYY
    ' DefaultHourlyRate - the hourly rate of the person
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request

        .Open "POST", "https://api.float.com/v3/people", False

        .setRequestHeader "Authorization", "Bearer " & Authorization
        .setRequestHeader "User-Agent", UserAgent
        .setRequestHeader "Content-Type", "application/json"
        
    End With
        
    Dim Body As String, Message As String
    Body = "{" & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & Name & Chr(34)
    
    If Email <> "" Then
        Body = Body & "," & Chr(34) & "email" & Chr(34) & ":" & Chr(34) & Email & Chr(34)
    End If
    
    If JobTitle <> "" Then
        Body = Body & "," & Chr(34) & "job_title" & Chr(34) & ":" & Chr(34) & JobTitle & Chr(34)
    End If
    
    If DepartmentID <> 0 Then
        Dim DepartmentString As String
        DepartmentString = "{" & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & "" & Chr(34) & ","
        DepartmentString = DepartmentString & Chr(34) & "department_id" & Chr(34) & ":" & DepartmentID & "}"
        
        Body = Body & "," & Chr(34) & "department" & Chr(34) & ":" & DepartmentString
    End If
    
    If Notes <> "" Then
        Body = Body & "," & Chr(34) & "notes" & Chr(34) & ":" & Chr(34) & Notes & Chr(34)
    End If
    
    If AutoEmail Then
        Body = Body & "," & Chr(34) & "auto_email" & Chr(34) & ":" & Chr(34) & 1 & Chr(34)
    End If
    
    If Not FullTime Then
        Body = Body & "," & Chr(34) & "employee_type" & Chr(34) & ":" & Chr(34) & 0 & Chr(34)
        
        If Not WorkDaysHours Is Nothing Then
        
            If WorkDaysHours.Count <> 7 Then
                Message = "No people created." & vbNewLine & vbNewLine
                Message = Message & "WorkDaysHours must have exactly 7 items in it."
                MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
                Exit Sub
            End If
        
            Dim WorkDaysHoursString As String, Hours As Variant
            WorkDaysHoursString = "["
            
            For Each Hours In WorkDaysHours
                WorkDaysHoursString = WorkDaysHoursString & Hours & ","
            Next Hours
            
            WorkDaysHoursString = Left(WorkDaysHoursString, Len(WorkDaysHoursString) - 1)
            WorkDaysHoursString = WorkDaysHoursString & "]"
            
            Body = Body & "," & Chr(34) & "work_days_hours" & Chr(34) & ":" & Chr(34) & WorkDaysHoursString & Chr(34)
        
        End If
        
    End If
    
    If Not Active Then
        Body = Body & "," & Chr(34) & "active" & Chr(34) & ":" & Chr(34) & 0 & Chr(34)
    End If
    
    If Contractor Then
        Body = Body & "," & Chr(34) & "people_type_id" & Chr(34) & ":" & 2
    End If
    
    If Not Tags Is Nothing Then
    
        Dim TagsString As String
        TagsString = "["
            
        Dim t As Variant
        For Each t In Tags
            TagsString = TagsString & "{" & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & t & Chr(34) & ","
            TagsString = TagsString & Chr(34) & "type" & Chr(34) & ":" & 1 & "},"
        Next t
            
        TagsString = Left(TagsString, Len(TagsString) - 1)
        TagsString = TagsString & "]"
            
        Body = Body & "," & Chr(34) & "tags" & Chr(34) & ":" & TagsString
    
    End If
    
    If StartDate <> "" Then
        Body = Body & "," & Chr(34) & "start_date" & Chr(34) & ":" & Chr(34) & StartDate & Chr(34)
    End If
    
    If EndDate <> "" Then
        Body = Body & "," & Chr(34) & "end_date" & Chr(34) & ":" & Chr(34) & EndDate & Chr(34)
    End If
    
    If DefaultHourlyRate <> 0 Then
        Body = Body & "," & Chr(34) & "default_hourly_rate" & Chr(34) & ":" & Chr(34) & DefaultHourlyRate & Chr(34)
    End If
    
    Body = Body & "}"
    Request.send Body
    
    ' Invalid field
    If Request.Status = 422 Then
        Message = "No person created." & vbNewLine & vbNewLine
        Message = Message & "Float API responded with: 422 Unprocessable Entity - The data supplied has failed validation." & vbNewLine & vbNewLine
        Message = Message & "Things to check:" & vbNewLine
        Message = Message & " - If a DepartmentID was passed, verify that it is correct" & vbNewLine
        Message = Message & " - If any date parameters were passed, verify that they are valid dates and in the form DD-MM-YYYY, YY-MM-DD, or YYYY-MM-DD"
        MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
    End If

End Sub


Public Sub CreatePhase(Authorization As String, UserAgent As String, ProjectID As Long, Name As String, StartDate As String, _
    EndDate As String, Optional Color As String, Optional Notes As String, Optional BudgetTotal As Double, _
    Optional DefaultHourlyRate As Double, Optional Billable As Boolean = True, Optional Tentative As Boolean = False, _
    Optional Active As Boolean = True)

    ' Purpose:
    ' Create a new Phase for a Project on Float
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' ProjectID - the ID of the project on Float
    ' Name - the name of the phase
    ' StartDate - the date the project starts in one of the following forms: YYYY-MM-DD, YY-MM-DD, DD-MM-YYYY
    ' EndDate - the date the project ends in one of the following forms: YYYY-MM-DD, YY-MM-DD, DD-MM-YYYY
    ' Color - the hexidecimal color the phase; defaults to the project color if nothing is passed
    ' Notes - notes on the phase
    ' BudgetTotal - hours or currency this phase has alloted to it depending on which parameter the project uses
    ' DefaultHourlyRate - hourly rate of people working on this phase
    ' Billable - whether or not the phase is billable
    ' Tentative - whether of not the phase is tentative
    ' Active - whether or not the phase is active
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request

        .Open "POST", "https://api.float.com/v3/phases", False

        .setRequestHeader "Authorization", "Bearer " & Authorization
        .setRequestHeader "User-Agent", UserAgent
        .setRequestHeader "Content-Type", "application/json"
        
    End With
    
    Dim Body As String, Message As String
    Body = "{" & Chr(34) & "project_id" & Chr(34) & ":" & Chr(34) & ProjectID & Chr(34)
    Body = Body & "," & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & Name & Chr(34)
    Body = Body & "," & Chr(34) & "start_date" & Chr(34) & ":" & Chr(34) & StartDate & Chr(34)
    Body = Body & "," & Chr(34) & "end_date" & Chr(34) & ":" & Chr(34) & EndDate & Chr(34)
    
    If Color <> "" Then
        Body = Body & "," & Chr(34) & "color" & Chr(34) & ":" & Chr(34) & Color & Chr(34)
    End If
    
    If Notes <> "" Then
        Body = Body & "," & Chr(34) & "notes" & Chr(34) & ":" & Chr(34) & Notes & Chr(34)
    End If
    
    If BudgetTotal <> 0 Then
        Body = Body & "," & Chr(34) & "budget_total" & Chr(34) & ":" & Chr(34) & BudgetTotal & Chr(34)
    End If
    
    If DefaultHourlyRate <> 0 Then
        Body = Body & "," & Chr(34) & "default_hourly_rate" & Chr(34) & ":" & Chr(34) & DefaultHourlyRate & Chr(34)
    End If
    
    If Not Billable Then
        Body = Body & "," & Chr(34) & "non_billable" & Chr(34) & ":" & 1
    End If
    
    If Tentative Then
        Body = Body & "," & Chr(34) & "tentative" & Chr(34) & ":" & 1
    End If
    
    If Not Active Then
        Body = Body & "," & Chr(34) & "active" & Chr(34) & ":" & 0
    End If
    
    Body = Body & "}"
    Request.send Body
    
    ' Invalid field
    If Request.Status = 422 Then
    
        Message = "No phase created." & vbNewLine & vbNewLine
        Message = Message & "Float API responded with: 422 Unprocessable Entity - The data supplied has failed validation." & vbNewLine & vbNewLine
        Message = Message & Request.responseText
        
        MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
        
    End If

End Sub


Public Sub CreateProject(Authorization As String, UserAgent As String, Name As String, Optional ClientID As Long, Optional Color As String, _
    Optional Notes As String, Optional Tags As Collection, Optional BudgetType As Long = 0, Optional BudgetTotal As Double, _
    Optional DefaultHourlyRate As Double, Optional NonBillable As Boolean = False, Optional Tentative As Boolean = False, _
    Optional Active As Boolean = True, Optional ProjectManagerID As Long, Optional AllPMsSchedule As Boolean = True)

    ' Purpose:
    ' Create a new project on Float
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' Name - name of the project
    ' ClientID - the client_id of the project in Float
    ' Color - the hexidecimal color the project
    ' Notes - notes on the project
    ' Tags - any tags related to the project
    ' BudgetType - type of buget for the project
    ' BudgetTotal - total budget for project when BudgetType is either 2 or 3
    ' DefaultHourlyRate - default hourly rate of the project
    ' NonBillable - whether or not the project is billable
    ' Tentative - whether or not the project is tentative
    ' Active - whether the project is active or archived
    ' ProjectManagerID - the people_id of the project manager on Float
    ' AllPMsSchedule - whether or not all PMs have scheduling rights
    
    ' Parameter enumerations:
    ' BudgetType - 0=No budget (defualt), 1=Hours by project, 2=Fee by project, 3=Hourly fee
    
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request

        .Open "POST", "https://api.float.com/v3/projects", False

        .setRequestHeader "Authorization", "Bearer " & Authorization
        .setRequestHeader "User-Agent", UserAgent
        .setRequestHeader "Content-Type", "application/json"
 
        Dim Body As String
        Body = "{" & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & Name & Chr(34)
 
        If ClientID <> 0 Then
            Body = Body & "," & Chr(34) & "client_id" & Chr(34) & ":" & Chr(34) & ClientID & Chr(34)
        End If
        
        If Color <> "" Then
            Body = Body & "," & Chr(34) & "color" & Chr(34) & ":" & Chr(34) & Color & Chr(34)
        End If
        
        If Notes <> "" Then
            Body = Body & "," & Chr(34) & "notes" & Chr(34) & ":" & Chr(34) & Notes & Chr(34)
        End If
        
        If Not Tags Is Nothing Then
            
            Dim TagsString As String
            TagsString = "["
            
            Dim t As Variant
            For Each t In Tags
                TagsString = TagsString & Chr(34) & t & Chr(34) & ","
            Next t
            
            TagsString = Left(TagsString, Len(TagsString) - 1)
            TagsString = TagsString & "]"
            
            Body = Body & "," & Chr(34) & "tags" & Chr(34) & ":" & TagsString
        
        End If

        Body = Body & "," & Chr(34) & "budget_type" & Chr(34) & ":" & Chr(34) & BudgetType & Chr(34)
        
        If BudgetTotal <> 0 Then
            Body = Body & "," & Chr(34) & "budget_total" & Chr(34) & ":" & Chr(34) & BudgetTotal & Chr(34)
        End If
        
        If DefaultHourlyRate <> 0 Then
            Body = Body & "," & Chr(34) & "default_hourly_rate" & Chr(34) & ":" & Chr(34) & DefaultHourlyRate & Chr(34)
        End If
        
        If NonBillable Then
            Body = Body & "," & Chr(34) & "non_billable" & Chr(34) & ":" & Chr(34) & "1" & Chr(34)
        End If
        
        If Tentative Then
            Body = Body & "," & Chr(34) & "tentative" & Chr(34) & ":" & Chr(34) & "1" & Chr(34)
        End If
        
        If Not Active Then
            Body = Body & "," & Chr(34) & "active" & Chr(34) & ":" & Chr(34) & "0" & Chr(34)
        End If
        
        If ProjectManagerID <> 0 Then
            Body = Body & "," & Chr(34) & "project_manager" & Chr(34) & ":" & Chr(34) & ProjectManagerID & Chr(34)
        End If
        
        If AllPMsSchedule Then
            Body = Body & "," & Chr(34) & "all_pms_schedule" & Chr(34) & ":" & Chr(34) & "1" & Chr(34)
        End If
        
        Body = Body & "}"
        .send Body
        
        ' Invalid field
        If .Status = 422 Then
            Dim Message As String
            Message = "No project created.  Float API says: " & vbNewLine & vbNewLine & Split(.responseText, Chr(34))(7)
            MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
        End If
        
    End With
    
End Sub


Public Sub CreateTask(Authorization As String, UserAgent As String, ProjectID As Long, Hours As Double, PeopleIDs As Collection, _
    Optional Name As String, Optional PhaseID As Long, Optional StartDate As String, Optional EndDate As String, Optional StartTime As String, _
    Optional Status As Long = 2, Optional Notes As String, Optional RepeatState As Long = 0, Optional RepeatEndDate As String)

    ' Purpose:
    ' Create a new Task for a Project on Float
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' ProjectID - the ID of the project on Float
    ' Hours - number of hours per day
    ' PeopleIDs - all of the the people_ids on Float that this task will be assigned to
    ' Name - the name of the task
    ' PhaseID - the ID of the project phase on Float
    ' StartDate - the date the task starts in the form YYYY-MM-DD
    ' EndDate - the date the task ends in the form YYYY-MM-DD
    '         - if passed without a StartDate, the task will only span 1 day starting today
    ' StartTime - the 24hr time the task starts in the form HH:MM
    ' Status - status of the task
    ' Notes - notes on the task
    ' RepeatState - how often the task repeats
    ' RepeatEndDate - the last date that the task's start date can repeat to in the form YYYY-MM-DD
    '               - only required if RepeatState is not 0 (default)
    
    ' Parameter enumerations:
    ' Status - 1=Tentative, 2=Confirmed (defualt), 3=Complete
    ' RepeatState - 0=No repeat (defualt), 1=Weekly, 2=Monthly, 3=Every 2nd week, 4=Every 3rd week, 5=Every 6th week, 6=Every 2 months,
    '             - 7=Every 3 months, 8=Every 6 months, 9=Yearly
    
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request

        .Open "POST", "https://api.float.com/v3/tasks", False

        .setRequestHeader "Authorization", "Bearer " & Authorization
        .setRequestHeader "User-Agent", UserAgent
        .setRequestHeader "Content-Type", "application/json"
        
    End With
    
    Dim Body As String, Message As String
    Body = "{" & Chr(34) & "project_id" & Chr(34) & ":" & Chr(34) & ProjectID & Chr(34)
    Body = Body & "," & Chr(34) & "hours" & Chr(34) & ":" & Chr(34) & Hours & Chr(34)
    
    If Not PeopleIDs Is Nothing Then
    
        If PeopleIDs.Count = 1 Then
            Body = Body & "," & Chr(34) & "people_id" & Chr(34) & ":" & Chr(34) & PeopleIDs(1) & Chr(34)
            
        Else
            Dim PeopleIDsString As String
            PeopleIDsString = "["
            
            Dim p As Variant
            For Each p In PeopleIDs
                PeopleIDsString = PeopleIDsString & Chr(34) & p & Chr(34) & ","
            Next p
            
            PeopleIDsString = Left(PeopleIDsString, Len(PeopleIDsString) - 1)
            PeopleIDsString = PeopleIDsString & "]"
            
            Body = Body & "," & Chr(34) & "people_ids" & Chr(34) & ":" & PeopleIDsString
        
        End If
        
    Else
        MsgBox Prompt:="No task created.  PeopleIDs must have at least 1 member.", Buttons:=vbCritical, Title:="Bad Parameters"
        Exit Sub
        
    End If
    
    If Name <> "" Then
        Body = Body & "," & Chr(34) & "name" & Chr(34) & ":" & Chr(34) & Name & Chr(34)
    End If
    
    If PhaseID <> 0 Then
         Body = Body & "," & Chr(34) & "phase_id" & Chr(34) & ":" & Chr(34) & PhaseID & Chr(34)
    End If
    
    If StartDate <> "" And EndDate <> "" Then
        Body = Body & "," & Chr(34) & "start_date" & Chr(34) & ":" & Chr(34) & StartDate & Chr(34)
        Body = Body & "," & Chr(34) & "end_date" & Chr(34) & ":" & Chr(34) & EndDate & Chr(34)
    End If
    
    If StartTime <> "" Then
        Body = Body & "," & Chr(34) & "start_time" & Chr(34) & ":" & Chr(34) & StartTime & Chr(34)
    End If
    
    If Status = 1 Or Status = 2 Or Status = 3 Then
        Body = Body & "," & Chr(34) & "status" & Chr(34) & ":" & Chr(34) & Status & Chr(34)
    End If
    
    If Notes <> "" Then
        Body = Body & "," & Chr(34) & "notes" & Chr(34) & ":" & Chr(34) & Notes & Chr(34)
    End If
    
    If RepeatState <> 0 Then
        Body = Body & "," & Chr(34) & "repeat_state" & Chr(34) & ":" & Chr(34) & RepeatState & Chr(34)
        Body = Body & "," & Chr(34) & "repeat_end_date" & Chr(34) & ":" & Chr(34) & RepeatEndDate & Chr(34)
    End If
    
    Body = Body & "}"
    Request.send Body
    
    ' Invalid field
    If Request.Status = 422 Then
    
        Message = "No task created." & vbNewLine & vbNewLine
        Message = Message & "Float API responded with: 422 Unprocessable Entity - The data supplied has failed validation." & vbNewLine & vbNewLine
        Message = Message & Request.responseText
        
        MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
        
    End If

End Sub


Public Sub CreateTimeOff(Authorization As String, UserAgent As String, TimeOffTypeID As Long, StartDate As String, EndDate As String, _
    PeopleIDs As Collection, Optional FullDay As Boolean = True, Optional Hours As Double = 8, Optional StartTime As String, _
    Optional TimeOffNotes As String, Optional RepeatState As Long = 0, Optional RepeatEnd As String)

    ' Purpose:
    ' Create a new Time Off in Float
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' TimeOffTypeID - the time_off_id of this type of time off in Float
    ' StartDate - the start date of this time off in the form YYYY-MM-DD
    ' EndDate - the end date of this time off in the form YYYY-MM-DD
    ' PeopleIDs - all of the people_ids on Float of the people who have this time off
    ' FullDay - whether or not this time off lasts the entire day
    '         - if True, the Hours parameter is ignored
    ' Hours - number of hours per day for this time off
    ' StartTime - the 24hr time the time off starts in the form HH:MM
    ' TimeOffNotes - any notes on this time off
    ' RepeatState - how often this time off repeats
    ' RepeatEnd - the last date that the time off's start_date can repeat to in the form YYYY-MM-DD
    '           - only required if RepeatState is not 0 (default)
    
    ' Parameter enumerations:
    ' RepeatState - 0=No repeat (defualt), 1=Weekly, 2=Monthly, 3=Every 2nd week, 4=Every 3rd week, 5=Every 6th week, 6=Every 2 months,
    '             - 7=Every 3 months, 8=Every 6 months, 9=Yearly
    
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request

        .Open "POST", "https://api.float.com/v3/timeoffs", False

        .setRequestHeader "Authorization", "Bearer " & Authorization
        .setRequestHeader "User-Agent", UserAgent
        .setRequestHeader "Content-Type", "application/json"
        
    End With
    
    Dim Body As String, Message As String
    Body = "{" & Chr(34) & "timeoff_type_id" & Chr(34) & ":" & Chr(34) & TimeOffTypeID & Chr(34)
    Body = Body & "," & Chr(34) & "start_date" & Chr(34) & ":" & Chr(34) & StartDate & Chr(34)
    Body = Body & "," & Chr(34) & "end_date" & Chr(34) & ":" & Chr(34) & EndDate & Chr(34)
    
    Dim PeopleIDsString As String
    PeopleIDsString = "["
    
    Dim p As Variant
    For Each p In PeopleIDs
        PeopleIDsString = PeopleIDsString & Chr(34) & p & Chr(34) & ","
    Next p
    
    PeopleIDsString = Left(PeopleIDsString, Len(PeopleIDsString) - 1)
    PeopleIDsString = PeopleIDsString & "]"
    
    Body = Body & "," & Chr(34) & "people_ids" & Chr(34) & ":" & PeopleIDsString
    
    If FullDay Then
        Body = Body & "," & Chr(34) & "full_day" & Chr(34) & ":" & Chr(34) & 1 & Chr(34)
    
    Else
        Body = Body & "," & Chr(34) & "full_day" & Chr(34) & ":" & Chr(34) & 0 & Chr(34)
        Body = Body & "," & Chr(34) & "hours" & Chr(34) & ":" & Chr(34) & Hours & Chr(34)
        
    End If
    
    Body = Body & "," & Chr(34) & "start_time" & Chr(34) & ":" & Chr(34) & StartTime & Chr(34)
    Body = Body & "," & Chr(34) & "timeoff_notes" & Chr(34) & ":" & Chr(34) & TimeOffNotes & Chr(34)
    Body = Body & "," & Chr(34) & "repeat_state" & Chr(34) & ":" & Chr(34) & RepeatState & Chr(34)
    Body = Body & "," & Chr(34) & "repeat_end" & Chr(34) & ":" & Chr(34) & RepeatEnd & Chr(34)
    
    Body = Body & "}"
    Request.send Body
    
    ' Invalid field
    If Request.Status = 422 Then
    
        Message = "No time off created." & vbNewLine & vbNewLine
        Message = Message & "Float API responded with: 422 Unprocessable Entity - The data supplied has failed validation." & vbNewLine & vbNewLine
        Message = Message & Request.responseText
        
        MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
        
    End If

End Sub
        
        
Public Sub CreateTimeOffType(Authorization As String, UserAgent As String, TimeOffTypeName As String, Optional Color As String)

    ' Purpose:
    ' Create a new TimeOff Type on Float
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' TimeOffTypeName - the name of this timeoff type
    ' Color - the hexidecimal color the timeoff type
    
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request

        .Open "POST", "https://api.float.com/v3/timeoff-types", False

        .setRequestHeader "Authorization", "Bearer " & Authorization
        .setRequestHeader "User-Agent", UserAgent
        .setRequestHeader "Content-Type", "application/json"
        
    End With
    
    Dim Body As String, Message As String
    Body = "{" & Chr(34) & "timeoff_type_name" & Chr(34) & ":" & Chr(34) & TimeOffTypeName & Chr(34)
    Body = Body & "," & Chr(34) & "color" & Chr(34) & ":" & Chr(34) & Color & Chr(34)
    
    Body = Body & "}"
    Request.send Body
    
    ' Invalid field
    If Request.Status = 422 Then
    
        Message = "No time off created." & vbNewLine & vbNewLine
        Message = Message & "Float API responded with: 422 Unprocessable Entity - The data supplied has failed validation." & vbNewLine & vbNewLine
        Message = Message & Request.responseText
        
        MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
        
    End If

End Sub
                            
                            
Public Function GetPeople(Authorization As String, UserAgent As String, Optional PeopleID As Long = 0) As String

    ' Purpose:
    ' Return the information on all people in the team or a specific person
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' PeopleID - the people_id on Float of a single person
    '          - if none is passed then all people are returned

    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request
    
        Dim Response As String
    
        If PeopleID = 0 Then
        
            Dim Page As Long, ResponseLength As Long
            Page = 1
            
            ' Empty responses have a length of 2
            Do While ResponseLength <> 2
            
                .Open "GET", "https://api.float.com/v3/people?page=" & Page & "&per-page=200", False ' 200 per page max
            
                .setRequestHeader "Authorization", "Bearer " & Authorization
                .setRequestHeader "User-Agent", UserAgent
                .setRequestHeader "Content-Type", "application/json"
            
                .send
                
                Response = Response & .responseText
                ResponseLength = Len(.responseText)
                Page = Page + 1
            
            Loop
            
        Else
        
            .Open "GET", "https://api.float.com/v3/people/" & PeopleID, False ' 200 per page max
            
            .setRequestHeader "Authorization", "Bearer " & Authorization
            .setRequestHeader "User-Agent", UserAgent
            .setRequestHeader "Content-Type", "application/json"
        
            .send
            
            Response = .responseText
            
            If .Status = 404 Then
            
                Dim Message As String
                Message = "People not found.  Nobody on the team has the people_id of " & PeopleID & "."
                
                MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="No People"
                Exit Function
                
            End If
            
        End If
    
    End With
    
    GetPeople = Response

End Function
                                    
                                    
Public Function GetProjects(Authorization As String, UserAgent As String, Optional ProjectID As Long = 0) As String

    ' Purpose:
    ' Return the information on all projects for the team or a single project
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' PeopleID - the project_id on Float of a single project
    '          - if none is passed then all projects are returned
    
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request
    
        Dim Response As String
        
        If ProjectID = 0 Then
        
            Dim Page As Long, ResponseLength As Long
            Page = 1
            
            ' Empty responses have a length of 2
            Do While ResponseLength <> 2
            
                .Open "GET", "https://api.float.com/v3/projects?page=" & Page & "&per-page=200", False ' 200 per page max
            
                .setRequestHeader "Authorization", "Bearer " & Authorization
                .setRequestHeader "User-Agent", UserAgent
                .setRequestHeader "Content-Type", "application/json"
            
                .send
                
                Response = Response & .responseText
                ResponseLength = Len(.responseText)
                Page = Page + 1
            
            Loop
            
        Else
        
            .Open "GET", "https://api.float.com/v3/projects/" & ProjectID, False ' 200 per page max
            
            .setRequestHeader "Authorization", "Bearer " & Authorization
            .setRequestHeader "User-Agent", UserAgent
            .setRequestHeader "Content-Type", "application/json"
        
            .send
            
            Response = .responseText
            
            If .Status = 404 Then
            
                Dim Message As String
                Message = "Project not found.  No projects for the team have the project_id of " & ProjectID & "."
                
                MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="No Projects"
                Exit Function
                
            End If
            
        End If
        
    End With
    
    GetProjects = Response

End Function
                                            
                                            
Public Function GetDepartments(Authorization As String, UserAgent As String, Optional DepartmentID As Long = 0) As String

    ' Purpose:
    ' Return the information on all departments for the team or a single department
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' DepartmentID - the department_id on Float of a single department
    '              - if none is passed then all departments are returned
    
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request
    
        Dim Response As String
        
        If DepartmentID = 0 Then
        
            Dim Page As Long, ResponseLength As Long
            Page = 1
            
            ' Empty responses have a length of 2
            Do While ResponseLength <> 2
            
                .Open "GET", "https://api.float.com/v3/departments?page=" & Page & "&per-page=200", False ' 200 per page max
            
                .setRequestHeader "Authorization", "Bearer " & Authorization
                .setRequestHeader "User-Agent", UserAgent
                .setRequestHeader "Content-Type", "application/json"
            
                .send
                
                Response = Response & .responseText
                ResponseLength = Len(.responseText)
                Page = Page + 1
            
            Loop
            
        Else
        
            .Open "GET", "https://api.float.com/v3/departments/" & DepartmentID, False
            
            .setRequestHeader "Authorization", "Bearer " & Authorization
            .setRequestHeader "User-Agent", UserAgent
            .setRequestHeader "Content-Type", "application/json"
        
            .send
            
            Response = .responseText
            
            If .Status = 404 Then
            
                Dim Message As String
                Message = "Department not found.  No departments for the team have the department_id of " & DepartmentID & "."
                
                MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="No Departments"
                Exit Function
                
            End If
            
        End If
        
    End With
    
    GetDepartments = Response

End Function
                                                    
                                                    
Public Sub DeleteDepartment(Authorization As String, UserAgent As String, DepartmentID As Long)

    ' Purpose:
    ' Delete the department with the corresponding DepartmentID
    
    ' Parameters:
    ' Authorization - your unique API token provided by Float
    ' UserAgent - organization name and email address ex. "John's Bakery (John.Doe@Bakery.com)"
    ' DepartmentID - the department_id on Float of the department
    
    
    Dim Request As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
    
    With Request
    
        Dim Response As String
        
        .Open "DELETE", "https://api.float.com/v3/departments/" & DepartmentID, False
        
        .setRequestHeader "Authorization", "Bearer " & Authorization
        .setRequestHeader "User-Agent", UserAgent
        .setRequestHeader "Content-Type", "application/json"
    
        .send
        
        Response = .responseText
        
        If .Status = 404 Then
        
            Dim Message As String
            Message = "Department not found.  No departments for the team have the department_id of " & DepartmentID & "."
            
            MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="No Department"
            Exit Sub
            
        End If
            
    End With
    
End Sub
