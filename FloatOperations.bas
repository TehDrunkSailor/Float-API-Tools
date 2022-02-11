Option Explicit

Public Sub CreatePerson(Authorization As String, UserAgent As String, Name As String, Optional Email As String, Optional JobTitle As String, _
    Optional DepartmentID As Long, Optional Notes As String, Optional AutoEmail As Boolean = False, Optional FullTime As Boolean = True, _
    Optional WorkDaysHours As Collection, Optional Active As Boolean = True, Optional Contractor As Boolean = False, _
    Optional Tags As Collection, Optional StartDate As String, Optional EndDate As String, Optional DefaultHourlyRate As Double)

    ' Purpose:
    ' Create a new Person on Float
    
    ' Notes:
    ' Date parameters must be in one of the following forms:
    ' - YYYY-MM-DD
    ' - YY-MM-DD
    ' - DD-MM-YYYY
    
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
        Message = "No people created." & vbNewLine & vbNewLine
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
    
    ' Notes:
    ' Date parameters must be in one of the following forms:
    ' - YYYY-MM-DD
    ' - YY-MM-DD
    ' - DD-MM-YYYY
    ' Color is a hexidecimal 6 character string which defaults to the project color if nothing is passed
    ' BudgetTotal is either hours or currency depending on which parameter the project uses
    
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
    Optional Notes As String, Optional Tags As Collection, Optional BudgetType As Long, Optional BudgetTotal As Double, _
    Optional DefaultHourlyRate As Double, Optional NonBillable As Boolean = False, Optional Tentative As Boolean = False, _
    Optional Active As Boolean = True, Optional ProjectManagerID As Long, Optional AllPMsSchedule As Boolean = False)

    ' Purpose:
    ' Create a new project on Float
    
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

        If BudgetType <> 0 Then
            Body = Body & "," & Chr(34) & "budget_type" & Chr(34) & ":" & Chr(34) & BudgetType & Chr(34)
        End If
        
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
            Message = "No projects created.  Float API says: " & vbNewLine & vbNewLine & Split(.responseText, Chr(34))(7)
            MsgBox Prompt:=Message, Buttons:=vbCritical, Title:="Bad Parameters"
        End If
        
    End With
    
End Sub




