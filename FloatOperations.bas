Attribute VB_Name = "FloatOperations"
Option Explicit

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

