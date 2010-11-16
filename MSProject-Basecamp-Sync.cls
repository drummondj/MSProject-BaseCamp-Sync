VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MSProject-Basecamp-Sync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' This module handles exporting tasks to basecamp

' BEGIN: EDIT THESE VARIABLES
Const SERVER_URL = "https://<your account>.basecamphq.com"
Const API_KEY = "<your API key account>"
Const PROJECT_NAME = "<your project name>"
' END: EDIT THESE VARIABLES

Dim PROJECT_ID As String
Dim PRIMARY_COMPANY_ID As String

Public Sub CleanTaskIDs()
   Dim T As Task
   For Each T In ActiveProject.Tasks
      SetMilestoneID T, ""
      SetTodoListID T, ""
      SetTodoItemID T, ""
   Next
End Sub
' Run this command to sync your project
Public Sub BasecampSync()
  PROJECT_ID = ""
  PRIMARY_COMPANY_ID = ""
  
  ' TODO - check we can talk to the server before we do anything
  '        just get the project name to see if it exists
  THING = GetProjectIDByName(PROJECT_NAME)
  
  If IsNull(THING) Then
    MsgBox ("Can't find project called '" & PROJECT_NAME & "'")
    Exit Sub
  End If
  
  Dim T As Task
  ' Search the projects tasks for milestones
  For Each T In ActiveProject.Tasks
    If T.Milestone = True Then
        ' If the task is a milestone then sync it
        CreateOrUpdateMilestone T
        
        ' Iterate through milestones sub-tasks and sync those too
        SyncSubTasks T, T
    End If
  Next
  
  MsgBox ("Basecamp Sync Complete")
End Sub

' This command resursively sync up sub-tasks and todo items for a given milestone (M)
Sub SyncSubTasks(T As Task, M As Task)
   Dim ST As Task
   ' Iterate through the specified tasks outline children
   For Each ST In T.OutlineChildren
      ' If this is not a milestone then process it
      If Not ST.Milestone Then
         If ST.OutlineChildren.Count = 0 Then
            ' If this task has no children then sync it to a todo_item
            CreateOrUpdateToDoItem ST, M
         Else
            ' Otherwise sync this task's children
            SyncSubTasks ST, M
         End If
      End If
   Next
End Sub
Sub CreateOrUpdateToDoItem(T As Task, M As Task)
    Dim Mode As String
    
    If GetTodoItemID(T) = "" Then
       Mode = "create"
       url = "/todo_lists/" & GetTodoListID(M) & "/todo_items.xml"
    Else
       Mode = "update"
       url = "/todo_items/" & GetTodoItemID(T) & ".xml"
    End If
        
    data = "<todo-item>" _
            & "<content>" & T.Name & "</content>" _
            & "<due-at>" & IsoDate(T.Finish) & "</due-at>"
        
    If Not T.ResourceNames = "" Then
        If InStr(T.ResourceNames, ",") > 0 Then
            MsgBox "Error, Basecamp cannot assign tasks to more than one person. Please update task " & T.Name
        Else
            PersonID = GetPersonIDByName(T.ResourceNames)
            If PersonID = "" Then
               MsgBox "Error, " & T.ResourceNames & "does not exist or is not assign to this project in Basecamp"
            Else
               data = data & "<responsible-party>" & PersonID & "</responsible-party>"
            End If
        End If
    Else
       data = data & "<responsible-party></responsible-party>"
    End If
    
    data = data & "<notify type=""boolean"">true</notify></todo-item>"
    
    If Mode = "create" Then
        Location = BasecampPostReturnLocation(url, data)
    Else
        Set Result = BasecampPut(url, data)
    End If
    
    If Mode = "create" Then
        SetTodoItemID T, Replace(Location, SERVER_URL & "/todo_items/", "")
    End If
    
    If T.PercentComplete = 100 Then
        CompleteToDoItem T
    Else
        UncompleteToDoItem T
    End If

End Sub
Sub CompleteToDoItem(T As Task)
   Set Result = BasecampPut("/todo_items/" & GetTodoItemID(T) & "/complete.xml", Null)
End Sub
Sub UncompleteToDoItem(T As Task)
   Set Result = BasecampPut("/todo_items/" & GetTodoItemID(T) & "/uncomplete.xml", Null)
End Sub
Sub CreateOrUpdateMilestone(T As Task)
    Dim Mode As String
    
    If GetMilestoneID(T) = "" Then
       Mode = "create"
       url = "/projects/" & GetProjectIDByName(PROJECT_NAME) & "/milestones/create"
    Else
       Mode = "update"
       url = "/milestones/update/" & GetMilestoneID(T)
    End If
    
    data = "<request><milestone>" _
      & "<title>" & T.Name & "</title>" _
      & "<deadline type=""date"">" & IsoDate(T.Finish) & "</deadline>" _
      & "<responsible-party>c" & GetPrimaryCompanyID() & "</responsible-party>" _
      & "<notify>true</notify>" _
      & "</milestone>"
    
    If Mode = "update" Then
       data = data & "<move-upcoming-milestones>false</move-upcoming-milestones>" _
                   & "<move-upcoming-milestones-off-weekends>false</move-upcoming-milestones-off-weekends>"
    End If
    
    data = data & "</request>"
        
    Set Result = BasecampPost(url, data)
    
    If Mode = "create" Then
        SetMilestoneID T, Result.SelectSingleNode("milestones/milestone/id").Text
    End If
    
    ' Mark this milestone as complete or not
    If T.PercentComplete = 100 Then
       url = "/milestones/complete/" & GetMilestoneID(T)
    Else
       url = "/milestones/uncomplete/" & GetMilestoneID(T)
    End If
    
    Set Result = BasecampGet(url)
    
    ' Create or update this milestones todo_list
    If GetTodoListID(T) = "" Then
        Mode = "create"
        url = "/projects/" & GetProjectIDByName(PROJECT_NAME) & "/todo_lists.xml"
    Else
        Mode = "update"
        url = "/todo_lists/" & GetTodoListID(T) & ".xml"
    End If
    
    data = "<todo-list>" _
        & "<name>" & T.Name & "</name>" _
        & "<description>To-Do items for " & T.Name & " milestone</description>" _
        & "<milestone-id>" & GetMilestoneID(T) & "</milestone-id>" _
        & "<private type=""boolean"">false</private>" _
        & "<tracked type=""boolean"">false</tracked>" _
        & "</todo-list>"
    
    If Mode = "create" Then
       Location = BasecampPostReturnLocation(url, data)
    Else
       Set Result = BasecampPut(url, data)
    End If
    
    If Mode = "create" Then
        SetTodoListID T, Replace(Location, SERVER_URL & "/todo_lists/", "")
    End If
    
End Sub

Function BasecampPost(ByVal url As String, ByVal data)
    Set BasecampPost = BasecampPostOrPut("POST", url, data)
End Function

Function BasecampPut(ByVal url As String, ByVal data)
    Set BasecampPut = BasecampPostOrPut("PUT", url, data)
End Function

Function BasecampPostOrPut(verb As String, ByVal url As String, ByVal data)
    Set xmlhttp = CreateObject("Microsoft.XMLHTTP")
    xmlhttp.Open verb, SERVER_URL + url, False, API_KEY, "X"
    xmlhttp.setRequestHeader "Content-Type", "application/xml"
    xmlhttp.setRequestHeader "Accept", "application/xml"
    xmlhttp.send data
    Sleep
    Set xmlDoc = CreateObject("Microsoft.XMLDOM")
    xmlDoc.async = False
    xmlDoc.LoadXML (xmlhttp.responseText)
    Sleep
    Set xmlhttp = Nothing
    Set BasecampPostOrPut = xmlDoc
End Function

Function BasecampPostReturnLocation(ByVal url As String, ByVal data)
    Set xmlhttp = CreateObject("Microsoft.XMLHTTP")
    xmlhttp.Open "POST", SERVER_URL + url, False, API_KEY, "X"
    xmlhttp.setRequestHeader "Content-Type", "application/xml"
    xmlhttp.setRequestHeader "Accept", "application/xml"
    xmlhttp.send data
    Sleep
    BasecampPostReturnLocation = xmlhttp.getResponseHeader("Location")
    Set xmlhttp = Nothing
End Function
Function BasecampGet(ByVal url As String)
    Set xmlhttp = CreateObject("Microsoft.XMLHTTP")
    xmlhttp.Open "GET", SERVER_URL + url, False, API_KEY, "X"
    xmlhttp.setRequestHeader "Content-Type", "application/xml"
    xmlhttp.setRequestHeader "Accept", "application/xml"
    xmlhttp.send
    Sleep
    Set xmlDoc = CreateObject("Microsoft.XMLDOM")
    xmlDoc.async = False
    xmlDoc.LoadXML (xmlhttp.responseText)
    Set xmlhttp = Nothing
    Set BasecampGet = xmlDoc
End Function

Function GetPrimaryCompanyID()
   If PRIMARY_COMPANY_ID = "" Then
      Set doc = BasecampGet("/account.xml")
       PRIMARY_COMPANY_ID = doc.SelectSingleNode("account/primary-company-id").Text
   End If
   GetPrimaryCompanyID = PRIMARY_COMPANY_ID
End Function

Function GetProjectIDByName(ByVal projectName As String)
   If PROJECT_ID = "" Then
      Set doc = BasecampGet("/projects.xml")
      For Each Node In doc.SelectNodes("projects/project")
         THING = Node.SelectSingleNode("id").Text
         ThisName = Node.SelectSingleNode("name").Text
         If ThisName = projectName Then
            PROJECT_ID = THING
            Exit For
         End If
      Next
   End If
   GetProjectIDByName = PROJECT_ID
End Function
Function GetPersonIDByName(ByVal personName As String)
      Set doc = BasecampGet("/projects/" & GetProjectIDByName(PROJECT_NAME) & "/people.xml")
      For Each Node In doc.SelectNodes("people/person")
         THING = Node.SelectSingleNode("id").Text
         ThisName = Node.SelectSingleNode("first-name").Text & " " & Node.SelectSingleNode("last-name").Text
         If ThisName = personName Then
            GetPersonIDByName = THING
            Exit For
         End If
      Next
End Function
Function IsoDate(ByVal dt)
   Dim y: y = Year(dt)
   Dim M: M = Month(dt)
   Dim d: d = Day(dt)
   Dim h: h = Hour(dt)
   Dim n: n = Minute(dt)
   Dim s: s = Second(dt)

   If M < 10 Then M = "0" & M
   If d < 10 Then d = "0" & d
   If h < 10 Then h = "0" & h
   If n < 10 Then n = "0" & n
   If s < 10 Then s = "0" & s
  
   IsoDate = y & "-" & M & "-" & d & "T" & h & ":" & n & ":" & s
End Function

Function GetMilestoneID(T As Task)
   GetMilestoneID = T.Text1
End Function

Sub SetMilestoneID(T As Task, ByVal THING As String)
   T.Text1 = THING
End Sub

Function GetTodoListID(T As Task)
   GetTodoListID = T.Text2
End Function

Sub SetTodoListID(T As Task, ByVal THING As String)
   T.Text2 = THING
End Sub

Function GetTodoItemID(T As Task)
   GetTodoItemID = T.Text3
End Function

Sub SetTodoItemID(T As Task, ByVal THING As String)
   T.Text3 = THING
End Sub

Sub Sleep()
    StartTime = Timer    ' Set start time.
    Do While Timer < StartTime + 0.5
    Loop
End Sub



