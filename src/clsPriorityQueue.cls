VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPriorityQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Dim priority_queue As Object
Dim duplicate_values As Collection

Private Sub Class_Initialize()
    Set priority_queue = CreateObject("Scripting.Dictionary")
    Set duplicate_values = New Collection
End Sub

Public Sub enqueue(priority As Long, element As Variant)
    Dim coll As Collection
    If priority_queue.Exists(priority) Then
        priority_queue(priority).Add element
    Else
        Set coll = New Collection
        coll.Add element
        priority_queue.Add priority, coll
    End If
End Sub

Public Function dequeue()
' Retrieve element with the lowest priority/value.
    Dim priority As Long
    
    priority = Application.WorksheetFunction.Min(priority_queue.keys)
    
    dequeue = priority_queue(priority)(1)
    priority_queue(priority).Remove (1)
    
    If priority_queue(priority).Count = 0 Then
        priority_queue.Remove (priority)
    End If
End Function

Public Function dequeue2()

    Dim priority As Long
    
    priority = Application.WorksheetFunction.Min(priority_queue.keys)
    
    dequeue2 = priority_queue(priority)(priority_queue(priority).Count)
    priority_queue(priority).Remove (priority_queue(priority).Count)
    
    If priority_queue(priority).Count = 0 Then
        priority_queue.Remove (priority)
    End If
End Function

Public Property Get get_all_keys() As Object
    Set get_all_keys = priority_queue.keys
End Property

Public Property Get Count() As Long
    Count = priority_queue.Count
End Property
