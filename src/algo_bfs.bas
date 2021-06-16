Attribute VB_Name = "algo_bfs"
Option Explicit


' Function returns a Array contain 2 key items:
'   1. A Dictionary data structure containing all the cells explored and its parent cell.
'   2. A Collection data structure containing all the visited_cells.
Public Function bfs(start_cell As String, target_cell As String) As Variant
    
    ' Main data structures for bfs.
    Dim queue As Object
    Dim visited_cells As Collection
    Dim predecessors As Object
    
    Dim current_cell As String
    Dim neighbour_cell As String
    
    Dim row_offset As Integer
    Dim column_offset As Integer
    
    Dim directions As Variant
    Dim direction_index As Byte
    
    Set queue = CreateObject("System.Collections.Queue")
    Set visited_cells = New Collection
    Set predecessors = CreateObject("Scripting.Dictionary")
    directions = directions_coll()
    
    queue.enqueue (start_cell)
    predecessors(start_cell) = Empty
    
    Do While queue.Count > 0
        current_cell = queue.dequeue()
        visited_cells.Add (current_cell)
        If current_cell = target_cell Then
            Exit Do
        End If
        For direction_index = 0 To 3
            neighbour_cell = process_offset(current_cell, offset_cells(directions(direction_index)))
            If Not predecessors.Exists(neighbour_cell) And legal_move(neighbour_cell) Then
                queue.enqueue (neighbour_cell)
                predecessors(neighbour_cell) = current_cell
            End If
        Next direction_index
    Loop
    
    bfs = Array(visited_cells, predecessors)
End Function

Public Sub run_bfs()
    Dim start_cell As String
    Dim end_cell As String
    Dim visited_cells As Collection
    Dim predecessors As Object
    Dim bfs_results As Variant
    
    On Error GoTo no_cell_found
    start_cell = find_cell_coordinates("A")
    end_cell = find_cell_coordinates("B")
    
    bfs_results = bfs(start_cell, end_cell)
    Set visited_cells = bfs_results(0)
    Set predecessors = bfs_results(1)
    
    If tools_form.mp1.Pages("p_advance").f_path.cbox_explored.Value Then
        Call show_visited_cells(visited_cells, tools_form.mp1.Pages("p_advance").f_path.tb_explored_delay.Value)
    End If
    If tools_form.mp1.Pages("p_advance").f_path.cbox_actual.Value Then
        Call show_path(predecessors, start_cell, end_cell, tools_form.mp1.Pages("p_advance").f_path.tb_actual_delay.Value)
    End If
    
    If predecessors(end_cell) = Empty Then
        MsgBox "No valid path found!"
    End If
    
    Exit Sub
no_cell_found:
    MsgBox "Start point or end point is missing!"
End Sub
