Attribute VB_Name = "algo_greedy_bfs"
Option Explicit


Public Function greedy_bfs(start_cell As String, target_cell As String) As Variant
    
    ' Main data structures for greedy_bfs.
    Dim priority_queue As clsPriorityQueue
    Dim visited_cells As Collection
    Dim predecessors As Object
    
    Dim current_cell As String
    Dim neighbour_cell As String
    Dim neighbour_priority As Long
    
    Dim new_cost As Long
    
    Dim row_offset As Integer
    Dim column_offset As Integer
    
    Dim directions As Variant
    Dim direction_index As Byte
    
    Dim x As Integer
    x = 0
    
    Set priority_queue = New clsPriorityQueue
    Set visited_cells = New Collection
    Set predecessors = CreateObject("Scripting.Dictionary")
    directions = directions_coll()
        
    priority_queue.enqueue 0, start_cell
    predecessors(start_cell) = Empty
    
    
    Do While priority_queue.Count > 0
        current_cell = priority_queue.dequeue()
        visited_cells.Add (current_cell)
        If current_cell = target_cell Then
            Exit Do
        End If
        For direction_index = 0 To 3
            neighbour_cell = process_offset(current_cell, offset_cells(directions(direction_index)))
            If legal_move(neighbour_cell) Then
                If Not predecessors.Exists(neighbour_cell) Then
                    neighbour_priority = calculate_heuristic(neighbour_cell, target_cell)
                    priority_queue.enqueue neighbour_priority, neighbour_cell
                    predecessors(neighbour_cell) = current_cell
                End If
            End If
        Next direction_index
    Loop
    
    greedy_bfs = Array(visited_cells, predecessors)
    
End Function

Function calculate_heuristic(current_cell As String, target_cell As String) As Long
    Dim h_cost As Long  ' Distance of current cell from the target cell
    
    Dim current_cell_arr As Variant
    Dim target_cell_arr As Variant
    
    current_cell_arr = string_to_array(current_cell)
    target_cell_arr = string_to_array(target_cell)
    
    h_cost = Abs(current_cell_arr(0) - target_cell_arr(0)) + Abs(current_cell_arr(1) - target_cell_arr(1))
    
    calculate_heuristic = h_cost
    
End Function

Public Sub run_greedy_bfs()
    Dim start_cell As String
    Dim end_cell As String
    Dim visited_cells As Collection
    Dim predecessors As Object
    Dim greedy_bfs_results As Variant
    
    On Error GoTo no_cell_found
    start_cell = find_cell_coordinates("A")
    end_cell = find_cell_coordinates("B")
    
    greedy_bfs_results = greedy_bfs(start_cell, end_cell)
    Set visited_cells = greedy_bfs_results(0)
    Set predecessors = greedy_bfs_results(1)
    
    If tools_form.mp1.Pages("p_advance").f_path.cbox_explored.Value Then
        Call show_visited_cells(visited_cells, tools_form.mp1.Pages("p_advance").f_path.tb_explored_delay.Value)
    End If
    If tools_form.mp1.Pages("p_advance").f_path.cbox_actual.Value Then
        Call show_path(predecessors, start_cell, end_cell, tools_form.mp1.Pages("p_advance").f_path.tb_actual_delay.Value)
    End If
    
'    If predecessors(end_cell) = Empty Then
'        MsgBox "No valid path found!"
'    End If
    
    Exit Sub
no_cell_found:
    MsgBox "Start point or end point is missing!"
      
'      Call show_visited_cells(visited_cells, 0)
'      Call show_path(predecessors, start_cell, end_cell, 0)
'
End Sub
