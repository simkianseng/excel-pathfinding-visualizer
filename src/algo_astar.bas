Attribute VB_Name = "algo_astar"
Option Explicit


' astar algo abit weird uses priority stack instead of priority queue

' Function returns a Array contain 2 key items:
'   1. A Dictionary data structure containing all the cells explored and its parent cell.
'   2. A Collection data structure containing all the visited_cells.
Public Function astar(start_cell As String, target_cell As String) As Variant
    
    ' Main data structures for astar.
    Dim priority_queue As clsPriorityQueue
    Dim visited_cells As Collection
    Dim predecessors As Object
    Dim g_values As Object
    
    Dim current_cell As String
    Dim neighbour_cell As String
    Dim f_cost As Long
    Dim new_cost As Long
    
    Dim row_offset As Integer
    Dim column_offset As Integer
    
    Dim directions As Variant
    Dim direction_index As Byte
    
    Set priority_queue = New clsPriorityQueue
    Set visited_cells = New Collection
    Set predecessors = CreateObject("Scripting.Dictionary")
    Set g_values = CreateObject("Scripting.Dictionary")
    directions = directions_coll()
        
    priority_queue.enqueue 0, start_cell
    predecessors(start_cell) = Empty
    g_values(start_cell) = 0
    
    
    Do While priority_queue.Count > 0
        current_cell = priority_queue.dequeue2()
        visited_cells.Add (current_cell)
        If current_cell = target_cell Then
            Exit Do
        End If
        For direction_index = 0 To 3
            neighbour_cell = process_offset(current_cell, offset_cells(directions(direction_index)))
            new_cost = g_values(current_cell) + 1
            If legal_move(neighbour_cell) Then
                If Not g_values.Exists(neighbour_cell) Or new_cost < g_values(neighbour_cell) Then
                    g_values(neighbour_cell) = new_cost
                    f_cost = new_cost + calculate_heuristic(neighbour_cell, target_cell)
                    priority_queue.enqueue f_cost, neighbour_cell
                    predecessors(neighbour_cell) = current_cell
                End If
            End If
        Next direction_index
    Loop
    
    astar = Array(visited_cells, predecessors)
    
End Function

Function calculate_heuristic(current_cell As String, target_cell As String) As Long
    Dim h_cost As Long ' Distance of current cell from the target cell
    
    Dim current_cell_arr As Variant
    Dim target_cell_arr As Variant
    
    current_cell_arr = string_to_array(current_cell)
    target_cell_arr = string_to_array(target_cell)
    
    h_cost = Abs(current_cell_arr(0) - target_cell_arr(0)) + Abs(current_cell_arr(1) - target_cell_arr(1))
    
    calculate_heuristic = h_cost
    
End Function

Public Sub run_astar()
'Dim start_time As Long ' For benchmarking purposes
'start_time = Timer() ' For benchmarking purposes

    Dim start_cell As String
    Dim end_cell As String
    Dim visited_cells As Collection
    Dim predecessors As Object
    Dim astar_results As Variant
    
    On Error GoTo no_cell_found
    start_cell = find_cell_coordinates("A")
    end_cell = find_cell_coordinates("B")
    
    astar_results = astar(start_cell, end_cell)
    Set visited_cells = astar_results(0)
    Set predecessors = astar_results(1)
    
    If tools_form.mp1.Pages("p_advance").f_path.cbox_explored.Value Then
        Call show_visited_cells(visited_cells, tools_form.mp1.Pages("p_advance").f_path.tb_explored_delay.Value)
    End If
    If tools_form.mp1.Pages("p_advance").f_path.cbox_actual.Value Then
        Call show_path(predecessors, start_cell, end_cell, tools_form.mp1.Pages("p_advance").f_path.tb_actual_delay.Value)
    End If
    
    If predecessors(end_cell) = Empty Then
        MsgBox "No valid path found!"
    End If
    
'Debug.Print (Timer() - start_time) ' For benchmarking purposes
    
    Exit Sub
no_cell_found:
    MsgBox "Start point or end point is missing!"
      
End Sub
