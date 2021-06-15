Attribute VB_Name = "dependencies"
Option Explicit


'~~~~~~~~~~~~~~~~~~~~~~~'
' Made by Sim Kian Seng '
'~~~~~~~~~~~~~~~~~~~~~~~'

' Programming concepts applied in this project:
'   1. Array manipulation
'   2. String manipulation
'   3. Dictionary manipulation
'   4. Application Programming Interface (API, Windows API used)
'   5. Functions (returns object/value)
'   6. Subroutine (Do not return anything)
'   7. Userforms
'   8. Classes and objects


' Here contain the dependent functions/sub procedures.


Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Public Function legal_move(cell As String) As Boolean
    Dim row_index As Integer
    Dim col_index As Integer
    
    row_index = string_to_array(cell)(0)
    col_index = string_to_array(cell)(1)
    
    If row_index < 1 Or col_index < 1 Then
        legal_move = False
        Exit Function
    End If
    If Cells(row_index, col_index).Value <> 1 Then
        legal_move = True
    Else
        legal_move = False
    End If
End Function

Public Function directions_coll() As Variant
    directions_coll = Array("right", "down", "left", "up")
End Function

Public Function offset_cells(direct As Variant) As Variant
    If direct = "right" Then
        offset_cells = Array(0, 1)
    ElseIf direct = "down" Then
        offset_cells = Array(1, 0)
    ElseIf direct = "left" Then
        offset_cells = Array(0, -1)
    ElseIf direct = "up" Then
        offset_cells = Array(-1, 0)
    End If
End Function

Public Function process_offset(cell As String, offset_cells As Variant) As String
' This function converts cell from string to integer and process the value.
    Dim row_index As Integer
    Dim col_index As Integer
    
    row_index = string_to_array(cell)(0) + offset_cells(0)
    col_index = string_to_array(cell)(1) + offset_cells(1)
    process_offset = row_index & "," & col_index
End Function

Public Function string_to_array(cell As Variant) As Variant
    Dim arr As Variant
    
    arr = Split(cell, ",")
    string_to_array = Array(CLng(arr(0)), CLng(arr(1)))
End Function

Public Sub show_discovered_cells(predecessors As Object, Optional delay As Integer)
    Dim item As Variant
    
    For Each item In predecessors
        Cells(string_to_array(item)(0), string_to_array(item)(1)).Interior.Color = vbCyan
        Sleep (delay)
    Next item
End Sub

Public Sub show_visited_cells(visited_cells As Object, Optional delay As Integer)
    Dim visited_cell As Variant
    
    For Each visited_cell In visited_cells
        Cells(string_to_array(visited_cell)(0), string_to_array(visited_cell)(1)).Interior.Color = vbCyan
        Sleep (delay)
    Next visited_cell
End Sub

Public Sub show_path(predecessors As Object, start_cell As String, target_cell As String, Optional delay As Integer)
    Dim prev_cell As String

    Cells(string_to_array(target_cell)(0), string_to_array(target_cell)(1)).Interior.Color = vbYellow
    prev_cell = predecessors(target_cell)
    Do While prev_cell <> Empty
        Sleep (delay)
        Cells(string_to_array(prev_cell)(0), string_to_array(prev_cell)(1)).Interior.Color = vbYellow
        prev_cell = predecessors(prev_cell)
    Loop
End Sub

Public Function find_cell_coordinates(cell_value) As String
    With Range("A:XFD").Find(cell_value, LookIn:=xlValues)
        find_cell_coordinates = .Row & "," & .Column
    End With
End Function

Public Sub clear_path()
    Cells.Interior.Color = xlNone
End Sub
