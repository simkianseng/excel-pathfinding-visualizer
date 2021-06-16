Attribute VB_Name = "maze_generator"
Option Explicit


Sub generateMaze(starting_cell_address As String, maze_row As Long, maze_col As Long, obstacle_density As Double)
    Application.ScreenUpdating = False

    ' Set start cell
    Dim startingCell As Range
    Set startingCell = ActiveSheet.Range(starting_cell_address)
    
    ' Set last cell
    Dim finishCell As Range
    Set finishCell = startingCell.Offset(maze_row - 1, maze_col - 1)
       
    Dim mazeBorders As Range
    Set mazeBorders = ActiveSheet.Range(startingCell, finishCell)
    mazeBorders.Value = 1
    
    Dim mazeField As Range
    Set mazeField = ActiveSheet.Range(startingCell.Offset(1, 1), finishCell.Offset(-1, -1))
    mazeField.Value = Empty
    
    ' Create random walls inside maze area
    Dim vCell As Range
    Dim rng As Double
    For Each vCell In mazeField
        If Not vCell.Address = starting_cell_address And Not vCell.Address = finishCell.Address Then
            rng = Rnd()
            If rng > 1 - obstacle_density Then
                vCell.Value = 1
            End If
        End If
    Next
    Application.ScreenUpdating = True
End Sub

