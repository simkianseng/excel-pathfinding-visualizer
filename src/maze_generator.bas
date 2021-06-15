Attribute VB_Name = "maze_generator"
Option Explicit


Sub generateMaze(starting_cell_address As String, maze_row As Long, maze_col As Long)
    Application.ScreenUpdating = False
    Dim startingCell As Range
    Set startingCell = ActiveSheet.Range(starting_cell_address)
    
    ' SET LAST CELL
    Dim finishCell As Range
    Set finishCell = startingCell.Offset(0, maze_row - 1).Offset(maze_col - 1, 0)
       
    Dim mazeBorders As Range
'    Set mazeBorders = ActiveSheet.Range(startingCell.Offset(-1, -1), finishCell.Offset(1, 1))
    Set mazeBorders = ActiveSheet.Range(startingCell, finishCell)
    mazeBorders.Value = 1
    
    Dim mazeField As Range
    Set mazeField = ActiveSheet.Range(startingCell.Offset(1, 1), finishCell.Offset(-1, -1))
    mazeField.Value = Empty
    
    ' create random walls inside maze area
    Dim vCell As Range
    Dim rng As Double
    For Each vCell In mazeField
        If Not vCell.Address = starting_cell_address And Not vCell.Address = finishCell.Address Then
            rng = Rnd()
            If rng > 0.65 Then
                vCell.Value = 1
            End If
        End If
        
    Next
    Application.ScreenUpdating = True
End Sub

Sub clearMazeArea(starting_cell_address As String, maze_row As Long, maze_col As Long)
    Range(starting_cell_address).Offset(-1, 0).CurrentRegion.ClearContents
End Sub

