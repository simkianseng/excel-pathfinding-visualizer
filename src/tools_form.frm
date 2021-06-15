VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} tools_form 
   Caption         =   "Tools"
   ClientHeight    =   7080
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3876
   OleObjectBlob   =   "tools_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "tools_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btn_generate_maze_Click()
    If tb_maze_row.Value < 2 Or tb_maze_col.Value < 2 Then
        MsgBox "Please generate a maze that is 2 by 2 or greater!", vbCritical
        Exit Sub
    End If
    
    Call generateMaze(tb_cell_reference.Value, tb_maze_row.Value, tb_maze_col.Value)
End Sub

Private Sub tgl_eraser_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    tgl_obstacles.Value = False
    tgl_start.Value = False
    tgl_target.Value = False
End Sub

Private Sub tgl_obstacles_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    tgl_eraser.Value = False
    tgl_start.Value = False
    tgl_target.Value = False
End Sub

Private Sub tgl_start_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    tgl_eraser.Value = False
    tgl_obstacles.Value = False
    tgl_target.Value = False
End Sub

Private Sub tgl_target_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    tgl_eraser.Value = False
    tgl_obstacles.Value = False
    tgl_start.Value = False
End Sub

Private Sub btn_clear_Click()
    Call deact_btn
    Call clear_path
End Sub

Private Sub btn_run_Click()
    Call deact_btn
    Call clear_path
    
    If ob_bfs Then
        Call run_bfs
    ElseIf ob_dfs Then
        Call run_dfs
    ElseIf ob_greedy_bfs Then
        Call run_greedy_bfs
    ElseIf ob_astar Then
        Call run_astar
    End If
End Sub

Private Sub UserForm_Activate()
    Me.StartUpPosition = 0
    'Upper left
    Dim Top As Double, Left As Double
'    Top = Abs(Application.Top) + (Application.Height - ActiveWindow.Height) + (Application.UsableHeight - ActiveWindow.UsableHeight)
'    Left = Abs(Application.Left) + ActiveWindow.Width - ActiveWindow.UsableWidth
    Me.Top = ActiveSheet.Range("A8").Top
    Me.Left = ActiveWindow.Left
End Sub

Private Sub UserForm_Initialize()
    Dim algo_type As String
    
    With Me
        .mp1.Pages("p_advance").f_path.cbox_explored.Value = Sheet3.Range("B1").Value
        .mp1.Pages("p_advance").f_path.cbox_actual.Value = Sheet3.Range("B2").Value
        .mp1.Pages("p_advance").f_path.tb_explored_delay.Value = Sheet3.Range("B3").Value
        .mp1.Pages("p_advance").f_path.tb_actual_delay.Value = Sheet3.Range("B4").Value
        
        .mp1.Pages("p_basic").f_maze_generator.tb_cell_reference.Value = Sheet3.Range("B6").Value
        .mp1.Pages("p_basic").f_maze_generator.tb_maze_row.Value = Sheet3.Range("B7").Value
        .mp1.Pages("p_basic").f_maze_generator.tb_maze_col.Value = Sheet3.Range("B8").Value
        
        algo_type = Sheet3.Range("B5").Value
        If algo_type = "bfs" Then
            .mp1.Pages("p_basic").f_algo.ob_bfs.Value = True
        ElseIf algo_type = "dfs" Then
            .mp1.Pages("p_basic").f_algo.ob_dfs.Value = True
        ElseIf algo_type = "greedy_bfs" Then
            .mp1.Pages("p_basic").f_algo.ob_greedy_bfs.Value = True
        ElseIf algo_type = "astar" Then
            .mp1.Pages("p_basic").f_algo.ob_astar.Value = True
        End If
    End With
End Sub

Private Sub UserForm_Terminate()
    With Me
        Sheet3.Range("B1").Value = .mp1.Pages("p_advance").f_path.cbox_explored.Value
        Sheet3.Range("B2").Value = .mp1.Pages("p_advance").f_path.cbox_actual.Value
        Sheet3.Range("B3").Value = .mp1.Pages("p_advance").f_path.tb_explored_delay.Value
        Sheet3.Range("B4").Value = .mp1.Pages("p_advance").f_path.tb_actual_delay.Value
        
        Sheet3.Range("B6").Value = .mp1.Pages("p_basic").f_maze_generator.tb_cell_reference.Value
        Sheet3.Range("B7").Value = .mp1.Pages("p_basic").f_maze_generator.tb_maze_row.Value
        Sheet3.Range("B8").Value = .mp1.Pages("p_basic").f_maze_generator.tb_maze_col.Value
        
        If .mp1.Pages("p_basic").f_algo.ob_bfs.Value = True Then
            Sheet3.Range("B5").Value = "bfs"
        ElseIf .mp1.Pages("p_basic").f_algo.ob_dfs.Value = True Then
            Sheet3.Range("B5").Value = "dfs"
        ElseIf .mp1.Pages("p_basic").f_algo.ob_greedy_bfs.Value = True Then
            Sheet3.Range("B5").Value = "greedy_bfs"
        ElseIf .mp1.Pages("p_basic").f_algo.ob_astar.Value = True Then
            Sheet3.Range("B5").Value = "astar"
        End If
    End With
End Sub

Sub deact_btn()
    tgl_eraser.Value = False
    tgl_obstacles.Value = False
    tgl_start.Value = False
    tgl_target.Value = False
End Sub
