VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub btn_show_tools_Click()
    tools_form.Show vbModeless
    Range(Cells(ActiveWindow.ScrollRow, ActiveWindow.ScrollColumn).Address).Select
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.EnableEvents = False
    With tools_form
        If .tgl_eraser.Value Then
            Target.Value = Empty
        ElseIf .tgl_obstacles.Value Then
            Target.Value = 1
        ElseIf .tgl_start.Value Then
            If Target.Count > 1 Then
                MsgBox "Can only set 1 start point!"
            Else
                On Error Resume Next
                Cells(string_to_array(find_cell_coordinates("A"))(0), string_to_array(find_cell_coordinates("A"))(1)).Value = Empty
                Target.Value = "A"
            End If
        ElseIf .tgl_target.Value Then
            If Target.Count > 1 Then
                MsgBox "Can only set 1 start point!"
            Else
                On Error Resume Next
                Cells(string_to_array(find_cell_coordinates("B"))(0), string_to_array(find_cell_coordinates("B"))(1)).Value = Empty
                Target.Value = "B"
            End If
        End If
    End With
    
    Range(Cells(ActiveWindow.ScrollRow, ActiveWindow.ScrollColumn).Address).Select
    Application.EnableEvents = True
End Sub

Sub disable_setup()
    Application.EnableEvents = False
    Application.CutCopyMode = True
    ActiveSheet.EnableSelection = xlNoRestrictions
End Sub

Sub setup()
    Application.EnableEvents = True
End Sub
