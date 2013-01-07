Attribute VB_Name = "mdlPivot"
Option Explicit

Public Sub CalculatePivot(ByRef dblSource() As Double, ByRef dblResult() As Double, ByVal intMatrixMode As Integer)
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim intColumn As Integer
    
    Dim dblPort() As Double
    
    ReDim dblPort(intMatrixMode - 1, intMatrixMode - 1) As Double
    
    For intLoop = 0 To intMatrixMode - 1
        For intRow = intLoop + 1 To intMatrixMode - 1
            If dblSource(intLoop, intLoop) = 0 Then
                dblPort(intRow, intLoop) = 0
            Else
                dblPort(intRow, intLoop) = dblSource(intRow, intLoop) / dblSource(intLoop, intLoop)
            End If
            
            dblSource(intRow, intLoop) = 0
            
            For intColumn = intLoop + 1 To intMatrixMode - 1
                dblSource(intRow, intColumn) = dblSource(intRow, intColumn) - dblPort(intRow, intLoop) * dblSource(intLoop, intColumn)
            Next intColumn
        Next intRow
    Next intLoop
    
    dblResult = dblSource
End Sub
