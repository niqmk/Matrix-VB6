Attribute VB_Name = "mdlLU"
Option Explicit

Public Sub CalculateLU( _
    ByRef dblSource() As Double, _
    ByRef dblResultUpper() As Double, _
    ByRef dblResultLower() As Double, _
    ByVal intMatrixMode As Integer)
    Dim dblPort() As Double
    Dim dblLower() As Double
    Dim dblUpper() As Double
    
    ReDim dblPort(intMatrixMode - 1) As Double
    ReDim dblLower(intMatrixMode - 1, intMatrixMode - 1) As Double
    ReDim dblUpper(intMatrixMode - 1, intMatrixMode - 1) As Double
    
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim intColumn As Integer
    
    Dim dblPosPort As Double
    Dim dblPosSource As Double
    
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            dblUpper(intRow, intColumn) = dblSource(intRow, intColumn)
            
            dblLower(intRow, intColumn) = 0
        Next intColumn
    Next intRow
    
    For intLoop = 0 To intMatrixMode - 1
        For intRow = intLoop To intMatrixMode - 1
            For intColumn = 0 To intMatrixMode - 1
                If intRow = intLoop Then
                    dblPort(intColumn) = dblUpper(intRow, intColumn)
                Else
                    If intColumn = intLoop Then
                        dblPosSource = dblUpper(intRow, intColumn)
                    End If
                    
                    dblPosPort = dblPort(intLoop)
                    
                    dblUpper(intRow, intColumn) = DivideElement(dblUpper(intRow, intColumn), dblPort(intColumn), dblPosSource, dblPosPort)
                    
                    If dblPosPort = 0 Then
                        dblLower(intRow, intColumn) = 0
                    Else
                        dblLower(intRow, intColumn) = dblPosSource / dblPosPort
                    End If
                End If
            Next intColumn
        Next intRow
    Next intLoop
    
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            If intRow = intColumn Then
                dblLower(intRow, intColumn) = 1
            ElseIf intRow < intColumn Then
                dblLower(intRow, intColumn) = 0
            End If
        Next intColumn
    Next intRow
    
    dblResultUpper = dblUpper
    dblResultLower = dblLower
End Sub

Private Function DivideElement( _
    ByVal dblSource As Double, _
    ByVal dblMultiple As Double, _
    ByVal dblPosSource As Double, _
    ByVal dblPosMultiple As Double) As Double
    If dblPosSource = 0 Then
        DivideElement = 0
    Else
        DivideElement = dblSource - (dblMultiple * (dblPosSource / dblPosMultiple))
    End If
End Function
