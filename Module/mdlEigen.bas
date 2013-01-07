Attribute VB_Name = "mdlEigen"
Option Explicit

'Private Const dblEpsilon As Double = 0.01

Private dblEigenValues() As Double
Private dblEigenVectors() As Double

Public Sub CalculateEigen( _
    ByRef dblSource() As Double, _
    ByVal intMatrixMode As Integer)
    Dim strPolinomial() As String
    
    ReDim dblEigenValues(intMatrixMode - 1) As Double
    ReDim dblEigenVectors(intMatrixMode - 1, intMatrix - 1) As Double
    
    ReDim strPolinomial(intMatrixMode) As String
    
    Dim intCounter As Integer
    
    For intCounter = 0 To intMatrixMode
        If intCounter = 0 Then
            strPolinomial(intCounter) = CStr((-1) ^ intMatrixMode) & "(" & Chr(182) & "^" & CStr(intMatrixMode) & ")"
        ElseIf intCounter = 1 Then
            strPolinomial(intCounter) = CStr((-1) ^ (intMatrixMode - 1)) & "*" & CStr(mdlProcedures.TraceMatrix(dblSource, intMatrixMode)) & "(" & Chr(182) & "^" & CStr(intMatrixMode - 1) & ")"
        ElseIf intCounter = intMatrixMode Then
            strPolinomial(intCounter) = CStr(mdlProcedures.GetDeterminant(dblSource, intMatrixMode))
        End If
    Next intCounter
    
    mdlProcedures.MessageBoxPolinomial strPolinomial
    
    Dim strBreak() As String
    
    strBreak = mdlProcedures.BreakPolinomial(strPolinomial, intMatrixMode)
    
    For intCounter = 0 To intMatrixMode - 1
        dblEigenValues(intCounter) = CDbl(strBreak(intCounter))
    Next intCounter
    
    CalculateEigenVectors dblSource, intMatrixMode
End Sub

Private Sub CalculateEigenVectors(ByRef dblSource() As Double, ByVal intMatrixMode As Integer)
    Dim dblX() As Double
    
    ReDim dblX(intMatrixMode - 1, intMatrixMode - 1, intMatrixMode) As Double
    
    Dim intCounter As Integer
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intCounter = 0 To intMatrixMode - 1
        For intRow = 0 To intMatrixMode - 1
            For intColumn = 0 To intMatrixMode - 1
                If intRow = intColumn Then
                    dblX(intCounter, intRow, intColumn) = dblSource(intRow, intColumn) - dblEigenValues(intCounter)
                Else
                    dblX(intCounter, intRow, intColumn) = dblSource(intRow, intColumn)
                End If
            Next intColumn
        Next intRow
    Next intCounter
    
    Dim dblResult() As Double
    Dim dblResult2() As Double
    Dim dblTemp() As Double
    
    For intCounter = 0 To intMatrixMode - 1
        ReDim dblTemp(intMatrixMode - 1, intMatrixMode - 1) As Double
        
        For intRow = 0 To intMatrixMode - 1
            For intColumn = 0 To intMatrixMode - 1
                dblTemp(intRow, intColumn) = dblX(intCounter, intRow, intColumn)
            Next intColumn
        Next intRow
        
        ReDim dblResult(intMatrixMode - 1) As Double
        
        dblResult2 = mdlProcedures.ReducedEchelonMatrix(dblTemp, dblResult, intMatrixMode)
        
        For intRow = 0 To intMatrixMode - 1
            dblEigenVectors(intCounter, intRow) = dblResult(intRow)
        Next intRow
    Next intCounter
End Sub

'Public Sub CalculateEigen( _
'    ByRef dblSource() As Double, _
'    ByVal intMatrixMode As Integer)
'    Dim dblSourceTemp() As Double
'    Dim dblX() As Double
'    Dim dblY As Double
'    Dim dblXT() As Double
'    Dim dblZ() As Double
'    Dim dblTemp() As Double
'
'    ReDim dblEigenVectors(intMatrixMode - 1, intMatrixMode - 1) As Double
'
'    dblSourceTemp = mdlProcedures.CopyMatrix(dblSource, intMatrixMode)
'
'    ReDim dblEigenValues(intMatrixMode - 1) As Double
'
'    Dim intRow As Integer
'    Dim intColumn As Integer
'
'    For intColumn = 0 To intMatrixMode - 1
'        CalculatePower dblSourceTemp, dblX, dblY, intMatrixMode
'
'        dblEigenValues(intColumn) = dblY
'
'        For intRow = 0 To intMatrixMode - 1
'            dblEigenVectors(intRow, intColumn) = dblX(intRow, 0)
'        Next intRow
'
'        dblXT = mdlProcedures.SetTranspose2(dblX)
'        dblTemp = mdlProcedures.MultiplyMatrix2(dblX, dblXT)
'        dblZ = mdlProcedures.MultiplyScalar2(dblTemp, dblY, UBound(dblTemp, 1), UBound(dblTemp, 2))
'        dblSourceTemp = mdlProcedures.SubtractMatrix(dblSourceTemp, dblZ)
'    Next intColumn
'End Sub
'
'Private Sub CalculatePower( _
'    ByRef dblSource() As Double, _
'    ByRef dblX() As Double, _
'    ByRef dblY As Double, _
'    ByVal intMatrixMode As Integer)
'    ReDim dblX(intMatrixMode - 1, 0)
'
'    Dim intCounter As Integer
'
'    For intCounter = 0 To intMatrixMode - 1
'        dblX(intCounter, 0) = 1#
'    Next intCounter
'
'    Dim dblYTemp() As Double
'    Dim dblMax() As Double
'
'    Dim blnLoop As Boolean
'
'    blnLoop = True
'
'    intCounter = -1
'
'    While blnLoop
'        intCounter = intCounter + 1
'
'        ReDim Preserve dblMax(intCounter) As Double
'
'        dblYTemp = mdlProcedures.MultiplyMatrix2(dblSource, dblX)
'
'        dblMax(intCounter) = mdlProcedures.GetMaximum(dblYTemp, UBound(dblYTemp, 1), UBound(dblYTemp, 2))
'
'        dblX = mdlProcedures.DivisionScalar2(dblYTemp, dblMax(intCounter))
'        dblY = dblMax(intCounter)
'
'        If intCounter > 0 Then
'            If Abs(dblMax(intCounter) - dblMax(intCounter - 1)) / dblMax(intCounter) <= dblEpsilon Then
'                blnLoop = False
'            End If
'        End If
'    Wend
'End Sub

Public Function EigenValues() As Double()
    EigenValues = dblEigenValues
End Function

Public Function EigenVectors() As Double()
    EigenVectors = dblEigenVectors
End Function
