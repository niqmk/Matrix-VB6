Attribute VB_Name = "mdlSVD"
Option Explicit

Private dblVT() As Double
Private dblA() As Double
Private dblE() As Double

Public Sub CalculateSVD(ByRef dblSource() As Double, ByVal intMatrixMode As Integer)
    Dim dblTranspose() As Double
    
    dblTranspose = mdlProcedures.SetTranspose(dblSource, intMatrixMode)
    
    Dim dblMultiply() As Double
    
    dblMultiply = mdlProcedures.MultiplyMatrix(dblTranspose, dblSource, intMatrixMode)
    
    mdlEigen.CalculateEigen dblMultiply, intMatrixMode
    
    Dim dblEigenValues() As Double
    Dim dblEigenVectors() As Double
    
    dblEigenValues = mdlEigen.EigenValues
    dblEigenVectors = mdlEigen.EigenVectors
    
'    dblEigenValues = mdlEigen.CalculateEigen(dblMultiply, intMatrixMode, intMatrixMode)
'
'    ReDim dblEigenVectors(mdlEigen.intIteration - 1, intMatrixMode - 1) As Double
'
'    Dim intRow As Integer
'    Dim intColumn As Integer
'
'    For intRow = 0 To mdlEigen.intIteration - 1
'        For intColumn = 0 To intMatrixMode - 1
'            dblEigenVectors(intRow, intColumn) = mdlEigen.dblY(intRow, intColumn) / dblEigenValues(intRow)
'        Next intColumn
'    Next intRow
'
'    '------
'
    Dim dblAV() As Double

'    ReDim dblAV(intMatrixMode - 1, mdlEigen.intIteration - 1) As Double

    ReDim dblAV(intMatrixMode - 1, intMatrixMode - 1) As Double
    
    Dim dblTemp() As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer

'    For intRow = 0 To mdlEigen.intIteration - 1
    For intRow = 0 To intMatrixMode - 1
        dblTemp = mdlProcedures.MultiplyMatrix3(dblSource, dblEigenVectors, intRow)

        For intColumn = 0 To intMatrixMode - 1
            dblAV(intColumn, intRow) = dblTemp(intColumn, 0)
        Next intColumn
    Next intRow

    Dim dblO() As Double

'    ReDim dblO(mdlEigen.intIteration - 1) As Double
    ReDim dblO(intMatrixMode - 1) As Double

'    For intRow = 0 To mdlEigen.intIteration - 1
    For intRow = 0 To intMatrixMode - 1
        dblO(intRow) = Math.Sqr(dblEigenValues(intRow))
    Next intRow

'    ReDim dblA(mdlEigen.intIteration - 1, intMatrixMode - 1) As Double
    ReDim dblA(intMatrixMode - 1, intMatrixMode - 1) As Double

'    For intRow = 0 To mdlEigen.intIteration - 1
    For intRow = 0 To intMatrixMode - 1
        dblTemp = mdlProcedures.MultiplyScalar2(dblAV, dblO(intRow), UBound(dblAV, 1), 0)

        For intColumn = 0 To intMatrixMode - 1
            dblA(intRow, intColumn) = dblTemp(intColumn, 0)
        Next intColumn
    Next intRow

    dblE = mdlProcedures.SetIdentity(intMatrixMode)

'    For intRow = 0 To mdlEigen.intIteration - 1
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            If dblE(intRow, intColumn) = 1# Then
                dblE(intRow, intColumn) = dblO(intRow)
            End If
        Next intColumn
    Next intRow

'    ReDim dblVT(mdlEigen.intIteration - 1, intMatrixMode - 1) As Double
    ReDim dblVT(intMatrixMode - 1, intMatrixMode - 1) As Double

'    For intRow = 0 To mdlEigen.intIteration - 1
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            dblVT(intRow, intColumn) = dblEigenVectors(intRow, intColumn)
        Next intColumn
    Next intRow
End Sub

Public Function MatrixA() As Double()
    MatrixA = dblA
End Function

Public Function MatrixVT() As Double()
    MatrixVT = dblVT
End Function

Public Function MatrixE() As Double()
    MatrixE = dblE
End Function
