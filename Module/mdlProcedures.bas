Attribute VB_Name = "mdlProcedures"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function InputMatrix(ByRef flxInput As MSFlexGrid, ByVal intMatrixRow As Integer, ByVal intMatrixColumn As Integer) As Double()
    Dim dblInputMatrix() As Double
    
    ReDim dblInputMatrix(intMatrixRow, intMatrixColumn) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    With flxInput
        For intRow = 0 To intMatrixRow
            For intColumn = 0 To intMatrixColumn
                If Not IsNumeric(.TextMatrix(intRow, intColumn)) Then
                    dblInputMatrix(intRow, intColumn) = 0
                Else
                    dblInputMatrix(intRow, intColumn) = CDbl(.TextMatrix(intRow, intColumn))
                End If
            Next intColumn
        Next intRow
    End With
    
    InputMatrix = dblInputMatrix
End Function

Public Sub SetMatrixToTextBox(ByRef txtResult As TextBox, ByRef dblResult() As Double, ByVal intMatrixMode As Integer)
    Dim intRow As Integer
    Dim intColumn As Integer
    
    txtResult.Text = ""
    
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            txtResult.Text = txtResult.Text & CStr(Format(dblResult(intRow, intColumn), "#0.00")) & vbTab
        Next intColumn
        
        txtResult.Text = Left(txtResult.Text, Len(txtResult.Text) - 1)
        
        txtResult.Text = txtResult.Text & vbCrLf
    Next intRow
    
    If Len(txtResult.Text) > 0 Then txtResult.Text = Left(txtResult.Text, Len(txtResult.Text) - 2)
End Sub

Public Sub SetMatrixToFlexGrid(ByRef flxMain As MSFlexGrid, ByRef dblSource() As Double, ByVal intMatrixMode As Integer, Optional ByVal blnClear As Boolean = False)
    Dim intRow As Integer
    Dim intColumn As Integer
    
    With flxMain
        For intRow = 0 To intMatrixMode - 1
            For intColumn = 0 To intMatrixMode - 1
                If blnClear Then
                    .TextMatrix(intRow, intColumn) = ""
                Else
                    .TextMatrix(intRow, intColumn) = Format(dblSource(intRow, intColumn), "0.00")
                End If
            Next intColumn
        Next intRow
    End With
End Sub

Public Sub MessageBoxPolinomial(ByRef strPolinomial() As String, Optional ByVal strOperator As String = "+")
    Dim strPoli As String
    
    strPoli = ""
    
    Dim intCounter As Integer
    
    For intCounter = 0 To UBound(strPolinomial)
        If Not intCounter = 0 Then
            strPoli = strPoli & strOperator
        End If
    
        strPoli = strPoli & "(" & strPolinomial(intCounter) & ")"
    Next intCounter
    
    MsgBox strPoli, vbOKOnly + vbInformation, "Polinomial"
End Sub

Public Function BreakPolinomial(ByRef strPolinomial() As String, ByVal intMatrixMode As Integer) As String()
    Dim strBreak() As String
    
    ReDim strBreak(UBound(strPolinomial) - 1) As String
    
    Dim dblDeterminant As Double
    
    dblDeterminant = CDbl(strPolinomial(UBound(strPolinomial)))
    
    Dim dblValue() As Double
    
    ReDim dblValue(UBound(strPolinomial)) As Double
    
    Dim intCounter As Integer
    
    For intCounter = 0 To UBound(strPolinomial)
        dblValue(intCounter) = GetValueFromEachEquation(strPolinomial(UBound(strPolinomial) - intCounter))
    Next intCounter
    
    Dim dblResult() As Double
    
    ReDim dblResult(intMatrixMode - 1) As Double
    
    Dim intResult As Integer
    
    intResult = mdlPoly.Solve(dblValue, intMatrixMode, dblResult)
    
    Dim strResult() As String
    
    If intResult = 0 Then
        ReDim strResult(intMatrixMode - 1) As String
        
        For intCounter = 0 To intMatrixMode - 1
            strResult(intCounter) = CStr(dblResult(intCounter))
        Next intCounter
    Else
        ReDim strResult(0) As String
        
        strResult(intCounter) = "0"
    End If
    
    BreakPolinomial = strResult
End Function

Private Function GetValueFromEachEquation(ByVal strEquation As String) As Double
    Dim strTemp As String
    
    If InStr(strEquation, "*") > 0 Then
        strTemp = Mid(strEquation, InStr(strEquation, "*") + 1)
        strTemp = Left(strTemp, InStr(strTemp, "(") - 1)
        
        If Left(strEquation, 1) = "-" Then
            GetValueFromEachEquation = -CDbl(strTemp)
        Else
            GetValueFromEachEquation = CDbl(strTemp)
        End If
    ElseIf InStr(strEquation, Chr(182)) > 0 Then
        GetValueFromEachEquation = 1#
    Else
        GetValueFromEachEquation = CDbl(strEquation)
    End If
End Function

Public Function TraceMatrix(ByRef dblSource() As Double, ByVal intMatrixMode As Integer) As Double
    Dim dblTrace As Double
    
    dblTrace = 0#

    Dim intRow As Integer
    
    For intRow = 0 To intMatrixMode - 1
        dblTrace = dblTrace + dblSource(intRow, intRow)
    Next intRow
    
    TraceMatrix = dblTrace
End Function

Public Function CopyMatrix(ByRef dblSource() As Double, ByVal intMatrixMode As Integer) As Double()
    Dim dblDuplicate() As Double
    
    ReDim dblDuplicate(intMatrixMode - 1, intMatrixMode - 1) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            dblDuplicate(intRow, intColumn) = dblSource(intRow, intColumn)
        Next intColumn
    Next intRow
    
    CopyMatrix = dblDuplicate
End Function

Public Sub CopyMatrix2(ByRef dblSource() As Double, ByRef dblResult() As Double, ByVal intMatrixMode As Integer)
    If UBound(dblResult, 1) < intMatrixMode - 1 Then
        Exit Sub
    ElseIf UBound(dblResult, 2) < intMatrixMode - 1 Then
        Exit Sub
    End If
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            dblResult(intRow, intColumn) = dblSource(intRow, intColumn)
        Next intColumn
    Next intRow
End Sub

Public Function SetTranspose(ByRef dblSource() As Double, ByVal intMatrixMode As Integer) As Double()
    Dim dblTranspose() As Double
    
    ReDim dblTranspose(intMatrixMode - 1, intMatrixMode - 1) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            dblTranspose(intRow, intColumn) = dblSource(intColumn, intRow)
        Next intColumn
    Next intRow
    
    SetTranspose = dblTranspose
End Function

Public Function SetTranspose2(ByRef dblSource() As Double)
    Dim dblTranspose() As Double
    
    ReDim dblTranspose(UBound(dblSource, 2), UBound(dblSource, 1)) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To UBound(dblTranspose, 1)
        For intColumn = 0 To UBound(dblTranspose, 2)
            dblTranspose(intRow, intColumn) = dblSource(intColumn, intRow)
        Next intColumn
    Next intRow
    
    SetTranspose2 = dblTranspose
End Function

Public Function SetIdentity(ByVal intMatrixMode As Integer) As Double()
    Dim dblIdentity() As Double
    
    ReDim dblIdentity(intMatrixMode - 1, intMatrixMode - 1) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            If intColumn = intRow Then
                dblIdentity(intRow, intColumn) = 1#
            Else
                dblIdentity(intRow, intColumn) = 0#
            End If
        Next intColumn
    Next intRow
    
    SetIdentity = dblIdentity
End Function

Public Function SetInverse(ByRef dblSource() As Double, ByVal intMatrixMode As Integer) As Double()
    Dim dblAdjoint() As Double
    
    dblAdjoint = mdlProcedures.SetAdjoint(dblSource, intMatrixMode)
    
    Dim dblDeterminant As Double
    
    dblDeterminant = mdlProcedures.GetDeterminant(dblSource, intMatrixMode)
    
    Dim dblInverse() As Double
    
    If CCur(dblDeterminant) = 0 Then
        ReDim dblInverse(intMatrixMode - 1, intMatrixMode - 1) As Double
    Else
        dblInverse = mdlProcedures.MultiplyScalar(dblAdjoint, 1 / dblDeterminant, intMatrixMode)
    End If
    
    SetInverse = dblInverse
End Function

Public Function SetAdjoint(ByRef dblSource() As Double, ByVal intMatrixMode As Integer) As Double()
    Dim dblAdjoint() As Double
    Dim dblMinor() As Double
    
    ReDim dblAdjoint(intMatrixMode - 1, intMatrixMode - 1) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            dblMinor = mdlProcedures.SetMinor(dblSource, intMatrixMode, intRow, intColumn)
            
            dblAdjoint(intRow, intColumn) = ((-1) ^ ((intRow + 1) + (intColumn + 1))) * (mdlProcedures.GetDeterminant(dblMinor, intMatrixMode - 1))
        Next intColumn
    Next intRow
    
    dblAdjoint = mdlProcedures.SetTranspose(dblAdjoint, intMatrixMode)
    
    SetAdjoint = dblAdjoint
End Function

Public Function SetMinor(ByRef dblSource() As Double, ByVal intMatrixMode As Integer, ByVal intRowInput As Integer, ByVal intColumnInput As Integer) As Double()
    Dim dblMinor() As Double
    
    ReDim dblMinor(intMatrixMode - 2, intMatrixMode - 2) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    Dim intM As Integer
    Dim intN As Integer
    
    intM = 0
    intN = 0
    
    For intRow = 0 To intMatrixMode - 1
        If Not intRow = intRowInput Then
            intN = 0
            
            For intColumn = 0 To intMatrixMode - 1
                If Not intColumn = intColumnInput Then
                    dblMinor(intM, intN) = dblSource(intRow, intColumn)
                    
                    intN = intN + 1
                End If
            Next intColumn
            
            intM = intM + 1
        End If
    Next intRow
    
    SetMinor = dblMinor
End Function

Public Function GetDeterminant(ByRef dblSource() As Double, ByVal intMatrixMode As Integer) As Double
    Dim dblTemp() As Double
    
    Dim dblResult As Double
    
    Dim blnUpperTringular As Boolean
    
    blnUpperTringular = mdlProcedures.IsUpperTringular(dblSource, intMatrixMode)
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    If Not blnUpperTringular Then
        If intMatrixMode > 2 Then
            mdlPivot.CalculatePivot dblSource, dblTemp, intMatrixMode
        End If
    End If
    
    dblResult = 1
    
    For intRow = 0 To intMatrixMode - 1
        If blnUpperTringular Then
            dblResult = dblResult * dblSource(intRow, intRow)
        Else
            If intMatrixMode > 2 Then
                dblResult = dblResult * dblTemp(intRow, intRow)
            Else
                dblResult = (dblSource(0, 0) * dblSource(1, 1)) - (dblSource(0, 1) * dblSource(1, 0))
            End If
        End If
    Next intRow
    
    GetDeterminant = dblResult
End Function

Public Function ReducedEchelonMatrix(ByRef dblSource() As Double, ByRef dblResult() As Double, ByVal intMatrixMode As Integer) As Double()
    Dim dblTemp() As Double
    
    Dim dblTempValue As Double
    
    ReDim dblResult(intMatrixMode - 1) As Double
    
    dblTemp = mdlProcedures.CopyMatrix(dblSource, intMatrixMode)

    Dim intRow As Integer
    Dim intColumn As Integer
    Dim intTemp As Integer
    Dim intCounter As Integer
    Dim intCounter2 As Integer
    
    intTemp = 0
    
    For intColumn = 0 To intMatrixMode - 1
        For intRow = intTemp To intMatrixMode - 1
            If Not dblTemp(intRow, intColumn) = 0# Then
                If Not intRow = intTemp Then
                    For intCounter = intColumn To intMatrixMode - 1
                        dblTempValue = dblTemp(intRow, intCounter)
                        
                        dblTemp(intRow, intCounter) = dblTemp(intTemp, intCounter)
                        dblTemp(intTemp, intCounter) = dblTempValue
                    Next intCounter
                    
                    dblTempValue = dblResult(intRow)
                    
                    dblResult(intRow) = dblResult(intTemp)
                    dblResult(intTemp) = dblTempValue
                End If
                
                dblTempValue = 1 / dblTemp(intTemp, intColumn)
                
                dblTemp(intTemp, intColumn) = 1#
                
                For intCounter = intColumn + 1 To intMatrixMode - 1
                    dblTemp(intTemp, intCounter) = dblTemp(intTemp, intCounter) * dblTempValue
                Next intCounter
                
                dblResult(intTemp) = dblResult(intTemp) * dblTempValue
                
                For intCounter2 = 0 To intMatrixMode - 1
                    If intCounter2 <> intTemp Then
                        dblTempValue = dblTemp(intCounter2, intColumn)
                        
                        dblTemp(intCounter2, intColumn) = 0#
                        
                        For intCounter = intColumn + 1 To intMatrixMode - 1
                            dblTemp(intCounter2, intCounter) = dblTemp(intCounter2, intCounter) - dblTempValue * dblTemp(intTemp, intCounter)
                        Next intCounter
                        
                        dblResult(intCounter2) = dblResult(intCounter2) - dblTempValue * dblResult(intTemp)
                    End If
                Next intCounter2
                
                intTemp = intTemp + 1
                
                Exit For
            End If
        Next intRow
    Next intColumn
    
    ReducedEchelonMatrix = dblTemp
End Function

Public Function MultiplyScalar(ByRef dblSource() As Double, ByVal dblMultiple As Double, ByVal intMatrixMode As Integer) As Double()
    Dim dblResult() As Double
    
    ReDim dblResult(intMatrixMode - 1, intMatrixMode - 1) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            dblResult(intRow, intColumn) = dblSource(intRow, intColumn) * dblMultiple
        Next intColumn
    Next intRow
    
    MultiplyScalar = dblResult
End Function

Public Function MultiplyScalar2(ByRef dblSource() As Double, ByVal dblMultiple As Double, ByVal intRowTemp As Integer, ByVal intColumnTemp As Integer) As Double()
    Dim dblResult() As Double
    
    ReDim dblResult(intRowTemp, intColumnTemp) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To intRowTemp
        For intColumn = 0 To intColumnTemp
            dblResult(intRow, intColumn) = dblSource(intRow, intColumn) * dblMultiple
        Next intColumn
    Next intRow
    
    MultiplyScalar2 = dblResult
End Function

Public Function MultiplyRow(ByRef dblSource() As Double, ByVal intRow As Integer, ByVal dblMultiple As Double, ByVal intMatrixMode As Integer)
    Dim intCounter As Integer
    
    For intCounter = 0 To intMatrixMode - 1
        dblSource(intRow, intCounter) = dblSource(intRow, intCounter) * dblMultiple
    Next intCounter
End Function

Public Function SubtractMatrix(ByRef dblSource1() As Double, ByRef dblSource2() As Double) As Double()
    Dim dblResult() As Double
    
    ReDim dblResult(UBound(dblSource1, 1), UBound(dblSource1, 2)) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To UBound(dblSource1, 1)
        For intColumn = 0 To UBound(dblSource1, 2)
            dblResult(intRow, intColumn) = dblSource1(intRow, intColumn) - dblSource2(intRow, intColumn)
        Next intColumn
    Next intRow
    
    SubtractMatrix = dblResult
End Function

Public Function DivisionScalar2(ByRef dblSource() As Double, ByVal dblDivision As Double) As Double()
    Dim dblResult() As Double
    
    ReDim dblResult(UBound(dblSource, 1), UBound(dblSource, 2)) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To UBound(dblSource, 1)
        For intColumn = 0 To UBound(dblSource, 2)
            If Not dblSource(intRow, intColumn) = 0# Then
                dblResult(intRow, intColumn) = dblSource(intRow, intColumn) / dblDivision
            End If
        Next intColumn
    Next intRow
    
    DivisionScalar2 = dblResult
End Function

Public Function MultiplyMatrix( _
    ByRef dblSource1() As Double, _
    ByRef dblSource2() As Double, _
    ByVal intMatrixMode As Integer) As Double()
    Dim dblResult() As Double
    
    ReDim dblResult(intMatrixMode - 1, intMatrixMode - 1) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    Dim intLoop As Integer
    
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            For intLoop = 0 To intMatrixMode - 1
                dblResult(intRow, intColumn) = _
                    dblResult(intRow, intColumn) + _
                    (dblSource1(intRow, intLoop) * dblSource2(intLoop, intColumn))
            Next intLoop
        Next intColumn
    Next intRow
    
    MultiplyMatrix = dblResult
End Function

Public Function MultiplyMatrix2( _
    ByRef dblSource1() As Double, _
    ByRef dblSource2() As Double) As Double()
    Dim intRow As Integer
    Dim intColumn As Integer
    
    intRow = UBound(dblSource1, 1)
    intColumn = UBound(dblSource2, 2)
    
    Dim dblResult() As Double
    
    ReDim dblResult(intRow, intColumn) As Double
    
    intRow = UBound(dblSource2, 1)
    intColumn = UBound(dblSource1, 2)
    
    If Not intColumn = intRow Then
        MultiplyMatrix2 = dblResult
        
        Exit Function
    End If
    
    intRow = UBound(dblSource1, 1)
    intColumn = UBound(dblSource2, 2)
    
    Dim intI As Integer
    Dim intJ As Integer
    Dim intK As Integer
    
    For intI = 0 To intRow
        For intJ = 0 To intColumn
            For intK = 0 To UBound(dblSource2, 1)
                dblResult(intI, intJ) = _
                    dblResult(intI, intJ) + _
                    (dblSource1(intI, intK) * dblSource2(intK, intJ))
            Next intK
        Next intJ
    Next intI
    
    MultiplyMatrix2 = dblResult
End Function

Public Function MultiplyMatrix3( _
    ByRef dblSource1() As Double, _
    ByRef dblSource2() As Double, _
    ByVal intPivot As Integer) As Double()
    Dim intRow As Integer
    Dim intColumn As Integer
    Dim intLoop As Integer
    
    intRow = UBound(dblSource2, 2)
    intColumn = 0
    
    Dim dblResult() As Double
    
    ReDim dblResult(intRow, intColumn) As Double
    
    For intRow = 0 To UBound(dblSource2, 1)
        For intColumn = 0 To 0
            For intLoop = 0 To UBound(dblSource2, 1)
                dblResult(intRow, intColumn) = _
                    dblResult(intRow, intColumn) + _
                    (dblSource1(intRow, intLoop) * dblSource2(intLoop, intPivot))
            Next intLoop
        Next intColumn
    Next intRow
    
    MultiplyMatrix3 = dblResult
End Function

Public Function IsSquare(ByRef dblSource() As Double, ByVal intMatrixMode As Integer) As Boolean
'    Dim dblMultiply() As Double
'
'    dblMultiply = mdlProcedures.MultiplyMatrix(dblSource, dblSource, intMatrixMode)
'
'    Dim dblMultiplyDeterminant As Double
'    Dim dblDeterminant As Double
'
'    dblMultiplyDeterminant = mdlProcedures.GetDeterminant(dblMultiply, intMatrixMode)
'    dblDeterminant = mdlProcedures.GetDeterminant(dblSource, intMatrixMode)
'
'    If CCur(dblMultiplyDeterminant) = CCur(dblDeterminant * dblDeterminant) Then
'        IsSquare = True
'    Else
'        IsSquare = False
'    End If
    IsSquare = True
End Function

Public Function IsNonSingular(ByRef dblSource() As Double, ByVal intMatrixMode As Integer) As Boolean
'    Dim dblInverse() As Double
'
'    dblInverse = mdlProcedures.SetInverse(dblSource, intMatrixMode)
'
'    Dim dblInverseDeterminant As Double
'    Dim dblDeterminant As Double
'
'    dblInverseDeterminant = (mdlProcedures.GetDeterminant(dblInverse, intMatrixMode))
'    dblDeterminant = mdlProcedures.GetDeterminant(dblSource, intMatrixMode)
'
'    Dim dblResult As Double
'
'    dblResult = dblInverseDeterminant * dblDeterminant
'
'    If CCur(dblResult) = CCur(1) Then
'        IsNonSingular = True
'    Else
'        IsNonSingular = False
'    End If

    Dim dblDeterminant As Double
    
    dblDeterminant = mdlProcedures.GetDeterminant(dblSource, intMatrixMode)
    
    If CCur(dblDeterminant) = 0 Then
        IsNonSingular = False
    Else
        IsNonSingular = True
    End If
End Function

Public Function IsSymmetric(ByRef dblSource() As Double, ByVal intMatrixMode As Integer) As Boolean
    Dim dblTranspose() As Double
    
    dblTranspose = mdlProcedures.SetTranspose(dblSource, intMatrixMode)
    
    Dim blnSymmetric As Boolean
    
    blnSymmetric = True
    
    Dim intRow As Integer
    
    For intRow = 0 To intMatrixMode - 1
        If Not CCur(dblTranspose(intRow, intRow)) = CCur(dblSource(intRow, intRow)) Then
            blnSymmetric = False
            
            Exit For
        End If
    Next intRow
    
    IsSymmetric = blnSymmetric
End Function

Public Function IsUpperTringular(ByRef dblSource() As Double, ByVal intMatrixMode As Integer) As Boolean
    Dim intRow As Integer
    Dim intColumn As Integer
    
    Dim blnTringular As Boolean
    
    blnTringular = True
    
    For intRow = 1 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            If intColumn = intRow Then
                Exit For
            Else
                If Not dblSource(intRow, intColumn) = 0 Then
                    blnTringular = False
                    
                    Exit For
                End If
            End If
        Next intColumn
        
        If Not blnTringular Then
            Exit For
        End If
    Next intRow
    
    IsUpperTringular = blnTringular
End Function

Public Function GetMaximum(ByRef dblSource() As Double, ByVal intRowTemp As Integer, ByVal intColumnTemp As Integer, Optional blnAbsolute As Boolean = True) As Double
    Dim dblMax As Double
    Dim dblTemp As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To intRowTemp
        For intColumn = 0 To intColumnTemp
            If blnAbsolute Then
                dblTemp = Abs(dblSource(intRow, intColumn))
            Else
                dblTemp = dblSource(intRow, intColumn)
            End If
            
            If intRow = 0 And intColumn = 0 Then
                dblMax = dblTemp
            ElseIf dblMax < dblTemp Then
                dblMax = dblTemp
            End If
        Next intColumn
    Next intRow
    
    GetMaximum = dblMax
End Function

Public Function IsValidRegion() As Boolean
    Dim dteValue As Date
    
    dteValue = "01/31/1990"
    
    Dim strValue As String
    
    strValue = CStr(dteValue)
    
    If Left(strValue, 2) = "31" Then
        IsValidRegion = False
    Else
        IsValidRegion = True
    End If
End Function
