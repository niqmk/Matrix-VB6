Attribute VB_Name = "mdlExcel"
Option Explicit

Private xlApp As Excel.Application
Private xlBook As Excel.Workbook

Public intMatrixChange As Integer

Public Function GetMatrix(ByVal strFileName As String, Optional ByVal intMatrixMode As Integer = 0) As Double()
    On Local Error GoTo ErrHandler
    
    OpenApplication strFileName, False
    
    Dim xlSheet As Excel.Worksheet
    
    Set xlSheet = xlBook.Worksheets(1)
    
    If intMatrixMode = 0 Then
        intMatrixMode = TestMatrix(xlSheet)
        
        mdlExcel.intMatrixChange = intMatrixMode
    End If
    
    Dim dblMatrix() As Double
    
    ReDim dblMatrix(intMatrixMode - 1, intMatrixMode - 1) As Double
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 1 To intMatrixMode
        For intColumn = 1 To intMatrixMode
            dblMatrix(intRow - 1, intColumn - 1) = ReadMatrix(xlSheet, intRow, intColumn)
        Next intColumn
    Next intRow
    
    Set xlSheet = Nothing
    
    CloseApplication
    
    GetMatrix = dblMatrix
    
ErrHandler:
End Function

Private Function TestMatrix(ByRef xlSheet As Excel.Worksheet) As Integer
    Dim intRow As Integer
    Dim intColumn As Integer
    
    Dim intStop As Integer
    
    intStop = 0
    
    With xlSheet
        For intColumn = 1 To mdlGlobal.intMatrix
            If Not IsNumeric(Trim(.Cells(1, intColumn).Value)) Or Trim(.Cells(1, intColumn).Value) = "" Then
                intStop = intColumn
                
                Exit For
            End If
        Next intColumn
        
        For intRow = 1 To mdlGlobal.intMatrix
            If Not IsNumeric(Trim(.Cells(intRow, 1).Value)) Or Trim(.Cells(intRow, 1).Value) = "" Then
                If Not intStop >= intRow Then
                    intStop = intRow
                End If
                
                Exit For
            End If
        Next intRow
    End With
    
    If intStop < 2 Then
        TestMatrix = 2
    Else
        TestMatrix = intStop
    End If
End Function

Public Sub SaveMatrix( _
    ByRef dblSource() As Double, _
    ByVal strFileName As String, _
    ByVal intMatrixMode As Integer)
    On Local Error GoTo ErrHandler
    
    OpenApplication
    
    Dim xlSheet As Excel.Worksheet
    
    Set xlSheet = xlBook.Worksheets.Add
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    For intRow = 0 To intMatrixMode - 1
        For intColumn = 0 To intMatrixMode - 1
            xlSheet.Cells(intRow + 1, intColumn + 1).Value = dblSource(intRow, intColumn)
        Next intColumn
    Next intRow
    
    Set xlSheet = Nothing
    
    CloseApplication strFileName
    
ErrHandler:
End Sub

Public Sub PrintMatrix( _
    ByRef dblSource() As Double, _
    ByRef dblInverse() As Double, _
    ByVal intMatrixMode As Integer)
    On Local Error GoTo ErrHandler
    
    OpenApplication , False
    
    Dim xlSheet As Excel.Worksheet
    
    Set xlSheet = xlBook.Worksheets.Add
    
    Dim intTemp As Integer
    Dim intCounter As Integer
    Dim intRow As Integer
    Dim intColumn As Integer
    
    intCounter = 1
    
    xlSheet.Cells(intCounter, 1).Value = "Matrix"
    
    intTemp = intCounter + 1
    
    For intRow = 0 To intMatrixMode - 1
        intCounter = intCounter + 1
        
        For intColumn = 0 To intMatrixMode - 1
            xlSheet.Cells(intCounter, intColumn + 1).Value = dblSource(intRow, intColumn)
        Next intColumn
    Next intRow
    
    Dim strAlphabet As String
    
    strAlphabet = GetAlphabetExcel(intMatrixMode)
    
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlEdgeRight).LineStyle = xlContinuous
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    intCounter = intCounter + 2
    
    xlSheet.Cells(intCounter, 1).Value = "Matrix Inverse"
    
    intTemp = intCounter + 1
    
    For intRow = 0 To intMatrixMode - 1
        intCounter = intCounter + 1
        
        For intColumn = 0 To intMatrixMode - 1
            xlSheet.Cells(intCounter, intColumn + 1).Value = dblInverse(intRow, intColumn)
        Next intColumn
    Next intRow
    
    strAlphabet = GetAlphabetExcel(intMatrixMode)
    
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlEdgeRight).LineStyle = xlContinuous
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlSheet.Range("A" & intTemp & ":" & strAlphabet & intCounter).Borders(xlEdgeBottom).LineStyle = xlContinuous

    xlSheet.PrintOut
    
    Set xlSheet = Nothing
    
    CloseApplication mdlGlobal.strPath & "Print.xls"
    
ErrHandler:
End Sub

Private Function GetAlphabetExcel(ByVal intNumber As Integer) As String
    Dim intMod As Integer
    Dim intDivision As Integer
    
    intMod = intNumber Mod 26
    intDivision = intNumber \ 26
    
    GetAlphabetExcel = GetAlphabet(intMod) & GetAlphabet(intDivision, True)
End Function

Private Function GetAlphabet(ByVal intNumber As Integer, Optional ByVal blnAlphabetNull As Boolean = False) As String
    If blnAlphabetNull Then
        If intNumber = 0 Then
            GetAlphabet = ""
            
            Exit Function
        End If
    End If
    
    GetAlphabet = Mid("ZABCDEFGHIJKLMNOPQRSTUVWXYZ", intNumber + 1, 1)
End Function

Private Sub OpenApplication(Optional ByVal strFileName As String = "", Optional ByVal blnVisible As Boolean = True)
    Set xlApp = New Excel.Application
    
    If Not Trim(strFileName) = "" Then
        Set xlBook = xlApp.Workbooks.Open(strFileName)
    Else
        Set xlBook = xlApp.Workbooks.Add
    End If
    
    xlApp.Visible = blnVisible
End Sub

Private Function ReadMatrix(ByRef xlSheet As Excel.Worksheet, ByVal intRow As Integer, ByVal intColumn As Integer) As Double
    With xlSheet
        If Not IsNumeric(.Cells(intRow, intColumn).Value) Then
            ReadMatrix = 0
        Else
            ReadMatrix = CDbl(.Cells(intRow, intColumn).Value)
        End If
    End With
End Function

Private Sub CloseApplication(Optional ByVal strFileName As String = "")
    On Local Error GoTo ErrHandler
    
    If Trim(strFileName) = "" Then
        xlBook.Close
    Else
        xlBook.Close True, strFileName
    End If
    
    Set xlBook = Nothing
    
    Set xlApp = Nothing

ErrHandler:
End Sub
