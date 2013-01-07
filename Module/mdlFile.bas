Attribute VB_Name = "mdlFile"
Option Explicit

Private Const strMatrix As String = "MATRIX"

Private Const strLText As String = "MATRIX L"
Private Const strUText As String = "MATRIX U"

Private Const strEIGENVALUEText As String = "EIGENVALUES"
Private Const strEIGENVECTORText As String = "EIGENVECTORS"

Private Const strSVDAText As String = "MATRIX A"
Private Const strSVDEText As String = "MATRIX E"
Private Const strSVDVTText As String = "MATRIX VT"

Public Sub WriteTextLU( _
    ByVal strFileName As String, _
    ByVal strL As String, _
    ByVal strU As String, _
    ByRef dblSource() As Double, _
    ByVal intMatrixMode As Integer)
    Dim txsFile As TextStream
    
    Set txsFile = mdlGlobal.fso.CreateTextFile(strFileName, False)
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    With txsFile
        .WriteLine strMatrix
        
        Dim strValue As String
        
        For intRow = 0 To intMatrixMode - 1
            strValue = ""
            
            For intColumn = 0 To intMatrixMode - 1
                If Not intColumn = 0 Then
                    strValue = strValue & vbTab
                End If
                
                strValue = strValue & Format(dblSource(intRow, intColumn), "0.00")
            Next intColumn
            
            .WriteLine strValue
        Next intRow
        
        .WriteLine
        .WriteLine strLText
        .WriteLine strL
        
        .WriteLine
        .WriteLine strUText
        .WriteLine strU
        .WriteLine
        
        .Close
    End With
    
    Set txsFile = Nothing
End Sub

Public Function ReadTextLU( _
    ByVal strFileName As String, _
    ByRef strL As String, _
    ByRef strU As String, _
    ByRef dblSource() As Double) As Integer
    Dim intMatrixMode As Integer
    
    intMatrixMode = 0
    
    Dim txsFile As TextStream
    
    Set txsFile = mdlGlobal.fso.OpenTextFile(strFileName)
    
    With txsFile
        Dim dblMatrix() As Double
        
        Dim intCounter As Integer
        Dim intRow As Integer
        
        Dim strValue As String
        
        Dim strSplit() As String
        
        While Not .AtEndOfStream
            Select Case .ReadLine
                Case strMatrix:
                    intRow = 0
                    
                    Do
                        If .AtEndOfStream Then Exit Do
                        
                        intRow = intRow + 1
                        strValue = .ReadLine
                        
                        If Not Trim(strValue) = "" Then
                            strSplit = Split(strValue, vbTab)
                            
                            If intMatrixMode = 0 Then
                                intMatrixMode = UBound(strSplit) + 1
                                
                                ReDim dblMatrix(intMatrixMode - 1, intMatrixMode - 1) As Double
                            End If
                            
                            For intCounter = 0 To UBound(strSplit)
                                If mdlProcedures.IsValidRegion Then
                                    strSplit(intCounter) = Replace(strSplit(intCounter), ",", ".")
                                Else
                                    strSplit(intCounter) = Replace(strSplit(intCounter), ".", ",")
                                End If
                                
                                dblMatrix(intRow - 1, intCounter) = CDbl(strSplit(intCounter))
                            Next intCounter
                        End If
                    Loop Until Trim(strValue) = ""
                Case strLText:
                    strL = ""
                    
                    Do
                        If .AtEndOfStream Then Exit Do
                        
                        strValue = .ReadLine
                        
                        If Not Trim(strValue) = "" Then
                            If Not Trim(strL) = "" Then
                                strL = strL & vbCrLf
                            End If
                            
                            strL = strL & strValue
                        End If
                    Loop Until Trim(strValue) = ""
                Case strUText:
                    strU = ""
                    
                    Do
                        If .AtEndOfStream Then Exit Do
                        
                        strValue = .ReadLine
                        
                        If Not Trim(strValue) = "" Then
                            If Not Trim(strU) = "" Then
                                strU = strU & vbCrLf
                            End If
                            
                            strU = strU & strValue
                        End If
                    Loop Until Trim(strValue) = ""
            End Select
        Wend
        
        .Close
    End With
    
    Set txsFile = Nothing
    
    dblSource = dblMatrix
    ReadTextLU = intMatrixMode
End Function

Public Sub WriteTextEigen( _
    ByVal strFileName As String, _
    ByVal strEigenValues As String, _
    ByVal strEigenVectors As String, _
    ByRef dblSource() As Double, _
    ByVal intMatrixMode As Integer)
    Dim txsFile As TextStream
    
    Set txsFile = mdlGlobal.fso.CreateTextFile(strFileName, False)
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    With txsFile
        .WriteLine strMatrix
        
        Dim strValue As String
        
        For intRow = 0 To intMatrixMode - 1
            strValue = ""
            
            For intColumn = 0 To intMatrixMode - 1
                If Not intColumn = 0 Then
                    strValue = strValue & vbTab
                End If
                
                strValue = strValue & Format(dblSource(intRow, intColumn), "0.00")
            Next intColumn
            
            .WriteLine strValue
        Next intRow
        
        .WriteLine
        .WriteLine strEIGENVALUEText
        .WriteLine strEigenValues
        
        .WriteLine
        .WriteLine strEIGENVECTORText
        .WriteLine strEigenVectors
        .WriteLine
        
        .Close
    End With
    
    Set txsFile = Nothing
End Sub

Public Function ReadTextEigen( _
    ByVal strFileName As String, _
    ByRef strEigenValues As String, _
    ByRef strEigenVectors As String, _
    ByRef dblSource() As Double) As Integer
    Dim intMatrixMode As Integer
    
    intMatrixMode = 0
    
    Dim txsFile As TextStream
    
    Set txsFile = mdlGlobal.fso.OpenTextFile(strFileName)
    
    With txsFile
        Dim dblMatrix() As Double
        
        Dim intCounter As Integer
        Dim intRow As Integer
        
        Dim strValue As String
        
        Dim strSplit() As String
        
        While Not .AtEndOfStream
            Select Case .ReadLine
                Case strMatrix:
                    intRow = 0
                    
                    Do
                        If .AtEndOfStream Then Exit Do
                        
                        intRow = intRow + 1
                        strValue = .ReadLine
                        
                        If Not Trim(strValue) = "" Then
                            strSplit = Split(strValue, vbTab)
                            
                            If intMatrixMode = 0 Then
                                intMatrixMode = UBound(strSplit) + 1
                                
                                ReDim dblMatrix(intMatrixMode - 1, intMatrixMode - 1) As Double
                            End If
                            
                            For intCounter = 0 To UBound(strSplit)
                                If mdlProcedures.IsValidRegion Then
                                    strSplit(intCounter) = Replace(strSplit(intCounter), ",", ".")
                                Else
                                    strSplit(intCounter) = Replace(strSplit(intCounter), ".", ",")
                                End If
                                
                                dblMatrix(intRow - 1, intCounter) = CDbl(strSplit(intCounter))
                            Next intCounter
                        End If
                    Loop Until Trim(strValue) = ""
                Case strEIGENVALUEText:
                    strEigenValues = ""
                    
                    Do
                        If .AtEndOfStream Then Exit Do
                        
                        strValue = .ReadLine
                        
                        If Not Trim(strValue) = "" Then
                            If Not Trim(strEigenValues) = "" Then
                                strEigenValues = strEigenValues & vbCrLf
                            End If
                            
                            strEigenValues = strEigenValues & strValue
                        End If
                    Loop Until Trim(strValue) = ""
                Case strEIGENVECTORText:
                    strEigenVectors = ""
                    
                    Do
                        If .AtEndOfStream Then Exit Do
                        
                        strValue = .ReadLine
                        
                        If Not Trim(strValue) = "" Then
                            If Not Trim(strEigenVectors) = "" Then
                                strEigenVectors = strEigenVectors & vbCrLf
                            End If
                            
                            strEigenVectors = strEigenVectors & strValue
                        End If
                    Loop Until Trim(strValue) = ""
            End Select
        Wend
        
        .Close
    End With
    
    Set txsFile = Nothing
    
    dblSource = dblMatrix
    ReadTextEigen = intMatrixMode
End Function

Public Sub WriteTextSVD( _
    ByVal strFileName As String, _
    ByVal strA As String, _
    ByVal strE As String, _
    ByVal strVT As String, _
    ByRef dblSource() As Double, _
    ByVal intMatrixMode As Integer)
    Dim txsFile As TextStream
    
    Set txsFile = mdlGlobal.fso.CreateTextFile(strFileName, False)
    
    Dim intRow As Integer
    Dim intColumn As Integer
    
    With txsFile
        .WriteLine strMatrix
        
        Dim strValue As String
        
        For intRow = 0 To intMatrixMode - 1
            strValue = ""
            
            For intColumn = 0 To intMatrixMode - 1
                If Not intColumn = 0 Then
                    strValue = strValue & vbTab
                End If
                
                strValue = strValue & Format(dblSource(intRow, intColumn), "0.00")
            Next intColumn
            
            .WriteLine strValue
        Next intRow
        
        .WriteLine
        .WriteLine strSVDAText
        .WriteLine strA
        
        .WriteLine
        .WriteLine strSVDEText
        .WriteLine strE
        
        .WriteLine
        .WriteLine strSVDVTText
        .WriteLine strVT
        .WriteLine
        
        .Close
    End With
    
    Set txsFile = Nothing
End Sub

Public Function ReadTextSVD( _
    ByVal strFileName As String, _
    ByRef strA As String, _
    ByRef strE As String, _
    ByRef strVT As String, _
    ByRef dblSource() As Double) As Integer
    Dim intMatrixMode As Integer
    
    intMatrixMode = 0
    
    Dim txsFile As TextStream
    
    Set txsFile = mdlGlobal.fso.OpenTextFile(strFileName)
    
    With txsFile
        Dim dblMatrix() As Double
        
        Dim intCounter As Integer
        Dim intRow As Integer
        
        Dim strValue As String
        
        Dim strSplit() As String
        
        While Not .AtEndOfStream
            Select Case .ReadLine
                Case strMatrix:
                    intRow = 0
                    
                    Do
                        If .AtEndOfStream Then Exit Do
                        
                        intRow = intRow + 1
                        strValue = .ReadLine
                        
                        If Not Trim(strValue) = "" Then
                            strSplit = Split(strValue, vbTab)
                            
                            If intMatrixMode = 0 Then
                                intMatrixMode = UBound(strSplit) + 1
                                
                                ReDim dblMatrix(intMatrixMode - 1, intMatrixMode - 1) As Double
                            End If
                            
                            For intCounter = 0 To UBound(strSplit)
                                If mdlProcedures.IsValidRegion Then
                                    strSplit(intCounter) = Replace(strSplit(intCounter), ",", ".")
                                Else
                                    strSplit(intCounter) = Replace(strSplit(intCounter), ".", ",")
                                End If
                                
                                dblMatrix(intRow - 1, intCounter) = CDbl(strSplit(intCounter))
                            Next intCounter
                        End If
                    Loop Until Trim(strValue) = ""
                Case strSVDAText:
                    strA = ""
                    
                    Do
                        If .AtEndOfStream Then Exit Do
                        
                        strValue = .ReadLine
                        
                        If Not Trim(strValue) = "" Then
                            If Not Trim(strA) = "" Then
                                strA = strA & vbCrLf
                            End If
                            
                            strA = strA & strValue
                        End If
                    Loop Until Trim(strValue) = ""
                Case strSVDEText:
                    strE = ""
                    
                    Do
                        If .AtEndOfStream Then Exit Do
                        
                        strValue = .ReadLine
                        
                        If Not Trim(strValue) = "" Then
                            If Not Trim(strE) = "" Then
                                strE = strE & vbCrLf
                            End If
                            
                            strE = strE & strValue
                        End If
                    Loop Until Trim(strValue) = ""
                Case strSVDVTText:
                    strVT = ""
                    
                    Do
                        If .AtEndOfStream Then Exit Do
                        
                        strValue = .ReadLine
                        
                        If Not Trim(strValue) = "" Then
                            If Not Trim(strVT) = "" Then
                                strVT = strVT & vbCrLf
                            End If
                            
                            strVT = strVT & strValue
                        End If
                    Loop Until Trim(strValue) = ""
            End Select
        Wend
        
        .Close
    End With
    
    Set txsFile = Nothing
    
    dblSource = dblMatrix
    ReadTextSVD = intMatrixMode
End Function
