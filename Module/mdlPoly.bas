Attribute VB_Name = "mdlPoly"
Option Explicit

'Thanks to Jim Carbine

Private Const prfMaxIterations As Integer = 2000
Private Const prfTolerance As Double = 0.0000000001
Private Const prfInitialGuess As Double = 0.005

Public Function Solve(ByRef A() As Double, ByVal intOrder As Integer, ByRef dblRoot() As Double) As Integer
    Dim Iteration As Integer
    Dim N As Integer
    Dim Nx As Integer
    Dim Ny As Integer
    Dim RootCounter As Integer
    Dim L As Integer
    Dim i As Integer
    Dim in1 As Integer
    Dim ic As Integer

    Dim C() As Double
    Dim X As Double
    Dim Xo As Double
    Dim X2 As Double
    Dim Xp As Double
    Dim Xt As Double

    Dim Y As Double
    Dim Yo As Double
    Dim Y2 As Double
    Dim Yp As Double
    Dim Yt As Double
    
    Dim dX As Double
    Dim uX As Double
    Dim dY As Double
    Dim uY As Double
    Dim u As Double
    Dim v As Double
    Dim Sq As Double
    Dim AL As Double
    Dim t As Double

    Dim Msg As String


    ReDim C(intOrder)
    Iteration = 0
    RootCounter = 0

    N = intOrder
    Nx = intOrder
    Ny = intOrder

'Reverse the order of the coefficients
    For L = 0 To intOrder
        i = intOrder - L
        C(i) = A(L)
    Next L

Rem: Set initial calculation values
lab15:  Xo = prfInitialGuess
        Yo = prfInitialGuess
        in1 = 0
        
lab8:   X = Xo
        Xo = -10 * Yo
        Yo = -10 * X
        X = Xo              'set X to current value
        Y = Yo              'set Y to current value
        in1 = in1 + 1
        GoTo lab3
        
lab10:  Iteration = 1
        Xp = X
        Yp = Y

Rem: Evaluate polynomial and derivatives
lab3:   ic = 0

lab7:   uX = 0
        uY = 0
        v = 0
        Yt = 0
        Xt = 1
        u = C(N)
        If u = 0 Then GoTo lab4

        For i = 1 To N
          L = N - i
          X2 = X * Xt - Y * Yt
          Y2 = X * Yt + Y * Xt
          u = u + C(L) * X2
          v = v + C(L) * Y2
          uX = uX + i * Xt * C(L)
          uY = uY - i * Yt * C(L)
          Xt = X2
          Yt = Y2
        Next i

        Sq = (uX ^ 2) + (uY ^ 2)
        If Sq = 0 Then GoTo lab5
        dX = (v * uY - u * uX) / Sq
        X = X + dX
        dY = -(u * uY + v * uX) / Sq
        Y = Y + dY
        If Abs(dY) + Abs(dX) - prfTolerance >= 0 Then
            ic = ic + 1
            If ic - prfMaxIterations < 0 Then GoTo lab7
            If Iteration = 0 Then
                If in1 - 5 < 0 Then
                    GoTo lab8
                Else
                    Solve = ic
                    Exit Function
                End If
            End If
        End If

'Rem: Set the step iteration counter
lab6:   For L = 0 To Ny
          i = intOrder - L
          t = A(i)
          A(i) = C(L)
          C(L) = t
        Next L
        i = N
        N = Nx
        Nx = i
        If Iteration <> 0 Then
            GoTo lab9
        Else
            GoTo lab10
        End If
        
lab5:   If Iteration = 0 Then GoTo lab8
        X = Xp
        Y = Yp

lab9:   Iteration = 0
        If Abs(Y) - (prfTolerance * Abs(X)) < 0 Then GoTo lab11
        AL = X + X
        Sq = (X ^ 2) + (Y ^ 2)
        N = N - 2
        GoTo lab12
        
lab4:   X = 0
        Nx = Nx - 1
        Ny = Ny - 1

lab11:  Y = 0
        Sq = 0
        AL = X
        N = N - 1
        
lab12:  C(1) = C(1) + AL * C(0)
        For L = 2 To N
          C(L) = C(L) + AL * C(L - 1) - Sq * C(L - 2)
        Next L


lab14:  dblRoot(RootCounter) = X
        RootCounter = RootCounter + 1
        
        If Sq = 0 Then
            If N > 0 Then
                GoTo lab15
            Else
                Solve = 0
                Exit Function
            End If
        Else
            Y = -Y
            Sq = 0
            GoTo lab14
        End If
End Function
