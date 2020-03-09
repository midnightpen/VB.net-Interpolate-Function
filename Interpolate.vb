    Public Function Interpolate(ByVal FindX As Double, ByVal X1 As Double, ByVal X2 As Double, ByVal Y1 As Double, ByVal Y2 As Double, Optional boolExtrapolate As Boolean = True) As Double
        Dim Interpolated As Double
        Dim UpperLimitX As Double
        Dim UpperLimitY As Double
        Dim LowerLimitX As Double
        Dim LowerLimitY As Double
        Try
            If Not boolExtrapolate Then
                If X2 > X1 Then
                    LowerLimitX = X1
                    UpperLimitX = X2
                    LowerLimitY = Y1
                    UpperLimitY = Y2
                Else
                    LowerLimitX = X2
                    UpperLimitX = X1
                    LowerLimitY = Y2
                    UpperLimitY = Y1
                End If

                If FindX <= LowerLimitX Then
                    'If FindX is less than or equal to X1, return Y will allways be Y1
                    Interpolated = LowerLimitY
                ElseIf FindX >= UpperLimitX Then
                    'If FindX is greater than or equal to X2, Return Y will allways be Y2
                    Interpolated = UpperLimitY
                End If
            Else
                Interpolated = (Y1 + (Y2 - Y1) * (FindX - X1) / (X2 - X1))
            End If
        Catch ex As Exception
            Dim Error_Location As String
            Error_Location = "Interpolate"
            MessageBox.Show(vbExclamation, Error_Location & ":" & Err.Number)
            Interpolated = 0
        End Try

        Return Interpolated
    End Function
