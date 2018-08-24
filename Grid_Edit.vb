Imports jakStd20_MathExt

Partial Public Class Grid

    Public Sub SetAllElevationTo(value As Double)
        For c = 0 To ColCount - 1
            For r = 0 To RowCount - 1
                GridDtm(r, c) = value
            Next
        Next
        CountValidPoints()
    End Sub

    Public Function ChangeGridSize(rows As Integer, columns As Integer) As Boolean
        Dim res As Boolean
        If rows > 0 Then
            If columns > 0 Then
                ReDim Preserve GridDtm(rows, columns)
                res = True
            Else
                ReDim Preserve GridDtm(rows, ColCount)
                res = True
            End If
        Else
            If columns > 0 Then
                ReDim Preserve GridDtm(RowCount, columns)
                res = True
            Else
                'No resize applicable
                res = False
            End If
        End If
        PntAll = RowCount * ColCount
        CountValidPoints()
        Return res
    End Function

    Public Function SetElevation(row As Integer, column As Integer, value As Double) As Boolean
        Dim res As Boolean
        If row >= 0 And row < RowCount Then
            If column >= 0 And column < ColCount Then
                GridDtm(row, column) = value
                res = True
            Else
                res = False
            End If
        Else
            res = False
        End If
        CountValidPoints()
        Return res
    End Function

    Public Sub SetGridOrigin(X As Double, Y As Double)
        If IsFinite(X) Then
            If IsFinite(Y) Then
                MinCoords.X = X
                MinCoords.Y = Y
            Else
                MinCoords.X = X
            End If
        Else
            If IsFinite(Y) Then
                MinCoords.Y = Y
            Else
                'Input not valid
            End If
        End If
    End Sub

End Class