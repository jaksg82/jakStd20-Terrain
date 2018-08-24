Imports jakStd20_MathExt

Public Class Grid
    Dim GridDtm(,) As Double 'GridDtm(row, col)

    'Dim Xmin, Xmax, Ymin, Ymax, Zmin, Zmax, ColInt, RowInt As Double
    Dim PntOk, PntAll As Integer

    'Dim LoadOk As Boolean
    Dim MinCoords, GridInterval As New Point3D

    Dim MaxZ As Double

    ReadOnly Property EastMin As Double
        Get
            Return MinCoords.X
        End Get
    End Property

    ReadOnly Property NorthMin As Double
        Get
            Return MinCoords.Y
        End Get
    End Property

    ReadOnly Property ElevationMin As Double
        Get
            Return MinCoords.Z
        End Get
    End Property

    ReadOnly Property EastMax As Double
        Get
            Return MinCoords.X + ((ColCount - 1) * GridInterval.X)
        End Get
    End Property

    ReadOnly Property NorthMax As Double
        Get
            Return MinCoords.Y + ((RowCount - 1) * GridInterval.Y)
        End Get
    End Property

    ReadOnly Property ElevationMax As Double
        Get
            Return MaxZ
        End Get
    End Property

    ReadOnly Property EastRange As Double
        Get
            Return ((ColCount - 1) * GridInterval.X)
        End Get
    End Property

    ReadOnly Property NorthRange As Double
        Get
            Return ((RowCount - 1) * GridInterval.Y)
        End Get
    End Property

    ReadOnly Property ElevationRange As Double
        Get
            Return MaxZ - MinCoords.Z
        End Get
    End Property

    ReadOnly Property AllCells As Integer
        Get
            Return PntAll
        End Get
    End Property

    ReadOnly Property ValidCells As Integer
        Get
            Return PntOk
        End Get
    End Property

    ReadOnly Property RowCount As Integer
        Get
            Return GridDtm.GetUpperBound(0) + 1
        End Get
    End Property

    Public Property RowInterval As Double
        Get
            Return GridInterval.Y
        End Get
        Set(value As Double)
            If IsFinite(value) Then
                GridInterval.Y = value
            End If
        End Set
    End Property

    ReadOnly Property ColCount As Integer
        Get
            Return GridDtm.GetUpperBound(1) + 1
        End Get
    End Property

    Public Property ColumnInterval As Double
        Get
            Return GridInterval.X
        End Get
        Set(value As Double)
            If IsFinite(value) Then
                GridInterval.X = value
            End If
        End Set
    End Property

    'ReadOnly Property IsGridLoaded As Boolean
    '    Get
    '        Return LoadOk
    '    End Get
    'End Property

    Public Property Name As String

    Private Sub CountValidPoints()
        Dim pnt As Integer = 0
        Dim tmpMin, tmpMax As Double
        tmpMax = Double.MinValue
        tmpMin = Double.MaxValue
        For c = 0 To ColCount - 1
            For r = 0 To RowCount - 1
                If IsFinite(GridDtm(r, c)) Then
                    pnt = pnt + 1
                    tmpMin = Math.Min(tmpMin, GridDtm(r, c))
                    tmpMax = Math.Max(tmpMax, GridDtm(r, c))
                End If
            Next
        Next
        PntOk = pnt
        MinCoords.Z = tmpMin
        MaxZ = tmpMax
    End Sub

    Public Function GetCellElevation(Row As Integer, Column As Integer) As Double
        If Row >= 0 And Row < RowCount Then
            If Column >= 0 And Column < ColCount Then
                Return GridDtm(Row, Column)
            Else
                Return Double.NaN
            End If
        Else
            Return Double.NaN
        End If
    End Function

    Public Function ToPoint3D(Row As Integer, Column As Integer) As Point3D
        Dim tmpPnt As New Point3D

        If Row >= 0 And Row < RowCount Then
            If Column >= 0 And Column < ColCount Then
                tmpPnt.X = MinCoords.X + GridInterval.X * Column
                tmpPnt.Y = MinCoords.Y + GridInterval.Y * Row
                tmpPnt.Z = GridDtm(Row, Column)
                Return tmpPnt
            Else
                Return tmpPnt
            End If
        Else
            Return tmpPnt
        End If
    End Function

    Public Function GetPointElevation(RefPoint As Point3D) As Double
        Dim SelCell(4), Triangle1(2), Triangle2(2), Triangle3(2), Triangle4(2) As Point3D 'the 4 dtm points nearer to the refpoint
        Dim SelCol, SelRow As Integer
        Dim CalcElev As Double

        'Check if the point are inside the grid
        If RefPoint.X > MinCoords.X And RefPoint.X < EastMax Then
            If RefPoint.Y > MinCoords.Y And RefPoint.Y < NorthMax Then
                'Retrieve the nearest column and row
                SelCol = CInt(Math.Floor((RefPoint.X - MinCoords.X) / GridInterval.X))
                SelRow = CInt(Math.Floor((RefPoint.Y - MinCoords.Y) / GridInterval.Y))
                'Prepare the array of point for the computation
                ' c0     c1
                '    c4
                ' c3     c2
                SelCell(0).X = MinCoords.X + (GridInterval.X * SelCol)
                SelCell(0).Y = MinCoords.Y + (GridInterval.Y * (SelRow + 1))
                SelCell(0).Z = GridDtm(SelRow + 1, SelCol)
                SelCell(1).X = MinCoords.X + (GridInterval.X * (SelCol + 1))
                SelCell(1).Y = MinCoords.Y + (GridInterval.Y * (SelRow + 1))
                SelCell(1).Z = GridDtm(SelRow + 1, SelCol + 1)
                SelCell(2).X = MinCoords.X + (GridInterval.X * (SelCol + 1))
                SelCell(2).Y = MinCoords.Y + (GridInterval.Y * SelRow)
                SelCell(2).Z = GridDtm(SelRow, SelCol + 1)
                SelCell(3).X = MinCoords.X + (GridInterval.X * SelCol)
                SelCell(3).Y = MinCoords.Y + (GridInterval.Y * SelRow)
                SelCell(3).Z = GridDtm(SelRow, SelCol)
                'Calculate how many Z value are available
                Dim ZAvail As Integer = 0
                ZAvail = ZAvail + If(Double.IsNaN(SelCell(0).Z), 1000, 2000)
                ZAvail = ZAvail + If(Double.IsNaN(SelCell(1).Z), 100, 200)
                ZAvail = ZAvail + If(Double.IsNaN(SelCell(2).Z), 10, 20)
                ZAvail = ZAvail + If(Double.IsNaN(SelCell(3).Z), 1, 2)

                Select Case ZAvail
                    Case 2222
                        'calculate the fourth point
                        SelCell(4).X = MinCoords.X + (GridInterval.X * SelCol) + (GridInterval.X / 2)
                        SelCell(4).Y = MinCoords.Y + (GridInterval.Y * SelRow) + (GridInterval.Y / 2)
                        SelCell(4).Z = (((SelCell(1).Z + SelCell(3).Z) / 2) + ((SelCell(0).Z + SelCell(2).Z) / 2)) / 2
                        'Calculate in witch of the 4 sub-cell triangles lay the point
                        If IsPointInsideTriangle(SelCell(0), SelCell(1), SelCell(4), RefPoint) Then
                            'the point is inside the first triangle
                            CalcElev = ExtractElevation(SelCell(0), SelCell(1), SelCell(4), RefPoint)
                        Else
                            If IsPointInsideTriangle(SelCell(1), SelCell(2), SelCell(4), RefPoint) Then
                                'the point is inside the second triangle
                                CalcElev = ExtractElevation(SelCell(1), SelCell(2), SelCell(4), RefPoint)
                            Else
                                If IsPointInsideTriangle(SelCell(2), SelCell(3), SelCell(4), RefPoint) Then
                                    'the point is inside the third triangle
                                    CalcElev = ExtractElevation(SelCell(2), SelCell(3), SelCell(4), RefPoint)
                                Else
                                    If IsPointInsideTriangle(SelCell(3), SelCell(0), SelCell(4), RefPoint) Then
                                        'the point is inside the fourth triangle
                                        CalcElev = ExtractElevation(SelCell(3), SelCell(0), SelCell(4), RefPoint)
                                    Else
                                        '??????
                                        CalcElev = Double.NaN
                                    End If
                                End If
                            End If
                        End If

                    Case 2221
                        If IsPointInsideTriangle(SelCell(0), SelCell(1), SelCell(2), RefPoint) Then
                            'the point is inside the fourth triangle
                            CalcElev = ExtractElevation(SelCell(0), SelCell(1), SelCell(2), RefPoint)
                        Else
                            '??????
                            CalcElev = Double.NaN
                        End If

                    Case 2212
                        If IsPointInsideTriangle(SelCell(0), SelCell(1), SelCell(3), RefPoint) Then
                            'the point is inside the fourth triangle
                            CalcElev = ExtractElevation(SelCell(0), SelCell(1), SelCell(3), RefPoint)
                        Else
                            '??????
                            CalcElev = Double.NaN
                        End If

                    Case 2122
                        If IsPointInsideTriangle(SelCell(0), SelCell(2), SelCell(3), RefPoint) Then
                            'the point is inside the fourth triangle
                            CalcElev = ExtractElevation(SelCell(0), SelCell(2), SelCell(3), RefPoint)
                        Else
                            '??????
                            CalcElev = Double.NaN
                        End If

                    Case 1222
                        If IsPointInsideTriangle(SelCell(1), SelCell(2), SelCell(3), RefPoint) Then
                            'the point is inside the fourth triangle
                            CalcElev = ExtractElevation(SelCell(1), SelCell(2), SelCell(3), RefPoint)
                        Else
                            '??????
                            CalcElev = Double.NaN
                        End If

                    Case Else
                        CalcElev = Double.NaN

                End Select
            Else
                'Point outside of grid
                CalcElev = Double.NaN
            End If
        Else
            'Point outside of grid
            CalcElev = Double.NaN
        End If

        Return CalcElev
    End Function

    Public Function GetValidPoints() As List(Of Point3D)
        Dim TmpArray As New List(Of Point3D)
        Dim cellvalue As Double
        Dim enow As Double
        Dim nnow As Double
        Dim orige As Double
        Dim orign As Double
        'Dim PntCount As Integer

        orige = MinCoords.X
        orign = MinCoords.Y
        enow = orige - GridInterval.X
        nnow = orign - GridInterval.Y

        'Start the loop for the columns
        For row = 0 To RowCount - 1
            nnow = nnow + GridInterval.Y
            enow = orige
            'Start the loop for the rows
            For col = 0 To ColCount - 1
                'Read Z value
                cellvalue = GetCellElevation(row, col)
                'Check if is a valid value
                If Double.IsNaN(cellvalue) Then
                    enow = enow + GridInterval.X
                Else
                    TmpArray.Add(New Point3D(enow, nnow, cellvalue))
                    enow = enow + GridInterval.X
                End If
            Next
        Next

        'Return the points
        Return TmpArray

    End Function

    Private Function IsPointInsideTriangle(V1 As Point3D, V2 As Point3D, V3 As Point3D, RefPoint As Point3D) As Boolean
        'float alpha = ((p2.y - p3.y)*(p.x - p3.x) + (p3.x - p2.x)*(p.y - p3.y)) / ((p2.y - p3.y)*(p1.x - p3.x) + (p3.x - p2.x)*(p1.y - p3.y));
        'float beta = ((p3.y - p1.y)*(p.x - p3.x) + (p1.x - p3.x)*(p.y - p3.y)) / ((p2.y - p3.y)*(p1.x - p3.x) + (p3.x - p2.x)*(p1.y - p3.y));
        'float gamma = 1.0f - alpha - beta;
        Dim Alpha, Beta, Gamma As Double
        Dim P0 As New Point3D
        P0 = RefPoint

        Alpha = ((V2.Y - V3.Y) * (P0.X - V3.X) + (V3.X - V2.X) * (P0.Y - V3.Y)) / ((V2.Y - V3.Y) * (V1.X - V3.X) + (V3.X - V2.X) * (V1.Y - V3.Y))
        Beta = ((V3.Y - V1.Y) * (P0.X - V3.X) + (V1.X - V3.X) * (P0.Y - V3.Y)) / ((V2.Y - V3.Y) * (V1.X - V3.X) + (V3.X - V2.X) * (V1.Y - V3.Y))
        Gamma = 1.0 - Alpha - Beta

        If Alpha >= 0 And Beta >= 0 And Gamma >= 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function ExtractElevation(V1 As Point3D, V2 As Point3D, V3 As Point3D, RefPoint As Point3D) As Double
        Dim P0, X1, X2, X3 As New Point3D
        Dim Z1 As Double
        Dim V1V2, V2V3, V3V1 As Boolean

        P0 = RefPoint

        'Calculate the intersection points with the same X of refpoint
        X1.X = P0.X
        X2.X = P0.X
        X3.X = P0.X

        'First side V1V2
        If V1.X - V2.X = 0 Then
            'Side V1-V2 have the same East
            X1.Y = Double.NaN
            X1.Z = Double.NaN
            V1V2 = False
        Else
            If V1.Y - V2.Y = 0 Then
                'Side V1-V2 have the same North
                V1V2 = True
                X1.Y = V1.Y
                X1.Z = V1.Z + ((P0.X - V1.X) * (V2.Z - V1.Z) / (V2.X - V1.X))
            Else
                'Side V1-V2 are inclined
                V1V2 = True
                X1.Y = V1.Y + ((P0.X - V1.X) * (V2.Y - V1.Y) / (V2.X - V1.X))
                X1.Z = V1.Z + ((X1.Y - V1.Y) * (V2.Z - V1.Z) / (V2.Y - V1.Y))
            End If
        End If

        'Second side V2V3
        If V3.X - V2.X = 0 Then
            'Side V3-V2 have the same East
            X2.Y = Double.NaN
            X2.Z = Double.NaN
            V2V3 = False
        Else
            If V3.Y - V2.Y = 0 Then
                'Side V3-V2 have the same North
                V2V3 = True
                X2.Y = V2.Y
                X2.Z = V2.Z + ((P0.X - V2.X) * (V3.Z - V2.Z) / (V3.X - V2.X))
            Else
                'Side V3-V2 are inclined
                V2V3 = True
                X2.Y = V2.Y + ((P0.X - V2.X) * (V3.Y - V2.Y) / (V3.X - V2.X))
                X2.Z = V2.Z + ((X2.Y - V2.Y) * (V3.Z - V2.Z) / (V3.Y - V2.Y))
            End If
        End If

        'Third side V3V1
        If V1.X - V3.X = 0 Then
            'Side V1-V3 have the same East
            X3.Y = Double.NaN
            X3.Z = Double.NaN
            V3V1 = False
        Else
            If V1.Y - V3.Y = 0 Then
                'Side V1-V3 have the same North
                V3V1 = True
                X3.Y = V3.Y
                X3.Z = V3.Z + ((P0.X - V3.X) * (V1.Z - V3.Z) / (V1.X - V3.X))
            Else
                'Side V1-V3 are inclined
                V3V1 = True
                X3.Y = V1.Y + ((P0.X - V1.X) * (V3.Y - V1.Y) / (V3.X - V1.X))
                X3.Z = V1.Z + ((X3.Y - V1.Y) * (V3.Z - V1.Z) / (V3.Y - V1.Y))
            End If
        End If

        If V1V2 Then
            If V2V3 Then
                Z1 = X1.Z + ((P0.Y - X1.Y) * (X2.Z - X1.Z) / (X2.Y - X1.Y))
            Else
                Z1 = X1.Z + ((P0.Y - X1.Y) * (X3.Z - X1.Z) / (X3.Y - X1.Y))
            End If
        Else
            Z1 = X3.Z + ((P0.Y - X3.Y) * (X2.Z - X3.Z) / (X2.Y - X3.Y))
        End If

        Return Z1

    End Function

End Class