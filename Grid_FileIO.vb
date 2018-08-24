Imports System.IO

Partial Public Class Grid

    Public Sub New()
        PntAll = 0
        PntOk = 0
        'LoadOk = False
        Name = "Grid[" & Date.Now.ToString("yyyyMMdd-HHmmss-fff", Globalization.CultureInfo.InvariantCulture) & "]"

    End Sub

    Public Sub New(GridStream As Stream)
        'Dim fs As New IO.FileStream(FileName, IO.FileMode.Open, IO.FileAccess.Read)
        Dim empty32 As Int32
        Dim CountRow, CountCol As Integer
        Dim CellValueSng, CellNullSng As Single
        Dim emptyDbl, CellNullDbl, CellValueDbl As Double
        Dim HdrString As String
        'Dim fInfo As New IO.FileInfo(FileName)
        'Dim FileLen As Int64 = fInfo.Length
        'Dim PrevProg As Integer = 0
        'Dim ActProg As Integer = 0
        'LoadOk = False
        Name = "Grid[" & Date.Now.ToString("yyyyMMdd-HHmmss-fff", Globalization.CultureInfo.InvariantCulture) & "]"

        Using LatFile As New IO.BinaryReader(GridStream)
            'Retrieve the first 4 byte for grid type identification
            HdrString = LatFile.ReadChars(4) ' Header string
            'Load the grid file as required
            Select Case HdrString
                Case "LKPG" 'Geopak Lattice File
                    'Read as Geopak Lattice
                    empty32 = LatFile.ReadInt32     '4
                    empty32 = LatFile.ReadInt32     '8
                    CountRow = LatFile.ReadInt32    '12
                    CountCol = LatFile.ReadInt32    '16
                    PntAll = LatFile.ReadInt32      '20
                    PntOk = LatFile.ReadInt32       '24
                    empty32 = LatFile.ReadInt32     '28
                    empty32 = LatFile.ReadInt32     '32
                    empty32 = LatFile.ReadInt32     '36
                    empty32 = LatFile.ReadInt32     '40
                    empty32 = LatFile.ReadInt32     '44
                    empty32 = LatFile.ReadInt32     '48
                    empty32 = LatFile.ReadInt32     '52
                    CellNullSng = LatFile.ReadSingle '56
                    empty32 = LatFile.ReadInt32     '60
                    GridInterval.Y = LatFile.ReadDouble     '64
                    GridInterval.X = LatFile.ReadDouble     '72
                    MinCoords.X = LatFile.ReadDouble       '80
                    MinCoords.Y = LatFile.ReadDouble       '88
                    MinCoords.Z = LatFile.ReadDouble       '96
                    emptyDbl = LatFile.ReadDouble       '104
                    emptyDbl = LatFile.ReadDouble       '112
                    MaxZ = LatFile.ReadDouble       '120
                    emptyDbl = LatFile.ReadDouble   '128 EastRange
                    emptyDbl = LatFile.ReadDouble   '136 NorthRange
                    emptyDbl = LatFile.ReadDouble   '144 ElevationRange

                    'Prepare the grid array
                    ReDim GridDtm(CountRow - 1, CountCol - 1)

                    'Jump to start data offset
                    For i = 0 To 81
                        CellValueSng = LatFile.ReadSingle
                    Next
                    'Start the loop for the columns
                    Try
                        For row = 0 To CountRow - 1
                            'Start the loop for the rows
                            For col = 0 To CountCol - 1
                                'Read Z value
                                CellValueSng = LatFile.ReadSingle
                                'Check if is a valid value
                                If CellValueSng = CellNullSng Then
                                    GridDtm(row, col) = Double.NaN
                                Else
                                    GridDtm(row, col) = CDbl(CellValueSng)
                                End If
                            Next
                        Next
                        'LoadOk = True
                    Catch ex As Exception
                        'LoadOk = False

                    End Try

                Case "DSRB" 'Surfer 7 File
                    Dim HeadSize, HeadVer As Int32
                    Dim emptyByte() As Byte
                    Dim grdok, dataok As Boolean
                    grdok = False
                    dataok = False

                    'Read as Surfer 7 File
                    HeadSize = LatFile.ReadInt32     '4
                    HeadVer = LatFile.ReadInt32      '8
                    If HeadSize > 4 Then emptyByte = LatFile.ReadBytes(HeadSize - 4)
                    'Loop through the sections of the file
                    Do
                        'Read next section header
                        HdrString = LatFile.ReadChars(4) ' Header string
                        HeadSize = LatFile.ReadInt32

                        Select Case HdrString
                            Case "GRID"
                                CountRow = LatFile.ReadInt32
                                CountCol = LatFile.ReadInt32
                                MinCoords.X = LatFile.ReadDouble
                                MinCoords.Y = LatFile.ReadDouble
                                GridInterval.X = LatFile.ReadDouble
                                GridInterval.Y = LatFile.ReadDouble
                                MinCoords.Z = LatFile.ReadDouble
                                MaxZ = LatFile.ReadDouble
                                emptyDbl = LatFile.ReadDouble 'Rotation Not Used as per format specification
                                CellNullDbl = LatFile.ReadDouble
                                'Calculate other values
                                PntAll = CountCol * CountRow
                                PntOk = 0
                                'MaxCoords.X = MinCoords.X + ((CountCol - 1) * GridInterval.X)
                                'MaxCoords.Y = MinCoords.Y + ((CountRow - 1) * GridInterval.Y)
                                'Prepare the grid array
                                ReDim GridDtm(CountRow - 1, CountCol - 1)
                                grdok = True
                                If dataok Then
                                    Exit Do
                                End If
                                If HeadSize > 72 Then emptyByte = LatFile.ReadBytes(HeadSize - 72)

                            Case "DATA"
                                'Start the loop for the columns
                                Try
                                    For row = 0 To CountRow - 1
                                        'Start the loop for the rows
                                        For col = 0 To CountCol - 1
                                            'Read Z value
                                            CellValueDbl = LatFile.ReadDouble
                                            'Check if is a valid value
                                            If CellValueDbl >= CellNullDbl Then
                                                GridDtm(row, col) = Double.NaN
                                            Else
                                                GridDtm(row, col) = CellValueDbl
                                                PntOk = PntOk + 1
                                            End If
                                        Next
                                    Next
                                    dataok = True
                                    If grdok Then
                                        'LoadOk = True
                                        Exit Do
                                    End If
                                Catch ex As Exception
                                    'LoadOk = False
                                    Exit Do
                                End Try

                            Case Else 'Fault info
                                emptyByte = LatFile.ReadBytes(HeadSize)

                        End Select

                    Loop

                Case Else
                    ' not compatible grid file
                    'LoadOk = False

            End Select
        End Using

    End Sub

    Public Function ToXml() As String
        Dim resXdoc As New XDocument()
        Dim xRoot As New XElement("TerrainGrid")

        'MinCoords, MaxCoords, GridInterval
        Dim xHead As New XElement("Header")

        xHead.Add(New XElement("MinCoords", MinCoords.X & "|" & MinCoords.Y & "|" & MinCoords.Z))
        xHead.Add(New XElement("MaxCoords", EastMax & "|" & NorthMax & "|" & MaxZ))
        xHead.Add(New XElement("Interval", GridInterval.X & "|" & GridInterval.Y))
        xHead.Add(New XElement("Dimension", ColCount & "|" & RowCount))

        Dim xmlRow As String

        Dim msRowZip As New MemoryStream
        Try
            Dim msRow As New MemoryStream
            Try
                Using binRow As New BinaryWriter(msRow)
                    For r = 0 To RowCount - 1
                        For c = 0 To ColCount - 1
                            binRow.Write(GridDtm(r, c))
                        Next
                    Next
                    Using gzip As Compression.DeflateStream = New Compression.DeflateStream(msRowZip, Compression.CompressionMode.Compress)
                        gzip.Write(msRow.ToArray, 0, CInt(msRow.Length))
                    End Using
                End Using
            Finally
                If msRow IsNot Nothing Then msRow.Dispose()
            End Try
            xmlRow = Convert.ToBase64String(msRowZip.ToArray)
        Finally
            If msRowZip IsNot Nothing Then msRowZip.Dispose()
        End Try

        Dim xrow As New XElement("GridData", xmlRow)

        xRoot.Add(xHead)
        xRoot.Add(xrow)

        resXdoc.Add(xRoot)
        resXdoc.Declaration = New XDeclaration("1.0", "UTF-8", "yes")
        Return resXdoc.ToString(SaveOptions.None)

    End Function

    Public Function ToXmlDocument() As XDocument
        Dim resXdoc As New XDocument()
        Dim xRoot As New XElement("TerrainGrid")

        'MinCoords, MaxCoords, GridInterval
        Dim xHead As New XElement("Header")

        xHead.Add(New XElement("MinCoords", MinCoords.X & "|" & MinCoords.Y & "|" & MinCoords.Z))
        xHead.Add(New XElement("MaxCoords", EastMax & "|" & NorthMax & "|" & MaxZ))
        xHead.Add(New XElement("Interval", GridInterval.X & "|" & GridInterval.Y))
        xHead.Add(New XElement("Dimension", ColCount & "|" & RowCount))

        Dim xmlRow As String

        Dim msRowZip As New MemoryStream
        Try
            Dim msRow As New MemoryStream
            Try
                Using binRow As New BinaryWriter(msRow)
                    For r = 0 To RowCount - 1
                        For c = 0 To ColCount - 1
                            binRow.Write(GridDtm(r, c))
                        Next
                    Next
                    Using gzip As Compression.DeflateStream = New Compression.DeflateStream(msRowZip, Compression.CompressionMode.Compress)
                        gzip.Write(msRow.ToArray, 0, CInt(msRow.Length))
                    End Using
                End Using
            Finally
                If msRow IsNot Nothing Then msRow.Dispose()
            End Try
            xmlRow = Convert.ToBase64String(msRowZip.ToArray)
        Finally
            If msRowZip IsNot Nothing Then msRowZip.Dispose()
        End Try

        Dim xrow As New XElement("GridData", xmlRow)

        xRoot.Add(xHead)
        xRoot.Add(xrow)

        resXdoc.Add(xRoot)
        resXdoc.Declaration = New XDeclaration("1.0", "UTF-8", "yes")
        Return resXdoc

    End Function

    Public Sub New(xDtm As XDocument)
        Dim resGrid As New Grid
        Dim tmpValues() As String
        Dim xnav As New XDocument()
        xnav = xDtm
        Dim xRoot As XElement = xnav.Root
        Dim load0, load1, load2, loadMin, loadMax, loadInt, loadCount, loadGrid As Boolean
        Dim CountRow, CountCol As Integer

        load0 = False
        load1 = False
        load2 = False

        Dim xHead = xRoot.Element("Header")
        Dim xrow = xRoot.Element("GridData")

        If xHead.HasElements Then
            If xHead.Element("MinCoords") IsNot Nothing Then
                tmpValues = xHead.Element("MinCoords").Value.Split("|"c)
                If tmpValues.Count = 3 Then
                    load0 = Double.TryParse(tmpValues(0), resGrid.MinCoords.X)
                    load1 = Double.TryParse(tmpValues(1), resGrid.MinCoords.Y)
                    load2 = Double.TryParse(tmpValues(2), resGrid.MinCoords.Z)
                End If
                If load0 And load1 And load2 Then
                    loadMin = True
                Else
                    loadMin = False
                End If
            Else
                loadMin = False
            End If

            If xHead.Element("MaxCoords") IsNot Nothing Then
                tmpValues = xHead.Element("MaxCoords").Value.Split("|"c)
                If tmpValues.Count = 3 Then
                    'load0 = Double.TryParse(tmpValues(0), resGrid.MaxCoords.X)
                    'load1 = Double.TryParse(tmpValues(1), resGrid.MaxCoords.Y)
                    load0 = True
                    load1 = True
                    load2 = Double.TryParse(tmpValues(2), resGrid.MaxZ)
                End If
                If load0 And load1 And load2 Then
                    loadMax = True
                Else
                    loadMax = False
                End If
            Else
                loadMax = False
            End If

            If xHead.Element("Interval") IsNot Nothing Then
                tmpValues = xHead.Element("Interval").Value.Split("|"c)
                If tmpValues.Count = 2 Then
                    load0 = Double.TryParse(tmpValues(0), resGrid.GridInterval.X)
                    load1 = Double.TryParse(tmpValues(1), resGrid.GridInterval.Y)
                End If
                If load0 And load1 Then
                    loadInt = True
                Else
                    loadInt = False
                End If
            Else
                loadInt = False
            End If

            If xHead.Element("Dimension") IsNot Nothing Then
                tmpValues = xHead.Element("Dimension").Value.Split("|"c)
                If tmpValues.Count = 2 Then
                    load0 = Integer.TryParse(tmpValues(0), CountCol)
                    load1 = Integer.TryParse(tmpValues(1), CountRow)
                End If
                If load0 And load1 Then
                    loadCount = True
                Else
                    loadCount = False
                End If
            Else
                loadCount = False
            End If

        End If

        ReDim resGrid.GridDtm(CountRow, CountCol)

        Dim grdBytes() As Byte

        If xHead.Element("GridData") IsNot Nothing Then
            Try
                grdBytes = Convert.FromBase64String(xHead.Element("GridData").Value)
                Dim msInput As New MemoryStream
                Try
                    msInput.Write(grdBytes, 0, grdBytes.Length)
                    msInput.Position = 0
                    Dim msOut As New MemoryStream
                    Try
                        Using gzip As New Compression.DeflateStream(msOut, Compression.CompressionMode.Decompress)
                            gzip.Write(msInput.ToArray, 0, CInt(msInput.Length))
                        End Using
                        Using grd As New BinaryReader(msOut)
                            For r = 0 To RowCount - 1
                                For c = 0 To ColCount - 1
                                    resGrid.GridDtm(r, c) = grd.ReadDouble()
                                Next
                            Next
                        End Using
                    Finally
                        If msOut IsNot Nothing Then msOut.Dispose()
                    End Try
                Finally
                    If msInput IsNot Nothing Then msInput.Dispose()
                End Try
                loadGrid = True
            Catch ex As Exception
                loadGrid = False
            End Try
        End If

        If loadCount And loadGrid And loadInt And loadMax And loadMin Then
            Me.GridDtm = resGrid.GridDtm
            Me.GridInterval = resGrid.GridInterval
            'Me.MaxCoords = resGrid.MaxCoords
            Me.MaxZ = resGrid.MaxZ
            Me.MinCoords = resGrid.MinCoords
            Me.PntAll = resGrid.PntAll
            Me.PntOk = resGrid.PntOk
        Else
            PntAll = 0
            PntOk = 0
        End If

    End Sub

    Public Function FromXml(xDtm As String) As Grid
        Dim resGrid As New Grid
        Dim tmpValues() As String
        Dim xnav As New XDocument()
        xnav = XDocument.Parse(xDtm)
        Dim xRoot As XElement = xnav.Root
        Dim load0, load1, load2, loadMin, loadMax, loadInt, loadCount, loadGrid As Boolean
        Dim CountRow, CountCol As Integer

        load0 = False
        load1 = False
        load2 = False

        Dim xHead = xRoot.Element("Header")
        Dim xrow = xRoot.Element("GridData")

        If xHead.HasElements Then
            If xHead.Element("MinCoords") IsNot Nothing Then
                tmpValues = xHead.Element("MinCoords").Value.Split("|"c)
                If tmpValues.Count = 3 Then
                    load0 = Double.TryParse(tmpValues(0), resGrid.MinCoords.X)
                    load1 = Double.TryParse(tmpValues(1), resGrid.MinCoords.Y)
                    load2 = Double.TryParse(tmpValues(2), resGrid.MinCoords.Z)
                End If
                If load0 And load1 And load2 Then
                    loadMin = True
                Else
                    loadMin = False
                End If
            Else
                loadMin = False
            End If

            If xHead.Element("MaxCoords") IsNot Nothing Then
                tmpValues = xHead.Element("MaxCoords").Value.Split("|"c)
                If tmpValues.Count = 3 Then
                    'load0 = Double.TryParse(tmpValues(0), resGrid.MaxCoords.X)
                    'load1 = Double.TryParse(tmpValues(1), resGrid.MaxCoords.Y)
                    load0 = True
                    load1 = True
                    load2 = Double.TryParse(tmpValues(2), resGrid.MaxZ)
                End If
                If load0 And load1 And load2 Then
                    loadMax = True
                Else
                    loadMax = False
                End If
            Else
                loadMax = False
            End If

            If xHead.Element("Interval") IsNot Nothing Then
                tmpValues = xHead.Element("Interval").Value.Split("|"c)
                If tmpValues.Count = 2 Then
                    load0 = Double.TryParse(tmpValues(0), resGrid.GridInterval.X)
                    load1 = Double.TryParse(tmpValues(1), resGrid.GridInterval.Y)
                End If
                If load0 And load1 Then
                    loadInt = True
                Else
                    loadInt = False
                End If
            Else
                loadInt = False
            End If

            If xHead.Element("Dimension") IsNot Nothing Then
                tmpValues = xHead.Element("Dimension").Value.Split("|"c)
                If tmpValues.Count = 2 Then
                    load0 = Integer.TryParse(tmpValues(0), CountCol)
                    load1 = Integer.TryParse(tmpValues(1), CountRow)
                End If
                If load0 And load1 Then
                    loadCount = True
                Else
                    loadCount = False
                End If
            Else
                loadCount = False
            End If

        End If

        ReDim resGrid.GridDtm(CountRow, CountCol)

        Dim grdBytes() As Byte

        If xHead.Element("GridData") IsNot Nothing Then
            Try
                grdBytes = Convert.FromBase64String(xHead.Element("GridData").Value)
                Using msInput As New MemoryStream
                    msInput.Write(grdBytes, 0, grdBytes.Length)
                    msInput.Position = 0
                    Dim msOut As New MemoryStream
                    Try
                        Using gzip As New Compression.DeflateStream(msOut, Compression.CompressionMode.Decompress)
                            gzip.Write(msInput.ToArray, 0, CInt(msInput.Length))
                        End Using
                        Using grd As New BinaryReader(msOut)
                            For r = 0 To RowCount - 1
                                For c = 0 To ColCount - 1
                                    resGrid.GridDtm(r, c) = grd.ReadDouble()
                                Next
                            Next
                        End Using
                    Finally
                        If msOut IsNot Nothing Then msOut.Dispose()
                    End Try
                End Using
                loadGrid = True
            Catch ex As Exception
                loadGrid = False
            End Try
        End If

        If loadCount And loadGrid And loadInt And loadMax And loadMin Then
            Return resGrid
        Else
            Return New Grid
        End If

    End Function

End Class