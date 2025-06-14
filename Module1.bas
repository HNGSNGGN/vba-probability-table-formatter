Attribute VB_Name = "Module1"
Sub Magic()

    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim rngSelection As Range
    Dim headerRow As Long
    Dim rankCol As Long, nameCol As Long, probCol As Long
    Dim startRow As Long, lastRow As Long
    Dim i As Long

    Dim totalRows As Long, thirdCount As Long

    On Error Resume Next
    Set rngSelection = Application.InputBox( _
        "Please select the header range for " & _
        ChrW(&H7B49) & ChrW(&H7D1A) & ", " & _
        ChrW(&H30A2) & ChrW(&H30A4) & ChrW(&H30C6) & ChrW(&H30E0) & ChrW(&H540D) & ", " & _
        ChrW(&H78BA) & ChrW(&H7387) & ":", Type:=8)
    On Error GoTo 0
    If rngSelection Is Nothing Then
        MsgBox "Operation canceled."
        Exit Sub
    End If

    Set wsSource = rngSelection.Worksheet
    headerRow = rngSelection.Row

    rankCol = 0: nameCol = 0: probCol = 0

    Dim strRankHeader As String
    Dim strNameHeader As String
    Dim strProbHeader As String

    strRankHeader = ChrW(&H7B49) & ChrW(&H7D1A)
    strNameHeader = ChrW(&H30A2) & ChrW(&H30A4) & ChrW(&H30C6) & ChrW(&H30E0) & ChrW(&H540D) ' «¢«¤«Æ«àÙ£
    strProbHeader = ChrW(&H78BA) & ChrW(&H7387)

    For i = 1 To rngSelection.Columns.Count
        Select Case wsSource.Cells(headerRow, rngSelection.Column + i - 1).Value
            Case strRankHeader
                rankCol = rngSelection.Column + i - 1
            Case strNameHeader
                nameCol = rngSelection.Column + i - 1
            Case strProbHeader
                probCol = rngSelection.Column + i - 1
        End Select
    Next i

    If rankCol = 0 Or nameCol = 0 Or probCol = 0 Then
        MsgBox "Required headers (" & strRankHeader & ", " & strNameHeader & ", " & strProbHeader & ") not found."
        Exit Sub
    End If

    startRow = headerRow + 1
    lastRow = wsSource.Cells(wsSource.Rows.Count, rankCol).End(xlUp).Row
    If lastRow < startRow Then
        MsgBox "No data found."
        Exit Sub
    End If

    totalRows = lastRow - startRow + 1
    If totalRows Mod 3 <> 0 Then
        MsgBox "Rows Number error."
        Exit Sub
    End If

    thirdCount = totalRows \ 3

    Dim firstJob() As Variant
    Dim secondJob() As Variant
    Dim thirdJob() As Variant

    ReDim firstJob(1 To thirdCount, 1 To 3)
    ReDim secondJob(1 To thirdCount, 1 To 3)
    ReDim thirdJob(1 To thirdCount, 1 To 3)

    Dim idx As Long
    idx = 1
    For i = startRow To startRow + thirdCount - 1
        firstJob(idx, 1) = wsSource.Cells(i, rankCol).Value
        firstJob(idx, 2) = wsSource.Cells(i, nameCol).Value
        firstJob(idx, 3) = wsSource.Cells(i, probCol).Text
        idx = idx + 1
    Next i

    idx = 1
    For i = startRow + thirdCount To startRow + (2 * thirdCount) - 1
        secondJob(idx, 1) = wsSource.Cells(i, rankCol).Value
        secondJob(idx, 2) = wsSource.Cells(i, nameCol).Value
        secondJob(idx, 3) = wsSource.Cells(i, probCol).Text
        idx = idx + 1
    Next i

    idx = 1
    For i = startRow + (2 * thirdCount) To lastRow
        thirdJob(idx, 1) = wsSource.Cells(i, rankCol).Value
        thirdJob(idx, 2) = wsSource.Cells(i, nameCol).Value
        thirdJob(idx, 3) = wsSource.Cells(i, probCol).Text
        idx = idx + 1
    Next i

    Dim wsResult As Worksheet
    Set wsResult = ThisWorkbook.Worksheets("3" & ChrW(&H30BB) & ChrW(&H30C3) & ChrW(&H30C8)) ' "3«»«Ã«È"

    wsResult.Range("G:AA").Clear

    wsResult.Range("G2").Value = ChrW(&H30E9) & ChrW(&H30F3) & ChrW(&H30AF)
    wsResult.Range("H2").Value = ChrW(&H30AF) & ChrW(&H30E9) & ChrW(&H30B9)
    wsResult.Range("I2").Value = strNameHeader
    wsResult.Range("J2").Value = ChrW(&H500B) & ChrW(&H5225) & ChrW(&H78BA) & ChrW(&H7387)

    wsResult.Range("K2").Value = ChrW(&H30AF) & ChrW(&H30E9) & ChrW(&H30B9)
    wsResult.Range("L2").Value = strNameHeader
    wsResult.Range("M2").Value = ChrW(&H500B) & ChrW(&H5225) & ChrW(&H78BA) & ChrW(&H7387)

    wsResult.Range("N2").Value = ChrW(&H30E9) & ChrW(&H30F3) & ChrW(&H30AF)
    wsResult.Range("O2").Value = ChrW(&H30AF) & ChrW(&H30E9) & ChrW(&H30B9)
    wsResult.Range("P2").Value = strNameHeader
    wsResult.Range("Q2").Value = ChrW(&H500B) & ChrW(&H5225) & ChrW(&H78BA) & ChrW(&H7387)

    Dim rowIndex As Long
    Dim originalText As String, valText As String
    Dim decimalCount As Long, dotPos As Long
    Dim dblVal As Double
    Dim fmt As String

    For rowIndex = 1 To thirdCount
        wsResult.Cells(rowIndex + 2, 7).Value = firstJob(rowIndex, 1)
        wsResult.Cells(rowIndex + 2, 8).Value = ChrW(&H30D3) & ChrW(&H30B7) & ChrW(&H30E7) & ChrW(&H30C3) & ChrW(&H30D7)
        wsResult.Cells(rowIndex + 2, 9).Value = firstJob(rowIndex, 2)

        originalText = firstJob(rowIndex, 3)
        valText = Replace(originalText, "%", "")
        valText = Trim(valText)
        dotPos = InStr(valText, ".")
        If dotPos > 0 Then
            decimalCount = Len(valText) - dotPos
        Else
            decimalCount = 0
        End If

        If IsNumeric(valText) Then
            dblVal = CDbl(valText) / 100#
            fmt = "0"
            If decimalCount > 0 Then
                fmt = fmt & "." & String(decimalCount, "0")
            End If
            fmt = fmt & "%"

            wsResult.Cells(rowIndex + 2, 10).NumberFormat = fmt
            wsResult.Cells(rowIndex + 2, 10).Value = dblVal
        Else
            wsResult.Cells(rowIndex + 2, 10).NumberFormat = "@"
            wsResult.Cells(rowIndex + 2, 10).Value = originalText
        End If
    Next rowIndex

    For rowIndex = 1 To thirdCount
        wsResult.Cells(rowIndex + 2, 11).Value = ChrW(&H30D1) & ChrW(&H30E9) & ChrW(&H30C7) & ChrW(&H30A3) & ChrW(&H30F3)
        wsResult.Cells(rowIndex + 2, 12).Value = secondJob(rowIndex, 2)

        originalText = secondJob(rowIndex, 3)
        valText = Replace(originalText, "%", "")
        valText = Trim(valText)
        dotPos = InStr(valText, ".")
        If dotPos > 0 Then
            decimalCount = Len(valText) - dotPos
        Else
            decimalCount = 0
        End If

        If IsNumeric(valText) Then
            dblVal = CDbl(valText) / 100#
            fmt = "0"
            If decimalCount > 0 Then
                fmt = fmt & "." & String(decimalCount, "0")
            End If
            fmt = fmt & "%"

            wsResult.Cells(rowIndex + 2, 13).NumberFormat = fmt
            wsResult.Cells(rowIndex + 2, 13).Value = dblVal
        Else
            wsResult.Cells(rowIndex + 2, 13).NumberFormat = "@"
            wsResult.Cells(rowIndex + 2, 13).Value = originalText
        End If
    Next rowIndex

    For rowIndex = 1 To thirdCount
        wsResult.Cells(rowIndex + 2, 14).Value = thirdJob(rowIndex, 1)
        wsResult.Cells(rowIndex + 2, 15).Value = ChrW(&H30D0) & ChrW(&H30FC) & ChrW(&H30C9)
        wsResult.Cells(rowIndex + 2, 16).Value = thirdJob(rowIndex, 2)

        originalText = thirdJob(rowIndex, 3)
        valText = Replace(originalText, "%", "")
        valText = Trim(valText)
        dotPos = InStr(valText, ".")
        If dotPos > 0 Then
            decimalCount = Len(valText) - dotPos
        Else
            decimalCount = 0
        End If

        If IsNumeric(valText) Then
            dblVal = CDbl(valText) / 100#
            fmt = "0"
            If decimalCount > 0 Then
                fmt = fmt & "." & String(decimalCount, "0")
            End If
            fmt = fmt & "%"

            wsResult.Cells(rowIndex + 2, 17).NumberFormat = fmt
            wsResult.Cells(rowIndex + 2, 17).Value = dblVal
        Else
            wsResult.Cells(rowIndex + 2, 17).NumberFormat = "@"
            wsResult.Cells(rowIndex + 2, 17).Value = originalText
        End If
    Next rowIndex

    Dim lastDataRow As Long
    lastDataRow = (thirdCount) + 2

    wsResult.Cells.Font.Name = "Meiryo UI"
    wsResult.Cells.Font.Size = 11

    If lastDataRow > 2 Then
        wsResult.Range("G3:G" & lastDataRow).Font.Size = 10
        wsResult.Range("I3:I" & lastDataRow).Font.Size = 10
        wsResult.Range("J3:J" & lastDataRow).Font.Size = 10
        wsResult.Range("L3:L" & lastDataRow).Font.Size = 10
        wsResult.Range("M3:M" & lastDataRow).Font.Size = 10
        wsResult.Range("N3:N" & lastDataRow).Font.Size = 10
        wsResult.Range("P3:P" & lastDataRow).Font.Size = 10
        wsResult.Range("Q3:Q" & lastDataRow).Font.Size = 10
    End If

    wsResult.Columns("G").ColumnWidth = 13.67
    wsResult.Columns("H").ColumnWidth = 16.58
    wsResult.Columns("I").ColumnWidth = 28
    wsResult.Columns("J").ColumnWidth = 18
    wsResult.Columns("K").ColumnWidth = 16.58
    wsResult.Columns("L").ColumnWidth = 28
    wsResult.Columns("M").ColumnWidth = 18
    wsResult.Columns("N").ColumnWidth = 13.67
    wsResult.Columns("O").ColumnWidth = 16.58
    wsResult.Columns("P").ColumnWidth = 28
    wsResult.Columns("Q").ColumnWidth = 18

    wsResult.Rows.RowHeight = 16.5
    wsResult.Rows(2).RowHeight = 15

    Application.DisplayAlerts = False

    Dim startMerge As Long, prevVal As Variant, colLetter As String

    If 3 <= lastDataRow Then
        colLetter = "G"
        startMerge = 3
        prevVal = wsResult.Range(colLetter & 3).Value
        For i = 4 To lastDataRow
            If wsResult.Range(colLetter & i).Value <> prevVal Then
                If i - 1 > startMerge Then
                    With wsResult.Range(colLetter & startMerge & ":" & colLetter & (i - 1))
                        .Merge
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlCenter
                    End With
                End If
                startMerge = i
                prevVal = wsResult.Range(colLetter & i).Value
            End If
        Next i
        If lastDataRow > startMerge Then
            With wsResult.Range(colLetter & startMerge & ":" & colLetter & lastDataRow)
                .Merge
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
            End With
        End If
    End If

    If 3 <= lastDataRow Then
        colLetter = "H"
        startMerge = 3
        prevVal = wsResult.Range(colLetter & 3).Value
        For i = 4 To lastDataRow
            If wsResult.Range(colLetter & i).Value <> prevVal Then
                If i - 1 > startMerge Then
                    With wsResult.Range(colLetter & startMerge & ":" & colLetter & (i - 1))
                        .Merge
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlCenter
                    End With
                End If
                startMerge = i
                prevVal = wsResult.Range(colLetter & i).Value
            End If
        Next i
        If lastDataRow > startMerge Then
            With wsResult.Range(colLetter & startMerge & ":" & colLetter & lastDataRow)
                .Merge
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
            End With
        End If
    End If

    If 3 <= lastDataRow Then
        colLetter = "K"
        startMerge = 3
        prevVal = wsResult.Range(colLetter & 3).Value
        For i = 4 To lastDataRow
            If wsResult.Range(colLetter & i).Value <> prevVal Then
                If i - 1 > startMerge Then
                    With wsResult.Range(colLetter & startMerge & ":" & colLetter & (i - 1))
                        .Merge
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlCenter
                    End With
                End If
                startMerge = i
                prevVal = wsResult.Range(colLetter & i).Value
            End If
        Next i
        If lastDataRow > startMerge Then
            With wsResult.Range(colLetter & startMerge & ":" & colLetter & lastDataRow)
                .Merge
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
            End With
        End If
    End If

    If 3 <= lastDataRow Then
        colLetter = "N"
        startMerge = 3
        prevVal = wsResult.Range(colLetter & 3).Value
        For i = 4 To lastDataRow
            If wsResult.Range(colLetter & i).Value <> prevVal Then
                If i - 1 > startMerge Then
                    With wsResult.Range(colLetter & startMerge & ":" & colLetter & (i - 1))
                        .Merge
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlCenter
                    End With
                End If
                startMerge = i
                prevVal = wsResult.Range(colLetter & i).Value
            End If
        Next i
        If lastDataRow > startMerge Then
            With wsResult.Range(colLetter & startMerge & ":" & colLetter & lastDataRow)
                .Merge
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
            End With
        End If
    End If

    If 3 <= lastDataRow Then
        colLetter = "O"
        startMerge = 3
        prevVal = wsResult.Range(colLetter & 3).Value
        For i = 4 To lastDataRow
            If wsResult.Range(colLetter & i).Value <> prevVal Then
                If i - 1 > startMerge Then
                    With wsResult.Range(colLetter & startMerge & ":" & colLetter & (i - 1))
                        .Merge
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlCenter
                    End With
                End If
                startMerge = i
                prevVal = wsResult.Range(colLetter & i).Value
            End If
        Next i
        If lastDataRow > startMerge Then
            With wsResult.Range(colLetter & startMerge & ":" & colLetter & lastDataRow)
                .Merge
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
            End With
        End If
    End If

    Application.DisplayAlerts = True

    wsResult.Range("G2:Q" & lastDataRow).HorizontalAlignment = xlCenter

    With wsResult.Range("G2:Q" & lastDataRow).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With wsResult.Range("G2:Q" & lastDataRow)
        .VerticalAlignment = xlCenter
    End With

End Sub

Sub ShiftColumnValues()

    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim jValues As Variant
    Dim mValues As Variant
    Dim qValues As Variant
    Dim i As Long

    Set wsResult = ThisWorkbook.Worksheets("3" & ChrW(&H30BB) & ChrW(&H30C3) & ChrW(&H30C8))

    lastRow = wsResult.Cells(wsResult.Rows.Count, "J").End(xlUp).Row

    If lastRow < 3 Then
        MsgBox "No data found."
        Exit Sub
    End If

    jValues = wsResult.Range("J3:J" & lastRow).Value
    mValues = wsResult.Range("M3:M" & lastRow).Value
    qValues = wsResult.Range("Q3:Q" & lastRow).Value

    If IsArray(jValues) = False Or IsArray(mValues) = False Or IsArray(qValues) = False Then
        MsgBox "One or more columns have no data."
        Exit Sub
    End If

    If UBound(jValues, 1) <> UBound(mValues, 1) Or UBound(jValues, 1) <> UBound(qValues, 1) Then
        MsgBox "Data ranges do not match."
        Exit Sub
    End If

    For i = 1 To UBound(jValues, 1)
        If IsEmpty(jValues(i, 1)) Then jValues(i, 1) = ""
        If IsEmpty(mValues(i, 1)) Then mValues(i, 1) = ""
        If IsEmpty(qValues(i, 1)) Then qValues(i, 1) = ""

        wsResult.Cells(i + 2, "M").Value = CStr(jValues(i, 1))
        wsResult.Cells(i + 2, "Q").Value = CStr(mValues(i, 1))
        wsResult.Cells(i + 2, "J").Value = CStr(qValues(i, 1))
    Next i

End Sub

