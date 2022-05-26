
Sub deleteConsecutiveRows()
    Dim wks As Excel.Worksheet
    Dim rng As Excel.Range
    Dim row As Long
    Dim lastRow As Long
    '-------------------------------------------------------------------------


    Set wks = Excel.ActiveSheet


    With wks
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).row

        For row = 2 To lastRow
            If .Cells(row, 4).Value = .Cells(row - 1, 4).Value Then

                If rng Is Nothing Then
                    Set rng = .Rows(row)
                Else
                    Set rng = Excel.Union(rng, .Rows(row))
                End If

            End If
        Next row

    End With


    'In order to avoid Run-time error check if [rng] range is not empty, before removing it.
    If Not rng Is Nothing Then
        Call rng.EntireRow.Delete
    End If