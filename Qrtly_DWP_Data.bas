Attribute VB_Name = "Qrtly_DWP_Data"
Option Explicit

Sub Gen_Qtrly_Data()

    Dim SQL_Data As Worksheet, Qrtly_DWP As Worksheet
    Dim tbl As ListObject
    Dim i As Long, j As Long, yr As Long, Mnth As Long, Qtr As Long
    Dim ST As String
    Dim Mnth_WP_amt As Double
    Dim found As Boolean

    'Set up SQL data worksheet and table
    Set SQL_Data = Sheet1
    Set tbl = SQL_Data.ListObjects("DWPDistrib")

    'Get or create Qrtly_DWP worksheet
    On Error Resume Next
    Set Qrtly_DWP = Worksheets("Qrtly_DWP")
    On Error GoTo 0

    If Qrtly_DWP Is Nothing Then
        Set Qrtly_DWP = Sheets.Add
        Qrtly_DWP.Name = "Qrtly_DWP"
    End If

    'Set up headers (no WP% column)
    Qrtly_DWP.Range("A3:D3").Value = Array("EXPOSURE_STATE", "PolYear", "PolQuarter", "QrtlyWP$")

    'Clear previous data
    Qrtly_DWP.Range("A4:D" & Qrtly_DWP.Rows.Count).ClearContents

    'Defines results variables
    Dim results() As Variant
    Dim Row_Data(1 To 4) As Variant
    Dim Temp_Row(1 To 4) As Variant
    Dim Row_Count As Long
    Row_Count = 0

    'Loop through each row of the SQL table
    For i = 1 To tbl.ListRows.Count
        With tbl.ListRows(i).Range
            ST = .Cells(1, tbl.ListColumns("EXPOSURE_STATE").Index).Value
            yr = .Cells(1, tbl.ListColumns("PolYear").Index).Value
            Mnth = .Cells(1, tbl.ListColumns("PolMonth").Index).Value
            Mnth_WP_amt = .Cells(1, tbl.ListColumns("WP_Tot").Index).Value
        End With

        'Convert month to quarter
        Select Case Mnth
            Case 1 To 3: Qtr = 1
            Case 4 To 6: Qtr = 2
            Case 7 To 9: Qtr = 3
            Case 10 To 12: Qtr = 4
        End Select

        'Check if combination already exists
        found = False
        For j = 0 To Row_Count - 1
            If results(j)(1) = ST And results(j)(2) = yr And results(j)(3) = Qtr Then
                results(j)(4) = results(j)(4) + Mnth_WP_amt
                found = True
                Exit For
            End If
        Next j

        'If not found, add new row
        If Not found Then
            Row_Data(1) = ST
            Row_Data(2) = yr
            Row_Data(3) = Qtr
            Row_Data(4) = Mnth_WP_amt

            ReDim Preserve results(0 To Row_Count)

            For j = 1 To 4
                Temp_Row(j) = Row_Data(j)
            Next j

            results(Row_Count) = Temp_Row
            Row_Count = Row_Count + 1
        End If
    Next i

    'Output results
    If Row_Count > 0 Then
        Dim output() As Variant
        ReDim output(1 To Row_Count, 1 To 4)

        For i = 0 To Row_Count - 1
            For j = 1 To 4
                output(i + 1, j) = results(i)(j)
            Next j
        Next i

        Qrtly_DWP.Range("A4").Resize(Row_Count, 4).Value = output
    End If

End Sub

