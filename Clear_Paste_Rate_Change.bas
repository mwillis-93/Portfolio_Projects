Attribute VB_Name = "Clear_Paste_Rate_Change"
Option Explicit

Sub Clear_Rate_Change()

    Dim ws As Worksheet
    Dim StartCell As Range
    Dim lastcol As Long, CurrRow As Long, lastrow As Long
    Dim EmptyCell As Integer
    Dim DataRng As Range
    
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "Change by Component", vbTextCompare) > 0 Then
            Set StartCell = ws.Range("B5")
    
            'Find last column
            lastcol = StartCell.Column
            Do While ws.Cells(StartCell.Row, lastcol + 1).Value <> ""
                lastcol = lastcol + 1
            Loop
            lastcol = lastcol
            
            'Move down until 2 successive empty rows
            CurrRow = StartCell.Row
            EmptyCell = 0
            
            Do While EmptyCell < 3
                CurrRow = CurrRow + 1
                If WorksheetFunction.CountA(ws.Range(ws.Cells(CurrRow, StartCell.Column), ws.Cells(CurrRow, lastcol))) = 0 Then
                    EmptyCell = EmptyCell + 1
                Else
                    EmptyCell = 0
                End If
            Loop
            
            'Clear selected data range
            lastrow = CurrRow - 2
            Set DataRng = ws.Range(StartCell, ws.Cells(lastrow, lastcol))
            DataRng.ClearContents
        End If
    Next ws
    
    MsgBox "Done Clearing Rate Change"
    
End Sub

Sub Paste_Rate_Change()
    
    Dim UserInput As String     'Grabs MM-YYYY to search for desired updated rate changes
    Dim YearInput As String     'Seperates year from UserInput to compile file path
    Dim ws As Worksheet
    Dim FilePath As String      'Gives consistent pieces of the file path
    Dim FullPath As String      'Compiles file path with MM-YYYY and YYYY
    Dim ST As String            'State abbreviation for finding correct tabs and rate change files by state
    Dim STRateChange As String  'State-specific xlsx file
    Dim TabMatch As String      'Matches the Specific Change by Component tab to pull data from
    Dim YrQtr As String         'Matches the Year and Quarter data in workbook ("B4")
    Dim YrQtrMatch As Variant   'Used in match function to find Year and Quarter data in updated workbook
        
    Dim wb As Workbook
    Dim WbNew As Workbook       'State-specific rate change workbook
    
    Dim wsSRC As Worksheet      'State-specific and tab-specific worksheet
    Dim CopyCell As Range
    Dim lastcol As Long, CurrRow As Long, lastrow As Long
    Dim EmptyCell As Integer
    Dim DataRng As Range        'Data to be copied from updated workbook
    
    Dim wsPaste As Worksheet
    
    UserInput = InputBox("Enter the month and year (MM-YYYY):", "Open Workbook")
    
    If UserInput = "" Then Exit Sub
    
    If Len(UserInput) <> 7 Or Mid(UserInput, 3, 1) <> "-" Then
        MsgBox "Enter in MM-YYYY format."
        Exit Sub
    End If
    
    YearInput = Right(UserInput, 4)
    
    'Find state abrreviation and compile filepath
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "Change by Component", vbTextCompare) > 0 Then
            ST = UCase(Right(Trim(ws.Name), 2))
            TabMatch = ws.Range("A2").Value
            YrQtr = ws.Range("B4").Value
            ws.Range("C1").Value = Date
    
            FilePath = "network:\filepath\" & YearInput & "\" & UserInput & "\"
            STRateChange = Dir(FilePath & "*" & ST & "*.xlsx")
    
            If STRateChange <> "" Then
                FullPath = FilePath & STRateChange
                Set WbNew = Workbooks.Open(FullPath)
            
                'Select correct range of data from updated workbook
                For Each wsSRC In WbNew.Worksheets
                    If StrComp(TabMatch, wsSRC.Range("A2"), vbBinaryCompare) = 0 Then
                        wsSRC.Activate
                    End If
                Next wsSRC
                Set wsSRC = WbNew.ActiveSheet
                YrQtrMatch = Application.Match(YrQtr, wsSRC.Rows(4), 0)
                    Set CopyCell = wsSRC.Rows(4).Cells(YrQtrMatch)
                    
                    'Find last column
                    lastcol = CopyCell.Column
                    Do While Cells(CopyCell.Row, lastcol + 1).Value <> ""
                        lastcol = lastcol + 1
                    Loop
                    
                    'Move down until 3 successive empty rows
                    CurrRow = CopyCell.Row
                    EmptyCell = 0
                    
                    Do While EmptyCell < 3
                        CurrRow = CurrRow + 1
                        If WorksheetFunction.CountA(wsSRC.Range(wsSRC.Cells(CurrRow, CopyCell.Column), wsSRC.Cells(CurrRow, lastcol))) = 0 Then
                            EmptyCell = EmptyCell + 1
                        Else
                            EmptyCell = 0
                        End If
                    Loop
                'Copy selected data range
                lastrow = CurrRow - 3
                Set DataRng = wsSRC.Range(CopyCell, wsSRC.Cells(lastrow, lastcol))
                DataRng.Copy
                
                'Paste in practice workbook
                For Each wsPaste In ThisWorkbook.Worksheets
                    If InStr(1, wsPaste.Name, ST, vbTextCompare) > 0 Then
                        wsPaste.Range("B4").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                        wsPaste.Range("B4").PasteSpecial Paste:=xlPasteFormats
                        Exit For
                    End If
                Next wsPaste
                Set wsPaste = ThisWorkbook.ActiveSheet
                          
                'Copy Column A headers, in case of change
                wsSRC.Range("A5:A100").Copy
                wsPaste.Range("A5").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                wsPaste.Range("A5").PasteSpecial Paste:=xlPasteFormats
                
            'Close updated workbooks without saving
            Application.DisplayAlerts = False
            WbNew.Close SaveChanges:=False
            Application.DisplayAlerts = True
            
            End If
        End If
    Next ws
        
End Sub

