# VBA-challenge
This module challenge uses VBA scripting to analyze generated stock market data.

Within this repository, you'll find individual screenshots illustrating the project's results, along with a PDF compilation for convenient reference. The VB file included here contains the full script utilized throughout the project. Additionally, due to its size, the macro-enabled 'multiple_year_stock_data' workbook is accessible via a Gmail link provided for your convenience. This workbook, complementing the VB script, offers a comprehensive overview of the project's functionality. Furthermore, a smaller test workbook has been included, complete with the full script, for experimentation and testing purposes.

Google Drive Link: https://drive.google.com/drive/folders/1Wkxq-lsB6J8iteFSvlgebRmOJS1qGXSe?usp=sharing



Module 2 Challenge: VBA

'Instructions
'Create a script that loops through all the stocks for one year and outputs the following information:
'The ticker symbol
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
'Use conditional formatting that will highlight positive change in green and negative change in red.

'Column Creation:
'On the same worksheet as the raw data, all columns created for:
'Ticker Symbol
'Total Stock Volume
'Yearly Change
'Percent Change
'Open Price (Additional column to calculate Yearly Change and Percent Change)
'Close Price(Additional column to calculate Yearly Change and Percent Change)
'Greatest % Increase
'Greatest % Decrease
'Greatest Total Volume

Sub ColumnCreation()
    Dim ws As Worksheet
    Dim SheetsArr() As Variant
    Dim sheetName As Variant
    
    'Array of sheets for headers
    SheetsArr = Array("2018", "2019", "2020")
    
    'Loop through each sheet in the array
    For Each sheetName In SheetsArr
        'Set the worksheet
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        'Add headers in the specified cells
        With ws
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"
            .Range("S1").Value = "Open Price"
            .Range("T1").Value = "Close Price"
        End With
    Next sheetName
End Sub

'Retrieval of Data:
'The script loops through one year of stock data and reads/ stores all the following values from each row:
'Ticker Symbol
'Volume of Stock
'Open Price (opening price at the beginning of a given year. First occurrence in <open>)
'Close Price (closing price at the end of that year. Last occurrence in <close>)

Sub RetrieveData()
    Dim ws As Worksheet
    Dim TickerDict As Object
    Dim cell As Range
    Dim lastRow As Long
    Dim i As Long
    Dim total As Double
    Dim tickerValue As Variant
    Dim lastOccurrence As Long
    Dim tickerLastRow As Object
    
    'Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets(Array("2018", "2019", "2020"))
        Set TickerDict = CreateObject("Scripting.Dictionary")
        i = 2 'Start writing from row 2 dude to headers
        
        'Find the last used row in Column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        'Find the Unique Tickers and Total Stock Value
        For Each cell In ws.Range("A2:A" & lastRow)
            If cell.Value <> "" And Not TickerDict.exists(cell.Value) Then
                TickerDict.Add cell.Value, Nothing
                
                'Find the total sum for the current ticker
                total = Application.WorksheetFunction.SumIf(ws.Range("A:A"), cell.Value, ws.Range("G:G"))
                
                'Find the last occurrence of the current ticker in column A
                Set tickerLastRow = ws.Range("A:A").Find(What:=cell.Value, After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious)
                If Not tickerLastRow Is Nothing Then
                    lastOccurrence = tickerLastRow.Row
                    'Paste the total sum in column L =Total Stock Value
                    ws.Cells(i, "L").Value = total
                    'Paste the ticker in column I = Unique Tickers (no duplicates)
                    ws.Cells(i, "I").Value = cell.Value
                    'Find the corresponding value in column C = Open
                    tickerValue = ws.Cells(cell.Row, "C").Value
                    'Paste the corresponding value in column S = Open Price
                    ws.Cells(i, "S").Value = tickerValue
                    'Find the corresponding value in column F = Close and Pate in T = Close Price
                    ws.Range("T" & i).Value = ws.Cells(lastOccurrence, "F").Value
                    i = i + 1
                End If
            End If
        Next cell
    Next ws
End Sub

'Calculate Yearly Change:
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'Yearly Change = Closing Price - Opening Price (J =T-S)

Sub CalculateYearlyChange()
    Dim SheetsArr As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    
    'Array of sheets for headers
    SheetsArr = Array("2018", "2019", "2020")
    
    'Loop through each sheet in the array
    For Each sheetName In SheetsArr
        'Set the worksheet
        Set ws = ThisWorkbook.Sheets(sheetName)
        
        'Find the last row in column S (Note: need the last row in column T or S to ensure calculations are not extending beyond the actual data.)
        lastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
        
        'Calculate Yearly Change
        ws.Range("J2:J" & lastRow).Formula = "=T2-S2"
        
        'Convert formulas to values
        ws.Range("J2:J" & lastRow).Value = ws.Range("J2:J" & lastRow).Value
    Next sheetName
End Sub

'Calculate Percent Change:
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'Percent Change = ((Yearly change)/Opening Price)(K = J/S)

Sub CalculatePercentChange()
    Dim lastRow As Long
    Dim i As Long
    Dim ws As Worksheet
    Dim SheetsArr As Variant
    
    'Array of sheets for headers
    SheetsArr = Array("2018", "2019", "2020")
    
    'Loop through each sheet in the array
    For Each wsName In SheetsArr
        'Set the worksheet
        Set ws = ThisWorkbook.Sheets(wsName)
        
        'Find the last row of data in column J or S
        lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        
        'Loop through each row to calculate the percent change and paste the result in column K
        For i = 2 To lastRow
            If IsNumeric(ws.Cells(i, "J").Value) And IsNumeric(ws.Cells(i, "S").Value) Then
                ws.Cells(i, "K").Value = (ws.Cells(i, "J").Value / ws.Cells(i, "S").Value)
                'Format the cell as percentage with two decimal places
                ws.Cells(i, "K").NumberFormat = "0.00%"
            End If
        Next i
    Next wsName
End Sub

'Apply Conditional Formatting For the Yearly Change:
'Conditional formatting that will highlight positive change in green and negative change in red

Sub ApplyConditionalFormatting()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets(Array("2018", "2019", "2020"))
        With ws.Range("J2:J" & ws.Cells(ws.Rows.Count, "J").End(xlUp).Row).FormatConditions
            .Delete
            ' Greater than 0 (Green)
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .Item(1).Interior.ColorIndex = 4
            
            ' Less than 0 (Red)
            .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .Item(2).Interior.ColorIndex = 3
        End With
    Next ws
End Sub

'Find the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume":

Sub FindMaxMinAndCorrespondingValues()
    Dim ws As Worksheet
    Dim maxK As Double, minK As Double, maxL As Double
    Dim maxI As Variant, minI As Variant, maxIL As Variant
    Dim maxIIndex As Long, minIIndex As Long, maxILIndex As Long
    
    'Loop through each worksheet in the array of sheets
    For Each ws In ThisWorkbook.Sheets(Array("2018", "2019", "2020"))
        'Find max and min values in column K
        maxK = WorksheetFunction.Max(ws.Columns("K"))
        minK = WorksheetFunction.Min(ws.Columns("K"))
        
        'Find max value in column L
        maxL = WorksheetFunction.Max(ws.Columns("L"))
        
        'Find the index of the maximum and minimum values in column K
        maxIIndex = Application.WorksheetFunction.Match(maxK, ws.Columns("K"), 0)
        minIIndex = Application.WorksheetFunction.Match(minK, ws.Columns("K"), 0)
        
        'Find the index of the maximum value in column L
        maxILIndex = Application.WorksheetFunction.Match(maxL, ws.Columns("L"), 0)
        
        'Retrieve the corresponding values in column I using the found indices
        maxI = ws.Cells(maxIIndex, "I").Value
        minI = ws.Cells(minIIndex, "I").Value
        maxIL = ws.Cells(maxILIndex, "I").Value

        'Format cells in column Q as percentage with two decimal places
        ws.Range("Q2:Q3").NumberFormat = "0.00%"

        
        'Paste values in specified cells
        ws.Range("P2").Value = maxI
        ws.Range("Q2").Value = maxK
        ws.Range("P3").Value = minI
        ws.Range("Q3").Value = minK
        ws.Range("P4").Value = maxIL
        ws.Range("Q4").Value = maxL
    Next ws
End Sub
