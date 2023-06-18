VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()

    'Define variables needed for the script to function
    Dim ws As Worksheet
    Dim ticker As String
    Dim vol As Double
    Dim lastrow As Long
    Dim summary_table_row As Integer
    Dim start As Long
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim formatRange As Range
    Dim maxPercentageChange As Double
    Dim maxPercentageIncreaseTicker As String
    Dim minPercentageChange As Double
    Dim minPercentageDecreaseTicker As String
    Dim TotalvolumeTicker As String
           
    'Variable for greatest % increse must be set outside of loop
     maxPercentageIncreaseTicker = ""
     minPercentageDecreaseTicker = ""
     maxTotalVolume = 0
     maxTotalVolumeTicker = 0
     

    'Add For each loop to gather data from all sheets
     For Each ws In ThisWorkbook.Sheets
    
    'Set locataion where the summary table will start
    summary_table_row = 2
    start = 2
        
    'Define the last row
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    'Insert column headers for the summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    

    
    'Add loop to look in each row for each sheet
    For i = 2 To lastrow
    
    'Check to see if the Ticker is still same or if it has changed
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Set ticker name
        ticker = ws.Cells(i, 1).Value
        
        
    'Calculate Yearly changes using the first open and the last close for each ticker
     yearlychange = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value
      
    'Calculate percentage change from the opening price at the beginning of a given year to the closing price at the end of that year
     percentchange = (yearlychange / ws.Cells(start, 3).Value)
     start = i + 1
     If percentchange > maxPercentageChange Then
       maxPercentageChange = percentchange
       maxPercentageChangeRow = summary_table_row
       maxPercentageIncreaseTicker = ticker
        End If
     
     If percentchange < minPercentageChange Then
      minPercentageChange = percentchange
      minPercentageChangeRow = summary_table_row
      minPercentageDecreaseTicker = ticker
    End If
  
                
    'Store volume data (Vol) to be printed later
    vol = vol + ws.Cells(i, 7).Value
    
    
     'Print total volume by ticker in summary table
    ws.Range("L" & summary_table_row).Value = vol
    If vol > maxTotalVolume Then
        maxTotalVolume = vol
        maxTotalVolumeTicker = ticker
    End If
    
    
   'Print the ticker name in the summary table on Sheet 2018
    ws.Range("I" & summary_table_row).Value = ticker
    
    'Print the Yearly Change in the summary table
    ws.Range("J" & summary_table_row).Value = yearlychange
            
    'Print percentage change in summary table
    ws.Range("K" & summary_table_row).Value = percentchange
    ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
    
    'Conditional formating of cells
    Set formatRange = ws.Range("J" & summary_table_row)
                With formatRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
                    .Interior.Color = RGB(0, 255, 0)
                    End With
    Set formatRange = ws.Range("J" & summary_table_row)
                With formatRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="0")
                    .Interior.Color = RGB(255, 0, 0)
                    End With

        
    'Add one count to the summary table
    summary_table_row = summary_table_row + 1
        
    'Reset counts for Volume, and yearlychange
    vol = 0
    yearlychange = 0
    percentchange = 0
 
            
  End If
  
       vol = vol + ws.Cells(i, 7).Value
  
 Next i
 
    ' Print highest % increased and highest % decreased
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("P2").Value = maxPercentageIncreaseTicker
    ws.Range("P3").Value = minPercentageDecreaseTicker
    ws.Range("O2").Value = "Greatest % Increased"
    ws.Range("O3").Value = "Greatest % Decreased"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P4").Value = maxTotalVolumeTicker
    ws.Range("Q2").Value = maxPercentageChange
    ws.Range("Q3").Value = minPercentageChange
    ws.Range("Q4").Value = maxTotalVolume
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
 
Next ws


    
End Sub
