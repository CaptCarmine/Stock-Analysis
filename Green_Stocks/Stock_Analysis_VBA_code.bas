Attribute VB_Name = "Module1"
Sub MacroCheck()

Dim Hotstuff As String

Hotstuff = "This is a box"

MsgBox (Hotstuff)


End Sub
Sub ClearWorksheet()

Cells.Clear

End Sub
Sub nested_practice()

Worksheets("Nested_Test").Activate

For i = 1 To 8
  For j = 1 To 8
    If ((i Mod 2 = 0) And (j Mod 2 = 0)) Then
    Cells(i, j).Interior.Color = vbGreen
    
    ElseIf ((i Mod 2 = 0) Or (j Mod 2 = 0)) Then
    Cells(i, j).Interior.Color = vbRed
    
    Else
    Cells(i, j).Interior.Color = vbGreen
    
    End If
  Next j
Next i
 
End Sub
Sub yearValueAnalysis()

yearValue = InputBox("What year would you like to run the analysis on?")

'3a) Initialize variables for starting price and ending price
Dim rowStart As Integer
Dim rowEnd As Integer
Dim TotalVolume As Long
Dim startingPrice As Double
Dim endingPrice As Double
Dim startTime As Single
Dim endTime  As Single

startTime = Timer

'1) Format the output sheet on All Stocks Analysis worksheet
Worksheets("All_Stocks_Analysis").Activate

  Cells(1, 1).Value = "All Stocks (" + yearValue + ")"
  'Create row headers
  Cells(3, 1).Value = "Ticker"
  Cells(3, 2).Value = "Total Daily Volume"
  Cells(3, 3).Value = "Return"


'2) Initialize array of all tickers
Dim tickers(12) As String
  tickers(0) = "AY"
  tickers(1) = "CSIQ"
  tickers(2) = "DQ"
  tickers(3) = "ENPH"
  tickers(4) = "FSLR"
  tickers(5) = "HASI"
  tickers(6) = "JKS"
  tickers(7) = "RUN"
  tickers(8) = "SEDG"
  tickers(9) = "SPWR"
  tickers(10) = "TERP"
  tickers(11) = "VSLR"

' 3b) Go to Data worksheet Worksheet
  Worksheets("2018").Activate

  rowStart = 2
  'rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
  rowEnd = Worksheets(yearValue).Range("A1").End(xlDown).Row
  'Set initial volume to zero
  Worksheets(yearValue).Activate

' 4) Nested loop to create all Ticker information
For t = 0 To 11
  ticker = tickers(t)
  'Set initial volume to zero
  TotalVolume = 0
  Worksheets(yearValue).Activate
    For i = rowStart To rowEnd
      'increase totalVolume
    '5a) Find total volume for current ticker
    If Cells(i, 1).Value = ticker Then
        TotalVolume = TotalVolume + Cells(i, 8).Value
        
    End If
    
    'Conditional to find startingprice Value
    If Cells(i, 1).Value = ticker And Cells(i - 1, 1).Value <> ticker Then
      startingPrice = Cells(i, 6).Value
    
    End If
    
    'Conditionnal to find startingprice Value
    If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
      endingPrice = Cells(i, 6).Value
    
    End If
    
    Next i
    
  Worksheets("All_Stocks_Analysis").Activate
  Cells(4 + t, 1).Value = ticker
  Cells(4 + t, 2).Value = TotalVolume
  Cells(4 + t, 3).Value = endingPrice / startingPrice - 1
    
  Next t

endTime = Timer
MsgBox ("This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue))

'reference code
'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists

End Sub
Sub Formatting_All_Stocks_AnalysisTable()

'formatting Price 2018
Worksheets("2018").Activate
Range("C:G").NumberFormat = "$0.00"

'Formatting All Stocks
Worksheets("All_Stocks_Analysis").Activate
Range("A3:C3").Font.Bold = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous

Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.00%"
Columns("B").AutoFit

dataRowStart = 4
dataRowEnd = Worksheets("All_Stocks_Analysis").Range("C4").End(xlDown).Row
'MsgBox (dataRowEnd)

For i = dataRowStart To dataRowEnd
  If Cells(i, 3) > 0 Then
    'Color the cell green
    Cells(i, 3).Interior.Color = vbGreen
  
  ElseIf Cells(i, 3) < 0 Then
    'Color the cell Red
    Cells(i, 3).Interior.Color = vbRed
    
  Else
    'Cell has no color
    Cells(i, 3).Interior.Color = xlNone
    
  End If
  
Next i

End Sub
Sub All_Stocks_Analysis()
'1) Format the output sheet on All Stocks Analysis worksheet
Worksheets("All_Stocks_Analysis").Activate

  Cells(1, 1).Value = "All Stocks (2018)"
  'Create row headers
  Cells(3, 1).Value = "Ticker"
  Cells(3, 2).Value = "Total Daily Volume"
  Cells(3, 3).Value = "Return"

'3a) Initialize variables for starting price and ending price
Dim rowStart As Integer
Dim rowEnd As Integer
Dim TotalVolume As Long
Dim startingPrice As Double
Dim endingPrice As Double

'2) Initialize array of all tickers
Dim tickers(12) As String
  tickers(0) = "AY"
  tickers(1) = "CSIQ"
  tickers(2) = "DQ"
  tickers(3) = "ENPH"
  tickers(4) = "FSLR"
  tickers(5) = "HASI"
  tickers(6) = "JKS"
  tickers(7) = "RUN"
  tickers(8) = "SEDG"
  tickers(9) = "SPWR"
  tickers(10) = "TERP"
  tickers(11) = "VSLR"

' 3b) Go to Data worksheet Worksheet
  Worksheets("2018").Activate

  rowStart = 2
  'rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
  rowEnd = Worksheets("2018").Range("A1").End(xlDown).Row
  'Set initial volume to zero
  Worksheets("2018").Activate

' 4) Nested loop to create all Ticker information
For t = 0 To 11
  ticker = tickers(t)
  'Set initial volume to zero
  TotalVolume = 0
  Worksheets("2018").Activate
    For i = rowStart To rowEnd
      'increase totalVolume
    '5a) Find total volume for current ticker
    If Cells(i, 1).Value = ticker Then
        TotalVolume = TotalVolume + Cells(i, 8).Value
        
    End If
    
    'Conditional to find startingprice Value
    If Cells(i, 1).Value = ticker And Cells(i - 1, 1).Value <> ticker Then
      startingPrice = Cells(i, 6).Value
    
    End If
    
    'Conditionnal to find startingprice Value
    If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
      endingPrice = Cells(i, 6).Value
    
    End If
    
    Next i
    
  Worksheets("All_Stocks_Analysis").Activate
  Cells(4 + t, 1).Value = ticker
  Cells(4 + t, 2).Value = TotalVolume
  Cells(4 + t, 3).Value = endingPrice / startingPrice - 1
    
  Next t

'reference code
'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists

End Sub
Sub DQAnalysis()

Dim rowStart As Integer
Dim rowEnd As Integer
Dim TotalVolume As Long
Dim startingPrice As Double
Dim endingPrice As Double

'Go to 2018 Worksheet
Worksheets("2018").Activate

  'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists


  rowStart = 2
  'rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
  rowEnd = Worksheets("2018").Range("A1").End(xlDown).Row
  'Set initial volume to zero
  TotalVolume = 0

For i = rowStart To rowEnd
    'increase totalVolume
    
  If Cells(i, 1).Value = "DQ" Then
      TotalVolume = TotalVolume + Cells(i, 8).Value
      
  End If
  
  'Conditional to find startingprice Value
  If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
    startingPrice = Cells(i, 6).Value

  End If
  
  'Conditionnal to find startingprice Value
  If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
    endingPrice = Cells(i, 6).Value

  End If
Next i

'Debug Commands to verify Values found
'MsgBox ("Total Volume: " + Str(TotalVolume) & vbNewLine & "Starting Price: " + Str(startingPrice) & vbNewLine & "Ending Price: " + Str(endingPrice)), vbInformation


'Go to DQ Analysis Workheet
 Worksheets("DQ_Analysis").Activate
 
'Add heder Files
 Range("A1").Value = "DAQO (Ticker: DQ)"
 
  Cells(3, 1).Value = "Year"
  Cells(3, 2).Value = "Total Daily Volume"
  Cells(3, 3).Value = "Return"
 
  Cells(4, 1).Value = 2018
  Cells(4, 2).Value = TotalVolume
  Cells(4, 3).Value = endingPrice / startingPrice - 1

End Sub
Sub DQAnalysis_Range()

'Go to DQ Analysis Workheet
 Worksheets("DQ_Analysis").Activate
 
'Add heder Files
 Range("A1").Value = "DAQO (Ticker: DQ)"
 
 Range("A3").Value = "Year"
 Range("B3").Value = "Total Daily Volume"
 Range("C3").Value = "Return"

End Sub

Sub DQAnalysis_Cells()

'Go to DQ Analysis Workheet
 Worksheets("DQ_Analysis").Activate
 
'Add heder Files
 
 Cells(1, 1).Value = "DAQO (Ticker: DQ)"
 
 Cells(3, 1).Value = "Year"
 Cells(3, 2).Value = "Total Daily Volume"
 Cells(3, 3).Value = "Return"

End Sub
Sub Clear()

'Clear Testing range
 Range("A1:J4").Value = ""
 Range(Cells(1, 1), Cells(10, 4)) = ""

End Sub
