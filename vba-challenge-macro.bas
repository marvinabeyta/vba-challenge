Attribute VB_Name = "Module1"
Sub hw2()

'Set initial varible for holding ticker symbols
Dim ticker As String
'Set initial varible to hold open stock price for each ticker
Dim op_price As Double
'Set initial varible to hold close stock price for each ticker
Dim cl_price As Double
'Set initial varible to hold the Percentage Change between open and close prices
Dim percentchg As Double
'Set initial varible to hold the Percentage Change between open and close prices
Dim yrlychg As Double
'Set initial varible to hold the total volume of of stick ticker
Dim totvol As Double
' set initil varilbe to save number of ticjer symbols
Dim num_tickers As Integer
num_tickers = 0


'Notes on what columun the data resides
'ticker columun A, open columun B, close columun F, volume column G

' Keep track of the location for each ticker in the summary table
Dim header_row As Integer
header_row = 1
Dim summary_table_row As Integer
summary_table_row = 2

' Flag varibile used to reset counters
openflag = 0

'Compute last row in worksheet
Dim ws As Worksheet
Dim lastrow As Long
Set ws = ActiveSheet
lastrow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row

'MsgBox (lastrow)

'Print Columun Headers
Range("I" & header_row).Value = "<TICKER>"
Range("J" & header_row).Value = "<OPEN>"
Range("K" & header_row).Value = "<CLOSE>"
Range("L" & header_row).Value = "<YEARLY CHG>"
Range("M" & header_row).Value = "<VOLUME>"
Range("N" & header_row).Value = "<% CHANGE>"

' Loop through all stock tickers
For i = 2 To lastrow
    'Sum total volume of each stock ticker
    totvol = totvol + Cells(i, 7)
    'Captue open stock price
    If openflag = 0 Then
        op_price = CDbl(Cells(i, 3).Value)
        openflag = 1
    End If
        
    'Check for change in ticker symbol
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        'Save ticker symbol
        ticker = Cells(i, 1).Value
        'save year end price
        cl_price = Cells(i, 6).Value
        'compute percent change
        percentchg = cl_price / op_price
        'compute yearly change
        yrlychg = op_price - cl_price
        '
        num_tickers = num_tickers + 1
        
        'Store data of this ticker
        Range("I" & summary_table_row).Value = ticker
        Range("J" & summary_table_row).Value = op_price
        Range("K" & summary_table_row).Value = cl_price
        Range("L" & summary_table_row).Value = yrlychg
        Range("M" & summary_table_row).Value = totvol
        Range("N" & summary_table_row).Value = percentchg
        
        ' Reset flags and volume storage
        summary_table_row = summary_table_row + 1
        tot_vol = 0
        openflag = 0
        
    End If
Next i

'create varibles for greatest computations
Dim g_in_ticker As String
Dim g_in_price As Double
Dim g_de_ticker As String
Dim g_de_price As Double
Dim g_totvol_ticker As String
Dim g_totvol As Double

header_row = 1
'Print columun headers for greatest computations
Range("O" & header_row + 1).Value = "<Greatest % Increase>"
Range("P" & header_row + 2).Value = "<Greatest % Decrease>"
Range("Q" & header_row + 3).Value = "<Greatest Total Volume>"
Range("P" & header_row).Value = "<TICKER>"
Range("Q" & header_row).Value = "<VALUE>"

'At start,values are in first row, they will update as vba loops through the rest
g_in_ticker = Cells(2, "I").Value
g_in_price = Cells(2, "L").Value
g_de_ticker = Cells(2, "I").Value
g_de_price = Cells(2, "L").Value
g_totvol_ticker = Cells(2, "I").Value
g_totvol = Cells(2, "M").Value

' loop through summary table to deternime greatest % increase, greatest % decrease,gretest total volume
For i = 2 To num_tickers

    'Check for greater increase value
    If Cells(i + 1, 12).Value > g_in_price Then
        g_in_price = Cells(i + 1, 12).Value
        g_in_ticker = Cells(i + 1, 9).Value
    End If
    
    'Check for greater decrease value
    If Cells(i + 1, 12).Value < g_de_price Then
        g_de_price = Cells(i + 1, 12).Value
        g_de_ticker = Cells(i + 1, 9).Value
    End If
    
    'Check for greater volume value
    If Cells(i + 1, 13).Value > g_totvol Then
        g_totvol = Cells(i + 1, 13).Value
        g_totvol_ticker = Cells(i + 1, 9).Value
    End If
Next i

'print computed values


End Sub
