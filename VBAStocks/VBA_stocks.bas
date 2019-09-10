Attribute VB_Name = "Module1"
Sub HWtest1():

'Declaring the variables
Dim Total_Vol As Double
Dim inital_open As Double
Dim close_val As Double
Dim lastrow As Double
Dim Ticker_name As String
Dim largest_vol As Double
Dim Counter As Integer
Dim worksheetMaxVal As Double



Counter = 0
largest_vol = 0
Greatest_increase = 0
Greatest_decrease = 0

'looping through every worksheets
For Each ws In Worksheets

'finding the last row number and storing it in a variable
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Initializing variables and allocating value for initial_open
Total_Vol = 0
Final_Table = 2
initial_open = ws.Cells(2, 3).Value

WorksheetName = ws.Name
MsgBox WorksheetName

'Printing titles for the results
ws.Cells(1, 11).Value = "Ticker_name"
ws.Cells(1, 12).Value = "Total_Volume"
ws.Cells(1, 13).Value = "Yearly_Change"
ws.Cells(1, 14).Value = "Percent_Change"


'looping through every row
For i = 2 To lastrow
    
   'If the value in row 3 is not equal to row 2 and so on
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    
    'Store the value of i,1 as Ticker_name and add value of cells i,7 to Total_Vol
    Ticker_name = ws.Cells(i, 1).Value
    Total_Vol = Total_Vol + ws.Cells(i, 7).Value
    
    'Store the value of i,6 as close_val
    close_val = ws.Cells(i, 6).Value
    
    'Printing the values for every Ticker_name, their Total_Vol and  calculating Yearly_change
    ws.Cells(Final_Table, 11).Value = Ticker_name
    ws.Cells(Final_Table, 12).Value = Total_Vol
    ws.Cells(Final_Table, 13).Value = close_val - initial_open
    
    'Calculating Percent_change
    
    If initial_open = 0 Then
    ws.Cells(Final_Table, 14).Value = 0
    Else
    ws.Cells(Final_Table, 14).Value = ws.Cells(Final_Table, 13).Value / initial_open
    End If
    
    'Formatting Percent_change
    ws.Cells(Final_Table, 14).NumberFormat = "0.00%"
    
    'Doing a condtional color formatting for postive(green) and negative values(red)
    If ws.Cells(Final_Table, 14).Value > 0 Then
    ws.Cells(Final_Table, 13).Interior.ColorIndex = 4
    Else
    ws.Cells(Final_Table, 13).Interior.ColorIndex = 3
    End If
    
    
    'Resetting Total_Vol, initla_open and Final_Table before the next ticker iteration
    Total_Vol = 0
    initial_open = ws.Cells(i + 1, 3).Value
    
    Final_Table = Final_Table + 1
    
    Else
    
    'if the ticker values are the same then keep adding Total_Vol
    Total_Vol = Total_Vol + ws.Cells(i, 7).Value
    
    'store the new close_val
    close_val = ws.Cells(i, 6).Value
    
    End If
    Next i
    
If Counter = 0 Then

    largest_vol = Application.WorksheetFunction.Max(ws.Columns("L"))
    Greatest_increase = Application.WorksheetFunction.Max(ws.Columns("N"))
    Greatest_decrease = Application.WorksheetFunction.Min(ws.Columns("N"))
       
Else
               
    worksheetMaxVal = Application.WorksheetFunction.Max(ws.Columns("L"))
    worksheetGreatest_increase = Application.WorksheetFunction.Max(ws.Columns("N"))
    worksheetGreatest_decrease = Application.WorksheetFunction.Min(ws.Columns("N"))
               
               If largest_vol < worksheetMaxVal Then
                 largest_vol = worksheetMaxVal
                 
               End If
               
               
               If Greatest_increase < worksheetGreatest_increase Then
                   Greatest_increase = worksheetGreatest_increase
                    
               End If
               
               If Greatest_decrease > worksheetGreatest_decrease Then
                   Greatest_decrease = worksheetGreatest_decrease
                    
               End If
               
End If
       
       Counter = Counter + 1
       Next ws
       
           Cells(2, 16).Value = "Greatest total volume"
           Cells(2, 17).Value = largest_vol
           Cells(3, 16).Value = "Greatest % increase"
           'Cells(3, 16).NumberFormat = "0.00%"
           Cells(3, 17).Value = Greatest_increase
           Cells(4, 16).Value = "Greatest % Decrease"
           Cells(4, 17).Value = Greatest_decrease
           'Cells(4, 16).NumberFormat = "0.00%"
         
        
         
    
End Sub







