Attribute VB_Name = "Module3"
Sub ChallengeTest()
        
'create Challenge Summary Output table
'Declare variables for summary values

    Dim ws As Worksheet
    Dim Greatest_Percent_Increase_Ticker As String
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease_Ticker As String
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Total_Volume_Ticker As String
    Dim Greatest_Total_Volume As Long
    Dim lastrow_summary_table As Integer
 
' Create Challenge summary titles

For Each ws In Worksheets
    
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
    'determine the last row of the summary table output
     lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row

    'Determine the maximum and minimum values in column "Percent Change" and
    'the maximum value in column "Total Stock Volume"
    'and assign to the corresponding ticker name and value cells in the summary output
    
      For i = 2 To lastrow_summary_table
        
            'Determine the maximum percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Determin the minimum percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Determine the maximum total stock volume
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        


End Sub
