VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub TickerPicker()

' Volume Total (vtotal).
' Index Row (inrow).
' Index Column (incol).
' Actual Change (valchange).
' Percentage Change (pchange).
' Index Column (incol).
' Start (start).
' Row Count (rowcount).
' Days (days).
' Daily Change (daychange).
' Average change (avchange).
' Worksheet (ws).

Dim vtotal As Double
Dim inrow As Long
Dim incol As Integer
Dim valchange As Double
Dim pchange As Double
Dim start As Long
Dim rowcount As Long
Dim days As Integer
Dim daychange As Single
Dim avchange As Double
Dim ws As Worksheet

' Looping through each worksheet.
For Each ws In Worksheets

    inrow = 2
    incol = 0
    vtotal = 0
    valchange = 0
    start = 2
    daychange = 0

' Setting Headers
    ws.Range("I1").value = "Ticker"
    ws.Range("J1").value = "Yearly Change ($)"
    ws.Range("K1").value = "Yearly Change (%)"
    ws.Range("L1").value = "Total Stock Volume"
    ws.Range("O1").value = "Ticker"
    ws.Range("P1").value = "Value"
    ws.Range("N2").value = "Greatest % Increase"
    ws.Range("N3").value = "Greatest % Decrease"
    ws.Range("N4").value = "Greatest Total Volume"

' Retrieving row # of the last row with data.
    rowcount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
           
    For inrow = 2 To rowcount
        
        ' Checking if the Ticker name is the same.
        If ws.Cells(inrow + 1, 1).value <> ws.Cells(inrow, 1).value Then
                 
            ' Store results in a variable.
            vtotal = vtotal + ws.Cells(inrow, 7).value
        
            If vtotal = 0 Then
                'Print the results
                ws.Range("I" & 2 + incol).value = ws.Cells(inrow, 1).value
                ws.Range("J" & 2 + incol).value = 0
                ws.Range("K" & 2 + incol).value = "%" & 0
                ws.Range("L" & 2 + incol).value = 0
            Else
                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To inrow
                        If ws.Cells(find_value, 3).value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
                            
                valchange = (ws.Cells(inrow, 6) - ws.Cells(start, 3))
                pchange = valchange / ws.Cells(start, 3)
                
                start = inrow + 1
                
                ' Print Values and format them according to specifications.
                ws.Range("I" & 2 + incol) = ws.Cells(inrow, 1).value
                ws.Range("J" & 2 + incol) = valchange
                ws.Range("J" & 2 + incol).NumberFormat = "0.00"
                ws.Range("K" & 2 + incol).value = pchange
                ws.Range("K" & 2 + incol).NumberFormat = "0.00%"
                ws.Range("L" & 2 + incol).value = vtotal
                
                'Set cell shading.
                Select Case valchange
                    Case Is > 0
                        ws.Range("J" & 2 + incol).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + incol).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + incol).Interior.ColorIndex = 0
                End Select
                                                     
            End If
            
            vtotal = 0
            valchange = 0
            incol = incol + 1
            days = 0
            daychange = 0
            
        Else
        ' If "Ticker" is still the same, add the results.
        vtotal = vtotal + ws.Cells(inrow, 7).value
                                
        End If
    
    Next inrow
    
    'Print the Min and Max Values in Column P.
    ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowcount)) * 100
    ws.Range("P3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowcount)) * 100
    ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & rowcount))
   
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowcount)), ws.Range("L2:L" & rowcount), 0)
    
    ' Print Ticker values in column O.
    ws.Range("O2") = ws.Cells(increase_number + 1, 9)
    ws.Range("O3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("O4") = ws.Cells(volume_number + 1, 9)
    
Next ws

End Sub
