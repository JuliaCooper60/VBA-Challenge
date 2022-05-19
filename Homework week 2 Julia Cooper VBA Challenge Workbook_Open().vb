Private Sub Workbook_Open()
     
End Sub

Private Sub RunProcess()
    Dim TickerSymbol As String
    Dim openStart As Double
    Dim closedEnd As Double
    Dim totalVolume As Single
    
    Dim greatestIncreaseTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecreaseTicker As String
    Dim greatestDecrease As Double
    Dim greatestVolumeTicker As String
    Dim greatestVolume As Double
    
    Dim countSummary  As Integer
    
    'CHALLENGE 2. Requirement: Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.
    
    'Loop through each sheet
    For Each Sheet In Worksheets
        Sheet.Activate
        'Select the first position in the sheet
        ActiveSheet.Range("A2").Activate
        
        'Reset all values between sheets
        TickerSymbol = ""
        openStart = 0
        closedEnd = 0
        totalVolume = 0
        greatestIncreaseTicker = ""
        greatestIncrease = 0
        greatestDecreaseTicker = ""
        greatestDecrease = 0
        greatestVolumeTicker = ""
        greatestVolume = 0
        countSummary = 2
        
        
        'Loop through all the cells in the first column until there are not values to process.
        Do While ActiveCell.Value <> ""
            If TickerSymbol <> ActiveCell.Value Then
                If TickerSymbol <> "" Then
                    ActiveSheet.Range("I" & countSummary) = TickerSymbol
                    ActiveSheet.Range("J" & countSummary) = (closedEnd - openStart)
                    If (closedEnd - openStart) < 0 Then
                        ActiveSheet.Range("J" & countSummary).Interior.Color = RGB(255, 0, 0)
                    Else
                        ActiveSheet.Range("J" & countSummary).Interior.Color = RGB(0, 255, 0)
                    End If
                    If openStart <> 0 Then
                        ActiveSheet.Range("K" & countSummary) = ((closedEnd - openStart) / openStart)
                    Else
                        ActiveSheet.Range("K" & countSummary) = 0
                    End If
                    ActiveSheet.Range("L" & countSummary) = totalVolume
                    If openStart <> 0 Then
                        If greatestIncrease < ((closedEnd - openStart) / openStart) Then
                            greatestIncreaseTicker = TickerSymbol
                            greatestIncrease = ((closedEnd - openStart) / openStart)
                        End If
                        If greatestDecrease > ((closedEnd - openStart) / openStart) Then
                            greatestDecreaseTicker = TickerSymbol
                            greatestDecrease = ((closedEnd - openStart) / openStart)
                        End If
                    Else
                        If greatestIncrease < 0 Then
                            greatestIncreaseTicker = TickerSymbol
                            greatestIncrease = 0
                        End If
                        If greatestDecrease > 0 Then
                            greatestDecreaseTicker = TickerSymbol
                            greatestDecrease = 0
                        End If
                    End If
                    If greatestVolume < totalVolume Then
                        greatestVolumeTicker = TickerSymbol
                        greatestVolume = totalVolume
                    End If
                    
                    countSummary = countSummary + 1
                End If
                openStart = ActiveCell.Offset(0, 2).Value
                totalVolume = 0
                TickerSymbol = ActiveCell.Value
            End If
            closedEnd = ActiveCell.Offset(0, 5).Value
            totalVolume = totalVolume + ActiveCell.Offset(0, 6).Value
            ActiveCell.Offset(1, 0).Activate
        Loop
        'Write out our overall values
        ActiveSheet.Range("P2") = greatestIncreaseTicker
        ActiveSheet.Range("Q2") = greatestIncrease
        ActiveSheet.Range("P3") = greatestDecreaseTicker
        ActiveSheet.Range("Q3") = greatestDecrease
        ActiveSheet.Range("P4") = greatestVolumeTicker
        ActiveSheet.Range("Q4") = greatestVolume
    Next
    Worksheets(1).Activate
    ActiveSheet.Range("A2").Activate
End Sub
