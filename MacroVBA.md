        Sub transaction_count()
        Dim currentValue As Boolean
        currentValue = Sheets("Monthly").Range("E13").Value
        currentValue = Sheets("Daily").Range("E37").Value
        currentValue = Sheets("DOW").Range("E14").Value
        currentValue = Sheets("Time Distribution").Range("E12").Value   
        If currentValue = False Then
        Sheets("Monthly").Range("E13").Value = True
        Sheets("Daily").Range("E37").Value = True
        Sheets("DOW").Range("E14").Value = True
        Sheets("Time Distribution").Range("E12").Value = True
        Else
        Sheets("Monthly").Range("E13").Value = False
        Sheets("Daily").Range("E37").Value = False
        Sheets("DOW").Range("E14").Value = False
        Sheets("Time Distribution").Range("E12").Value = False
        End If
        End Sub

Sub quantity_sold()
    Dim currentValue As Boolean
    currentValue = Sheets("Monthly").Range("F13").Value
    currentValue = Sheets("Daily").Range("F37").Value
    currentValue = Sheets("DOW").Range("F14").Value
    currentValue = Sheets("Time Distribution").Range("F12").Value
If currentValue = False Then
        Sheets("Monthly").Range("F13").Value = True
        Sheets("Daily").Range("F37").Value = True
        Sheets("DOW").Range("F14").Value = True
        Sheets("Time Distribution").Range("F12").Value = True
Else
        Sheets("Monthly").Range("F13").Value = False
        Sheets("Daily").Range("F37").Value = False
        Sheets("DOW").Range("F14").Value = False
        Sheets("Time Distribution").Range("F12").Value = False
End If
End Sub

Sub revenue()
    Dim currentValue As Boolean
    currentValue = Sheets("Monthly").Range("G13").Value
    currentValue = Sheets("Daily").Range("G37").Value
    currentValue = Sheets("DOW").Range("G14").Value
    currentValue = Sheets("Time Distribution").Range("G12").Value
If currentValue = False Then
        Sheets("Monthly").Range("G13").Value = True
        Sheets("Daily").Range("G37").Value = True
        Sheets("DOW").Range("G14").Value = True
        Sheets("Time Distribution").Range("G12").Value = True
Else
        Sheets("Monthly").Range("G13").Value = False
        Sheets("Daily").Range("G37").Value = False
        Sheets("DOW").Range("G14").Value = False
        Sheets("Time Distribution").Range("G12").Value = False
End If
End Sub
