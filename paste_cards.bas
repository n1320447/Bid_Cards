Sub Copy_Paste_Cards_to_Sheets()


'LOOP FOR CYCLING THROUGH SHEET NAMES
    Sheets("SHEET CREATOR").Select
    Dim x As Integer
    Application.ScreenUpdating = False
    ' Set numrows = number of rows of data.
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    'Debug.Print NumRows
    ' Select cell a1.
    'Range("A1").Select
    ' Establish "For" loop to loop "numrows" number of times.
    'var = ""
    Dim arr() As String
    For x = 1 To NumRows
        Sheets("SHEET CREATOR").Select
        y = "A" + CStr(x)
        Range(y).Select
        ' Insert your code here.
        ' Selects cell down 1 row from active cell.
        ' ActiveCell.Offset(1, 0).Select
        Dim var As Variant
        var = Range(y).Value

        ReDim Preserve arr(x)
        arr(x) = var

        

    Next
        
'START OF RANGE LOOP
    Sheets("CARD DUMP").Select
    Dim SrchRng As Range, cel As Range
    Set SrchRng = Range("A1:S100000")
    Start = 0
    end1 = 0
    wallet = 0
    Start2 = 0
    end2 = 0
    end3 = 0
    Banana = 0
    phone = 0
    monkey = 0
    Dim arr1() As String
    Dim arr2() As String
    Dim arr3() As String
    Dim arr4() As String
    For Each cel In SrchRng
        If InStr(1, cel.Value, "DESCRIPTION") > 0 Then

            Start = "A" + CStr(cel.Row)
            'Start2 = Trim(Replace(Start, "A", ""))
            Start2 = CStr(cel.Row + 10)
            ReDim Preserve arr1(wallet * 2)
            arr1(wallet * 2) = Start
            
            ReDim Preserve arr4(monkey + 10)
            arr4(monkey + 10) = Start2
            monkey = monkey + 1

        End If

        If InStr(1, cel.Value, "CARD TOTAL MC2") > 0 Then


            end1 = "T" + CStr(cel.Row + 1)
            end2 = Trim(Replace(end1, "T", ""))
            end2 = CStr(cel.Row)
            end3 = CStr(cel.Row + 1)

            ReDim Preserve arr1((wallet * 2) + 1)
            arr1((wallet * 2) + 1) = end1
            wallet = wallet + 1
       
            ReDim Preserve arr2(Banana + 1)
            arr2(Banana + 1) = end2
            Banana = Banana + 1
            
            ReDim Preserve arr3(phone + 1)
            arr3(phone + 1) = end3
            phone = phone + 1

        End If
      
    
    Next cel
    Dim total As Integer
    For n = 0 To (NumRows - 1)
        Sheets("CARD DUMP").Select
        r = CStr(arr3(n + 1)) 'Cell to put total from sum formula
        r1 = CStr(arr2(n + 1)) 'Lower bound of range to sum
        r2 = CStr(arr4(n + 10)) 'Upper bound of range to sum
        Range("M" + r).Value = "=Sum(M" + r1 + ":M" + r2 + ")"
        Range("N" + r).Value = "=Sum(N" + r1 + ":N" + r2 + ")"
        Range("O" + r).Value = "=Sum(O" + r1 + ":O" + r2 + ")"
        Range("P" + r).Value = "=Sum(P" + r1 + ":P" + r2 + ")"
        Range("Q" + r).Value = "=Sum(Q" + r1 + ":Q" + r2 + ")"
        Range("R" + r).Value = "=Sum(R" + r1 + ":R" + r2 + ")"
        Range("S" + r).Value = "=Sum(S" + r1 + ":S" + r2 + ")"
        Range("T" + r).Value = "=Sum(T" + r1 + ":T" + r2 + ")"
        
    Next
    For j = 0 To (NumRows - 1)
        Sheets("CARD DUMP").Select
        Range(arr1(j * 2), arr1((j * 2) + 1)).Copy
        Sheets(arr(j + 1)).Select
        'Call AddOutsideBorders(ActiveWorkbook.Worksheets(arr(j + 1)).Range("A3:S10"))
        Range("A1").Select
        Range("M:T").ColumnWidth = 14
        ActiveSheet.Paste
        ActiveSheet.Range("A3:T10").BorderAround xlContinuous, xlThick      
        Application.ScreenUpdating = True
    Next
    

Application.Goto Reference:=Sheets("SHEET CREATOR").Range("A1")
End Sub
