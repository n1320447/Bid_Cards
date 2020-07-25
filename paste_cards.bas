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
        'Debug.Print var
        'Debug.Print x
        ReDim Preserve arr(x)
        arr(x) = var
        'range start
        'Partial_Text = var
        
'        'START OF RANGE LOOP
'        Sheets("CARD DUMP").Select
'        Dim SrchRng As Range, cel As Range
'        Set SrchRng = Range("A1:S100000")
'        start = 0
'        end1 = 0
'        For Each cel In SrchRng
'            If InStr(1, cel.Value, "DESCRIPTION") > 0 Then
'
'                'Debug.Print cel
'                start = "A" + CStr(cel.Row)
'                Debug.Print start
'
'            End If
'
'            If InStr(1, cel.Value, "CARD TOTAL MC2") > 0 Then
'
'                'Debug.Print cel
'                end1 = "S" + CStr(cel.Row + 1)
'                Debug.Print end1
'
'            End If
'
'        Next cel
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

            'Debug.Print cel
            Start = "A" + CStr(cel.Row)
            'Start2 = Trim(Replace(Start, "A", ""))
            Start2 = CStr(cel.Row + 10)
            'Debug.Print Start2
            ReDim Preserve arr1(wallet * 2)
            arr1(wallet * 2) = Start
            
            ReDim Preserve arr4(monkey + 10)
            arr4(monkey + 10) = Start2
            monkey = monkey + 1
            'monkey = monkey + 1
            'Debug.Print arr1(wallet * 2)
            'Debug.Print Start2
            'Debug.Print arr4(monkey + 10)
            'comment
        End If

        If InStr(1, cel.Value, "CARD TOTAL MC2") > 0 Then

            'Debug.Print cel
            end1 = "T" + CStr(cel.Row + 1)
            end2 = Trim(Replace(end1, "T", ""))
            end2 = CStr(cel.Row)
            end3 = CStr(cel.Row + 1)
            'Debug.Print CStr(cel.Row + 1)
            'Debug.Print end1
            'Debug.Print wallet
            'Debug.Print end3
            ReDim Preserve arr1((wallet * 2) + 1)
            arr1((wallet * 2) + 1) = end1
            wallet = wallet + 1
            'ReDim Preserve arr1(
            'Debug.Print end2
            
            ReDim Preserve arr2(Banana + 1)
            arr2(Banana + 1) = end2
            Banana = Banana + 1
            
            ReDim Preserve arr3(phone + 1)
            arr3(phone + 1) = end3
            phone = phone + 1
            'Debug.Print arr2(Banana)
            'Debug.Print arr3(phone)
        End If
      
    
    Next cel
    'Debug.Print "this array 3 " + arr3(phone)
    Dim total As Integer
    For n = 0 To (NumRows - 1)
        Sheets("CARD DUMP").Select
        r = CStr(arr3(n + 1)) 'Cell to put total from sum formula
        'Debug.Print r
        r1 = CStr(arr2(n + 1)) 'Lower bound of range to sum
        'Debug.Print r1
        r2 = CStr(arr4(n + 10)) 'Upper bound of range to sum
        'Debug.Print r2
        'Debug.Print arr3(n)
        'Debug.Print arr2(n + 1)
        'Debug.Print arr4(n + 10)
        'Debug.Print CStr(arr3(n))
        'Debug.Print ("M" & r1 & ":M" & r2)
        'Debug.Print (("M" & r1) & ":" & ("M" & r2))
        'Debug.Print ActiveSheet.Range("M" & r)
        'Range("M" & r).Formula = "=Sum((M & r1):(M & r2))"
        Range("M" + r).Value = "=Sum(M" + r1 + ":M" + r2 + ")"
        'Debug.Print Report.Cells("M" & r).Value = Excel.WorksheetFunction.Sum(Report.Range("(M & r1):(M & r2)"))
        'Debug.Print Report.Cells(13, r)
        'Range("M" & r) = WorksheetFunction.Sum(Worksheets("CARD DUMP").Range(("M" & r1), ("M" & r2)))
       
       
        
    Next
    'Debug.Print Range("M79").Formula = "SUM(M78:M11)"
    For j = 0 To (NumRows - 1)
        Sheets("CARD DUMP").Select
        'Debug.Print arr(j + 1)
        'Debug.Print arr1(j * 2)
        'Debug.Print arr1((j * 2) + 1)
        'Range ((M & CStr(cel.Row + 1))) = WorksheetFunction.Sum(Worksheets(arr(j+1)).Range((M & CStr(cel.Row)),(M))
        'Range(
        'Debug.Print arr3
        Range(arr1(j * 2), arr1((j * 2) + 1)).Copy
        'Range("M":(end2(CStr(cel.Row))))
        Sheets(arr(j + 1)).Select
        'Call AddOutsideBorders(ActiveWorkbook.Worksheets(arr(j + 1)).Range("A3:S10"))
        Range("A1").Select
        Range("M:T").ColumnWidth = 14
        ActiveSheet.Paste
        ActiveSheet.Range("A3:T10").BorderAround xlContinuous, xlThick
        'Debug.Print Start2
        'ActiveSheet.Range("A3:S
        'lRow = ThisWorkbook.Sheets(CurrentSheet).Range("A10000000000").End(xlUp).Row
        'Debug.Print lRow
        
        Application.ScreenUpdating = True
    Next
    
   'comment2

Application.Goto Reference:=Sheets("SHEET CREATOR").Range("A1")
End Sub
