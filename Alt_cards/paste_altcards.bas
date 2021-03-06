Sub Copy_Pase_Alt_or_Phase_Cards()
'CODE FOR ALT/PHASE CARDS

'LOOP FOR CYCLING THROUGH SHEET NAMES
    Sheets("SHEET CREATOR").Select
    Dim x As Integer
    Application.ScreenUpdating = False
    ' Set numrows = number of rows of data.
    NumRows = Range("M1", Range("M1").End(xlDown)).Rows.Count
    'Debug.Print NumRows
    ' Select cell a1.
    'Range("A1").Select
    ' Establish "For" loop to loop "numrows" number of times.
    'var = ""
    Dim arr() As String
    For x = 1 To NumRows
        Sheets("SHEET CREATOR").Select
        y = "M" + CStr(x)
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
    Sheets("ALT PHASE CARD DUMP").Select
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
    mouse = 0
    Dim arr1() As String
    Dim arr2() As String
    Dim arr3() As String
    Dim arr4() As String
    Dim arr5() As String
    Dim rng2 As Range
        Set rng2 = Range("A12:A100000") ' Identify your range
        n = 0
    For Each cel In SrchRng
        If InStr(1, cel.Value, "DESCRIPTION") > 0 Then

            Start = "A" + CStr(cel.Row)
            'Start2 = Trim(Replace(Start, "A", ""))
            Start2 = CStr(cel.Row + 10)
            Start3 = CStr(cel.Row + 2)
            ReDim Preserve arr1(wallet * 2)
            arr1(wallet * 2) = Start

            ReDim Preserve arr4(monkey + 10)
            arr4(monkey + 10) = Start2
            monkey = monkey + 1

            ReDim Preserve arr5(mouse + 2) 'Gives top left corner for thick vertical border
            arr5(mouse + 2) = Start3
            mouse = mouse + 1


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
'    Dim total As Integer
'    For n = 0 To (NumRows - 1)
'        Sheets("ALT PHASE CARD DUMP").Select
'        r = CStr(arr3(n + 1)) 'Cell to put total from sum formula
'        r1 = CStr(arr2(n + 1)) 'Lower bound of range to sum
'        r2 = CStr(arr4(n + 10)) 'Upper bound of range to sum
'        Range("M" + r).Value = "=Sum(M" + r1 + ":M" + r2 + ")"
'        Range("N" + r).Value = "=Sum(N" + r1 + ":N" + r2 + ")"
'        Range("O" + r).Value = "=Sum(O" + r1 + ":O" + r2 + ")"
'        Range("P" + r).Value = "=Sum(P" + r1 + ":P" + r2 + ")"
'        Range("Q" + r).Value = "=Sum(Q" + r1 + ":Q" + r2 + ")"
'        Range("R" + r).Value = "=Sum(R" + r1 + ":R" + r2 + ")"
'        Range("S" + r).Value = "=Sum(S" + r1 + ":S" + r2 + ")"
'        Range("T" + r).Value = "=Sum(T" + r1 + ":T" + r2 + ")"
'
'    Next

    For T = 0 To (NumRows - 1)
        Sheets("ALT PHASE CARD DUMP").Select
        r5 = CStr(arr5(T + 2)) 'Gives A3 on each card
        r = CStr(arr3(T + 1)) 'Cell to put total from sum formula
        r2 = CStr(arr4(T + 10)) 'Upper bound of range to sum
        r3 = CStr((arr4(T + 10) + 1)) 'Upper bound of range sum + 1
        r6 = CStr((arr5(T + 2)) + 2) 'Gives A5 on each card
        Range("G" + r2 + ":H" + r).BorderAround xlContinuous, xlMedium
        Range("G" + r2 + ":I" + r).BorderAround xlContinuous, xlMedium
        Range("A" + r3 + ":T" + r3).BorderAround xlContinuous, xlMedium
        Range("A" + r5 + ":A" + r6).BorderAround xlContinuous, xlMedium
        Range("A" + r5 + ":L" + r).BorderAround xlContinuous, xlThick
        Range("A" + r5 + ":T" + r).BorderAround xlContinuous, xlThick
        Range("A" + r5 + ":J" + r).BorderAround xlContinuous, xlThick
        Range("K" + r + ":T" + r).BorderAround xlContinuous, xlThick
        Range("A:T").Interior.Color = RGB(171, 255, 171)

        'fixes fonts below
        Range("M" + r2 + ":T" + r).NumberFormat = "$#,##0"
        Range("A" + ":T").Font.Name = "Calibri"
        'center all cells where bidders will type
        Range("M" + ":T").HorizontalAlignment = xlCenter

    Next

    For Each c In rng2.Cells
        If c.Value <> "" And c.Value = "ALT" Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
            n = c.Row
                'adds orange highlight to cards
                Range("M" + n + ":T" + n).Interior.ColorIndex = 44
        End If
    Next c




    
    For j = 0 To (NumRows - 1)
        Sheets("ALT PHASE CARD DUMP").Select
        Range(arr1(j * 2), arr1((j * 2) + 1)).Copy
        Sheets(arr(j + 1)).Select
        On Error GoTo addsheet
        Dim arrk() As String
        k = 0
        pen = 0
        d = 0
        v = 0
        v1 = 0
        v2 = 0
        With Sheets(arr(j + 1))
            k = .Cells(.Rows.Count, 7).End(xlUp).Row
            k = k + 5
            ReDim Preserve arrk(pen)
            arrk(pen) = k
            k = pen + 1
            'Debug.Print arrk(pen)
        End With
        w = CStr(arrk(d))
        w1 = CStr(arrk(d) + 9)
        v = CStr(arrk(d) + 3)
        v1 = CStr(arrk(d) + 9)
        v2 = CStr(arrk(d) + 10)
        Range("A" + CStr(arrk(d))).Select
        d = d + 1
        Range("M:T").ColumnWidth = 14
        ActiveSheet.Paste
        'Range("A3:T5").BorderAround xlContinuous, xlMedium
        'ActiveSheet.Range("A3:T10").BorderAround xlContinuous, xlThick
        'ActiveSheet.Range("A3:T11").BorderAround xlContinuous, xlThick
        'ActiveSheet.Range("A" + v + ":T" + v1).BorderAround xlContinuous, xlThick
        'ActiveSheet.Range("A" + v + ":T" + v2).BorderAround xlContinuous, xlThick
        Dim rng1 As Range
        Set rng1 = Range("A12:A100000") ' Identify your range
        n = 0
            For Each k In rng1.Cells
                If k.Value <> "" And k.Value = "CARD TOTAL MC2:" Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
                    n = k.Row
                    n2 = n + 8
                    n3 = n + 9
                    Range("G" + (CStr(n2))).Value = "Subcontractor in Add/Cut is:"
                    Range("G" + (CStr(n3))).Value = "Bid Amount in Add/Cut is:"
                    Range("M" + (CStr(n2))).Value = "(Only Bid Captain fills in, let them know if this does not match bid card.)"
                    Range("M" + (CStr(n3))).Value = "(Only Bid Captain fills in, let them know if this does not match bid card.)"
                    Range("G" + (CStr(n2))).Font.Size = "14"
                    Range("G" + (CStr(n3))).Font.Size = "14"
                    Range("K" + (CStr(n2)) + ":L" + (CStr(n2))).Merge
                    Range("K" + (CStr(n3)) + ":L" + (CStr(n3))).Merge
                    Range("K" + (CStr(n3)) + ":L" + (CStr(n3))).NumberFormat = "$#,##0"
                    Range("K" + (CStr(n2)) + ":L" + (CStr(n2))).BorderAround xlContinuous, xlThick
                    Range("K" + (CStr(n3)) + ":L" + (CStr(n3))).BorderAround xlContinuous, xlThick
                End If
            Next k
        'Debug.Print arrk(d) + 1
        'removes excess rows
        Range("A" + v + ":A" + w1).EntireRow.Delete
'        Debug.Print v
'        Debug.Print v1
'        Debug.Print v2

        Dim rng As Range
        Set rng = Range("K12:K100000") ' Identify your range
        d = 0
            For Each c In rng.Cells
                If c.Value <> "" And c.Value = "Sub Name:" Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""

                    d = c.Row
                    Rows(d).EntireRow.Delete
                    'Debug.Print d
                End If
            Next c
            
'        Dim rng3 As Range
'        Set rng3 = Range("A1:A100000") ' Identify your range
'            'highlight alternate rows
'            For Each c In rng3.Cells
'                If c.Value <> "" And c.Value = "ALTERNATE 01: ADD 2ND LEVEL" Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
'
'                    d = c.Row
'                    Debug.Print d
'
'                End If
'            Next c

            


        Application.ScreenUpdating = True
    Next
addsheet:
        On Error GoTo end1
        Sheets.Add.Name = arr(j + 1)
        Resume Next
end1:
        MsgBox ("Alternate Bid Cards are now copied to existing sheets. If new scopes were added, move those new sheets to follow numerical order of existing sheets.")





Application.Goto Reference:=Sheets("SHEET CREATOR").Range("A1")
End Sub




