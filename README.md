# VBA-Pattern-match
'Below is an a code I have been working on in VBA to search a 5 character long cell value against another sheet/database which has 1 less character.
'The idea here is to cater for instances where there a typo error has added an additional charater or the value is represented as a text with a zero at the front
'Looking to improve match time by limiting match percentage of two values being compared

Sub Pattern_Match()
    'Rudamentary pattern matching code to find a matches between a database with less characters than the search reference,
    'i.e. looking for 03183(search ref) in 3183(database/other excel spreadsheet)
    'in Visual Basic Application
    Dim PROD, WCRESULT As Variant
    Dim PRODCHECKR As Variant
    Dim PRODCHECKL As Variant
    Dim PRODCHECKBS As Variant
    Dim PRODLEN As Long
    Dim PRODCOLNUM As Long
    Dim RESCOLNUM As Long
    Dim CONDCOLNUM As Long
    
    Dim DBREADY As Long
    DBREADY = MsgBox("Is the report to check against open, otherwise cancel and restart", vbOKCancel, "File to match against")
    If DBREADY = 1 Then
        GoTo DBONHAND
    ElseIf DBREADY = 2 Then
        Exit Sub
    End If
DBONHAND:
    
    Dim INDRPT1COL As Range   'Index formula range
        Set INDRPT1COL = Workbooks("YF ILR - New Database.xlsx").Worksheets("Products").Range("F1:F" & Workbooks("YF ILR - New Database.xlsx").Worksheets("Products").Range("a1").CurrentRegion.Rows.Count)  'dynamic added 2022-07-26

    Dim MTHRPTCOL As Range   'Match formula range
        Set MTHRPTCOL = Workbooks("YF ILR - New Database.xlsx").Worksheets("Products").Range("F1:F" & Workbooks("YF ILR - New Database.xlsx").Worksheets("Products").Range("a1").CurrentRegion.Rows.Count)   'dynamic added 2022-07-26
    'Above two ranges can be assigned to arrays ot speed up search process
    
    Dim RPTime As Double  'added to see how long process takes, in order to look in to improvements
    RPTime = Timer
    
    Lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).row
    Lastcol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    PRODCOLNUM = InputBox("Product column to check", , "Number value here")
    CONDCOLNUM = InputBox("LKUP result column for If statement", , "Number value here")
    
    RESCOLNUM = Lastcol + 1  'adding an additional column to publish the result in
    For i = 2 To Lastrow
        'If i = 83 Then  'This is for debugging at a specific row where special attention is needed
            'Stop
        'End If
        If Cells(i, CONDCOLNUM).Text = "#N/A" Or Cells(i, CONDCOLNUM).Text = "" Then   'Or WorksheetFunction.IsNumber(WorksheetFunction.Find("Not found", Cells(i, CONDCOLNUM).Text))
            PROD = ActiveWorkbook.ActiveSheet.Cells(i, PRODCOLNUM).Text
            PRODLEN = Len(ActiveSheet.Cells(i, PRODCOLNUM).value)
            PRODCHECK = WorksheetFunction.Replace(PROD, 1, 1, "")
            PRODCHECKR = Left(PROD, PRODLEN - 1)  'to take charater away from front(left side)
            PRODCHECKL = Right(PROD, PRODLEN - 1)  'to take character away form back(Right side
            PRODCHECKBS = Left(Right(PROD, PRODLEN - 1), PRODLEN - 2)  'to take charater away from front and back (both ends) simultaneously
RECHECK:
            On Error Resume Next
            
            'Debug.Print PRODCHECK
            Cells(i, RESCOLNUM) = WorksheetFunction.Index(INDRPT1COL, _
                                WorksheetFunction.Match("*" & PRODCHECK & "*", _
                                MTHRPTCOL, _
                                0), _
                            1) & "- TH " & _
                            WorksheetFunction.Index(INDRPT5COL, _
                                WorksheetFunction.Match("*" & PRODCHECK & "*", _
                                MTHRPTCOL, _
                                0), _
                            5)   'this needs to be given as column number you want from DB being checked
            If Cells(i, RESCOLNUM) = "" And PRODLEN > 3 Then
                PRODCHECK = WorksheetFunction.Replace(PRODCHECK, 1, 1, "")
                PRODLEN = PRODLEN - 1
                On Error GoTo 0
                On Error GoTo -1
                GoTo RECHECK
            ElseIf PRODLEN = 3 Then
                Cells(i, RESCOLNUM) = "No pattern match"
            Else
                Cells(i, RESCOLNUM) = "No pattern match"
            End If
        End If
    Next i
MsgBox ("Time is:" & (Timer - RPTime) & "secs")

INDRPT1COL = Nothing
INDRPT5COL = Nothing
MTHRPTCOL = Nothing
PRODCHECK = Nothing
PRODLEN = Empty
PROD = Nothing
PRODCOLNUM = Empty
CONDCOLNUM = Empty
RESCOLNUM = Empty
DBREADY = Empty
RPTime = Empty
End Sub
