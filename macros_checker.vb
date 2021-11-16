'At first add the library -> Tools -> References -> Microsoft VBScript Regular Expressions 5.5 -> ok
Public Const range_plan = "AC63:AC150" 'Change for the range with plan mud parameters
Public Const range_fact = "M63:Y150" 'Change for the range with fact mud parameters
'Make Sure if there is correct Data format(day/month) of time in Windows 10 it should be United States format

Function RegExGet(aString As String, myexp As String) As Variant
Dim RegEx As New VBScript_RegExp_55.RegExp
Dim newArray() As String
Dim x As Integer
Dim cnt As Integer
RegEx.Pattern = myexp
RegEx.IgnoreCase = True
RegEx.Global = True
Set Matches = RegEx.Execute(aString)
x = Matches.Count
ReDim newArray(x - 1) As String
cnt = 0
    For Each Match In Matches
        newArray(cnt) = Match.Value
        cnt = cnt + 1
    Next
     RegExGet = newArray()
End Function

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
Dim sht As Worksheet

If wb Is Nothing Then Set wb = ThisWorkbook
On Error Resume Next
Set sht = wb.Sheets(shtName)
On Error GoTo 0
WorksheetExists = Not sht Is Nothing
End Function
'
'Public total_params As Integer
'Public wrong_params As Integer

Private Sub mud_checker()
    Dim arr() As Double
    Set ws = Sheets(ActiveSheet.Name)
    ws.Select
'Take all plan parameters
    Set rg = ws.Range(range_plan)
    
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Dim params As Variant
    params = rg.Value
    Dim y As Long, x As Long
    Dim y_row As Integer, x_column As Integer
    y_row = CInt(RegExGet(range_fact, "\d+")(0))
    x_column = CInt(Asc(RegExGet(range_fact, "\w")(0)) - 64) '64 because A in Asc code = 65, B = 66, C = 67 etc...
	Dim celvalue As Double
    'Make visible only this area
    ActiveWindow.ScrollRow = y_row
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.Zoom = 30
    Dim trim_str As String
    Dim LArray() As String
    ReDim arr(1 To rg.Rows.Count, 2) '0 - first parameter, 1 - second parameter
    For Z = 1 To rg.Rows.Count
        arr(Z, 0) = 0.0 '-min possible value
        arr(Z, 1) = 999999.999 '-max possible value for LSRV 30000.000 ? haven't seen anything more than that...
    Next Z
    Dim gel_min As Double, gel_max As Double
    Dim gel As Boolean
    gel = False
    Dim f1 As Double, f2 As Double
    For y = LBound(params) To UBound(params)
        For x = LBound(params, 2) To UBound(params, 2)
            If Not IsEmpty(params(y, x)) And Not params(y, x) = "-" Then 'skip empty cells and cells which contain only '-' sign

                f1 = 0'min
                f2 = 999999.999'max
                'clean with replacing empty cells
                trim_str = WorksheetFunction.Trim(params(y, x))
                trim_str = Replace(trim_str, " ", "")
				if Not trim_str Like "*[0-9]*" Then
					Debug.Print "Not Found Any number in string:", trim_str
                ElseIf InStr(trim_str, "±") Then 'density if ± defines, CDec takes string and converts to Decimal or Float?
                    f1 = CDec(RegExGet(trim_str, "\d\.\d+")(0)) - CDec(RegExGet(trim_str, "\d\.\d+")(1)) 'min
                    f2 = CDec(RegExGet(trim_str, "\d\.\d+")(0)) + CDec(RegExGet(trim_str, "\d\.\d+")(1)) 'max
                ElseIf InStr(trim_str, "-") Then
                    If InStr(trim_str, "/") Then
                        LArray = Split(trim_str, "/")
                        f1 = CDec(RegExGet(trim_str, "\d+")(0))
                        f2 = CDec(RegExGet(trim_str, "\d+")(1))
                        gel_min = CDec(RegExGet(trim_str, "\d+")(2))
                        gel_max = CDec(RegExGet(trim_str, "\d+")(3))
                    ElseIf InStr(trim_str, ".") Then
                        LArray = Split(trim_str, "-")
                        f1 = CDec(LArray(0))
                        f2 = CDec(LArray(1))
                    Else
						f1 = CDec(RegExGet(trim_str, "\d+")(0))
						f2 = CDec(RegExGet(trim_str, "\d+")(1))
                    End If
                ElseIf InStr(trim_str, ChrW(&H2265)) Then 'greater or equal in (u)nicode for instance >=oil percentage
                    f1 = CDec(RegExGet(trim_str, "\d+")(0))
                    'Debug.Print "ASC", f1, Asc(trim_str)
                ElseIf InStr(trim_str, ChrW(&H2264)) Then 'less or equal in (u)nicode for instance <=MBT
                    LArray = Split(trim_str, ChrW(&H2264))
                    f2 = CDec(LArray(1))
                ElseIf InStr(trim_str, "<") Then
                    LArray = Split(trim_str, "<")
                    f2 = CDec(LArray(1)) - 0.01
                End If
            arr(y, 0) = CDbl(f1)
            arr(y, 1) = CDbl(f2)
            'Debug.Print "min = " & arr(y, 0), "max = " & arr(y, 1), y
            End If
        Next x
    Next y
'Take all fact parameters and check with plan
    Set rg = ws.Range(range_fact)
    params = rg.Value
    curr_day = ActiveSheet.Name
    For y = LBound(params) To UBound(params) 'LBound - first position; UBound- last position
        For x = LBound(params, 2) To UBound(params, 2)
        If Not IsEmpty(params(y, x)) Then
            Set mc = Worksheets(curr_day).Cells(y_row + y - 1, x_column + x) 'BUG FIX IT WHY IT IS -1 I DON't KNOW PREVIOUSLY IT WORKS WITHOUT '-1'
            Range(mc.Address()).Select
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            If params(y, x) = 0 Then
                Selection.Interior.Color = 49407 'Orange color
            Else '<> the same as != in C++
                If arr(y, 0) <> 0 Or arr(y, 1) <> 999999.999 Then
'Get Gels 10s/10min
                    If InStr(params(y, x), Chr(47)) Then '/ the same as Chr(47)
                        LArray = Split(params(y, x), Chr(47))
                        If CDec(LArray(0)) < arr(y, 0) Or CDec(LArray(0)) > arr(y, 1) Then
                            'wrong_params = wrong_params + 1
                            Debug.Print "Detected_gel10s", LArray(0), "min = " & arr(y, 0), "max = " & arr(y, 1) 'arr 0 1 only for Gels 10s
                            If gel = False Then 'If just 1 gel wrong - orange color if 2 gel wrong - red color
                                Selection.Interior.Color = 49407 'Orange color
                                gel = True
                            Else
                                Selection.Interior.Color = 255 'RED color
                            End If
                        ElseIf LArray(1) < gel_min Or LArray(1) > gel_max Then 'BUG !!!SOMETIMES THAT CONDITIONS GOES BEFORE first IF STATEMENT!
                            'wrong_params = wrong_params + 1
                            Debug.Print "Detected_gel10m", LArray(1), "min = " & gel_min, "max = " & gel_max
                            If gel = False Then 'If just 1 gel wrong - orange color if 2 gel wrong - red color
                                Selection.Interior.Color = 49407 'Orange color
                                gel = True
                            Else
                                Selection.Interior.Color = 255 'RED color
                            End If
                        Else
                            'Debug.Print LArray(0), "min = " & arr(y, 0), "max = " & arr(y, 1), "CORRECT GELS"
                        End If
					Else
'All other parameters	
						if params(y, x) = "-" Then
							celvalue = 0.0
						else
							celvalue = Cdbl(params(y, x))
						end if
						If celvalue < arr(y, 0) Or celvalue > arr(y, 1) Then
							'wrong_params = wrong_params + 1
							Selection.Interior.Color = 255 'RED color
							Debug.Print "Detected", celvalue, "min = " & arr(y, 0), "max = " & arr(y, 1)
						Else
							Debug.Print celvalue, "min = " & arr(y, 0), "max = " & arr(y, 1)
						End If
					End if
                Else
                    'total_params = total_params + 1
                    'Debug.Print params(y, x), "min = " & arr(y, 0), "max = " & arr(y, 1), "CORRECT PARAMETERS"
                End If
            End If
        End If
        Next x
    Next y
End Sub

Private Sub new_day()
'Assign new values and make some calculations
    Dim curr_day As String, next_day As String, prev_day As String
    curr_day = ActiveSheet.Name
    oldDate_f = Replace(curr_day, ".", "/")
    prev_day = DateAdd("d", -1, oldDate_f)
    prev_day = Format(prev_day, "dd.mm")
    next_day = DateAdd("d", 1, oldDate_f) 'add 1 day only
    next_day = Format(next_day, "dd.mm")
'Copy and create a new sheet
    Set ws = Sheets(ActiveSheet.Name)
    ws.Select                   '1 SELECTION
    If WorksheetExists("inventory") = True Then
        ws.Copy Before:=Sheets("inventory") 'create new day
    Else
        ws.Copy After:=Sheets(Sheets.Count) 'create new day
    End If
    
    ActiveSheet.Name = next_day
'Replace dates in formulas
     Cells.Replace What:=prev_day, Replacement:=curr_day, LookAt:=xlPart, _
       SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
      ReplaceFormat:=False
'An erray of cells Define which of them equals to 0
    Sheets(curr_day).Select     '2  SELECTION
    Dim rg As Range
    Set rg = ThisWorkbook.Worksheets(curr_day).Range(range_fact)
'Paint cells which we going to check
    rg.Select
    'Selection.Interior.Color = 15773696 'Blue color ?
'Copy&Paste values which contain info or skip empty cells
    Dim params As Variant
    params = rg.Value
    Dim y As Long, x As Long 'y - vertical (1,2,3...), x - horizontal(A,B,C...)
    Dim y_row As Integer, x_column As Integer
    y_row = CInt(RegExGet(range_fact, "\d+")(0))
    x_column = CInt(Asc(RegExGet(range_fact, "\w")(0)) - 64) '64 because A in Asc code = 65, B = 66, C = 67 etc...
    For y = LBound(params) To UBound(params) 'LBound - first position; UBound- last position
        For x = LBound(params, 2) To UBound(params, 2)
        If Not IsEmpty(params(y, x)) Then
            If params(y, x) = 0 Then
                Set mc = Worksheets(curr_day).Cells(y + y_row, x + x_column) 'The problem is in here we return precisely address for instance 1 1 will be A1 but in our Case it is actually M63
                Range(mc.Address()).Select
                Selection.Interior.Color = 49407 'just highlight that cells
            Else
                Set mc = Worksheets(curr_day).Cells(y + y_row, x + x_column)
                Range(mc.Address()).Select
                Selection.Copy
            'KLUDGE FIX ME UP FROM GARBAGE
                Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                    xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
            'KLUDGE FIX ME UP FROM GARBAGE
            End If
        End If
        Next x
    Next y
End Sub

Sub mud_reports()
    Result = MsgBox("Would you like to check the parameters on this day?", vbYesNoCancel + vbDefaultButton2)
    If Result = vbYes Then
        Call mud_checker
        'Debug.Print "in the sum " & total_params "parameters was checked " & wrong_params & " wrong parameters was found"
    Else
    End If
    Result = MsgBox("Would you like to create a new day?", vbYesNoCancel + vbDefaultButton2)
    If Result = vbYes Then
        Call new_day
    Else
    End If
    Result = MsgBox("Would you like to check all the parameters for all days?", vbYesNoCancel + vbDefaultButton2)
    If Result = vbYes Then
        Dim i As Integer
        For i = 1 To Worksheets.Count
            Sheets(i).Select
            Call mud_checker
        Next i
    Else
    End If
End Sub


