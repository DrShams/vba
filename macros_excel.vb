Sub new_day()
'
' new_day Macro
' that macro create a new day for the report
'
' Keyboard Shortcut: Ctrl+n
'
	'Assign new values and make some calculations
    Dim curr_day As String, next_day As String, prev_day As String
    curr_day = ActiveSheet.Name
    oldDate_f = Replace(ActiveSheet.Name, ".", "/")
    prev_day = DateAdd("d", -1, oldDate_f)
    prev_day = Format(prev_day, "dd.mm")
    next_day = DateAdd("d", 1, oldDate_f) 'add 1 day only
    next_day = Format(next_day, "dd.mm")
    'Copy and create a new sheet
    Set ws = Sheets(ActiveSheet.Name)
    ws.Select					'1 SELECTION
    ws.Copy Before:=Sheets("inventory") 'create new day
    ActiveSheet.Name = next_day
    'Replace dates in formulas
     Cells.Replace What:=prev_day, Replacement:=curr_day, LookAt:=xlPart, _
       SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
      ReplaceFormat:=False
    'An erray of cells Define which of them equals to 0
	Sheets(curr_day).Select		'2	SELECTION
    Dim rg As Range
    Set rg = ThisWorkbook.Worksheets(curr_day).Range("M63:Y150")
    Dim params As Variant
    params = rg.Value
    Dim y As Long, x As Long'y - vertical (1,2,3...), x - horizontal(A,B,C...)
    For y = LBound(params) To UBound(params) 'LBound - first position; UBound- last position
        For x = LBound(params, 2) To UBound(params, 2)
        If Not IsEmpty(params(y, x)) Then
            If params(y, x) = 0 Then
                Set mc = Worksheets(curr_day).Cells(y + 63, x + 13) 'The problem is in here we return precisely address for instance 1 1 will be A1 but in our Case it is actually M63
                'Range("M63:Y150")
					'A B C D E F G H I G K L M - on x row +13
					'63 - on y row +63
					'Kludge - Fixme UP!
                Range(mc.Address()).Select
                With Selection.Interior
                .Color = 49407 'just highlight that cells
                End With
            Else
				' Macrocolor Macro - color cells from the previous report to white
				'Kludge - Fixme UP!
				Set mc = Worksheets(curr_day).Cells(y + 63, x + 13)
				Range(mc.Address()).Select
				Selection.Copy
				Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
					xlNone, SkipBlanks:=False, Transpose:=False
				Application.CutCopyMode = False
				With Selection.Interior
					.PatternColorIndex = xlAutomatic
					.ThemeColor = xlThemeColorDark1
				End With
            End If
        End If
        Next x
    Next y
End Sub

