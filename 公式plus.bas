Attribute VB_Name = "¹«Ê½Plus"
Public data_title

Function P_ABOUT()
    about_form.Show
    about_form.TextBox1.Locked = True
End Function
Function P_HELP()
    ActiveWorkbook.FollowHyperlink "https://ladeng6666.github.io/wiki/formula-plus"
    P_HELP = "ÇëÔÚ´ò¿ªµÄÍøÒ³ÖÐ²é¿´"
End Function
Function P_SUM_MERGE(ByVal sum_area As range, Optional ByVal direction As Integer = 0)
    Dim thisCell As range, new_area As range
    Set thisCell = Application.thisCell
    
    If thisCell.MergeCells Then
        If direction = 0 Then
            Set new_area = Rows(thisCell.Row & ":" & thisCell.Row + thisCell.MergeArea.Rows.Count - 1)
        Else
            Set new_area = Columns(Chr(thisCell.Column + 64) & ":" & Chr(thisCell.Column + thisCell.MergeArea.Columns.Count - 1 + 64))
 
        End If
        Set sum_area = Intersect(sum_area, new_area)
    End If
    
    If Not sum_area Is Nothing Then
        P_SUM_MERGE = WorksheetFunction.Sum(sum_area)
    Else
        P_SUM_MERGE = 0
    End If
End Function
Function MAXIFS(ByVal max_range, ByVal criteria_range1, ByVal criteria1, ParamArray criterias())
    Dim var_len As Integer, last_r As Long
    Dim offset_c As Integer, checkOK As Boolean, max_value
    Dim cri_range As range, cri As String

    var_len = UBound(criterias) + 1
    If var_len Mod 2 = 1 Then
        Application.Evaluate ("" * 1)
    End If
    

    last_r = WorksheetFunction.Min(max_range.Rows.Count, ActiveSheet.usedRange.Rows.Count)
    max_value = -Cells.Rows.Count
    For r = 1 To last_r
        offset_c = criteria_range1.Column - max_range.Column
        checkOK = max_range.Cells(r).Offset(0, offset_c) = criteria1

        
        If var_len >= 2 Then
            For i = 0 To var_len - 2 Step 2
                Set cri_range = criterias(i)
                offset_c = cri_range.Column - max_range.Column
                cri = criterias(i + 1)
                checkOK = checkOK And max_range.Cells(r).Offset(0, offset_c) = criterias(i + 1)
            Next
        End If

        If checkOK Then
            max_value = WorksheetFunction.Max(max_value, max_range.Cells(r))
        End If
    Next
    If max_value = -Cells.Rows.Count Then
        max_value = "" + 1
    End If
    MAXIFS = max_value
End Function

Function MINIFS(ByVal min_range, ByVal criteria_range1, ByVal criteria1, ParamArray criterias())
    Dim var_len As Integer, last_r As Long
    Dim offset_c As Integer, checkOK As Boolean, min_value
    Dim cri_range As range, cri As String

    var_len = UBound(criterias) + 1
    If var_len Mod 2 = 1 Then
        Application.Evaluate ("" * 1)
    End If
    

    last_r = WorksheetFunction.Min(min_range.Rows.Count, ActiveSheet.usedRange.Rows.Count)
    min_value = Cells.Rows.Count
    For r = 1 To last_r
        offset_c = criteria_range1.Column - min_range.Column
        checkOK = min_range.Cells(r).Offset(0, offset_c) = criteria1
        
        If var_len >= 2 Then
            For i = 0 To var_len - 2 Step 2
                Set cri_range = criterias(i)
                offset_c = cri_range.Column - min_range.Column
                cri = criterias(i + 1)
                checkOK = checkOK And min_range.Cells(r).Offset(0, offset_c) = criterias(i + 1)
            Next
        End If

        If checkOK Then
            min_value = WorksheetFunction.Min(min_value, min_range.Cells(r))
        End If
    Next
    If min_value = Cells.Rows.Count Then
        min_value = "" + 1
    End If
    MINIFS = min_value
End Function

Function P_RMB(money As String)
    Dim rmb As String, ints As String, decimals As String
    
    If Val(money) = 0 Then rmb = "": Exit Function
    
    Dim yy() As String
    rmb = Application.WorksheetFunction.text(Round(Val(money), 2) + 0.001, "[DBNum2]")
    ints = VBA.Split(rmb, ".")(0)
    decimals = VBA.Split(rmb, ".")(1)
    
    j = Mid$(decimals, 1, 1)
    f = Mid$(decimals, 2, 1)
    If ints = "Áã" Then ints = "" Else ints = ints & "Ôª"
    If f = "Áã" Then
        If j = "Áã" Then rmb = ints & "Õû"
        If j <> "Áã" Then rmb = ints & j & "½ÇÕû"
    Else
        rmb = ints & j & "½Ç" & f & "·Ö"
    End If
    P_RMB = rmb
End Function


Function P_EN(ByVal txt As String, Optional needNum As Integer = 0)
    Dim biaodianCN As String, biaodianEN
    biaodianCN = "\u3002|\uff1f|\uff01|\uff0c|\u3001|\uff1b|\uff1a|\u201c|\u201d|\u2018|\u2019|\uff08|\uff09|\u300a|\u300b|\u3008|\u3009|\u3010|\u3011|\u300e|\u300f|\u300c|\u300d|\ufe43|\ufe44|\u3014|\u3015|\u2026|\u2014|\uff5e|\ufe4f|\uffe5"
    biaodianEN = "*,\.;\:"" '!/"
    p = ""
    
    If needNum Then
        p = P_REG_REPLACE(txt, "[^a-zA-Z\d\s\-\\" & biaodianCN & biaodianEN & "]")
    Else
        p = P_REG_REPLACE(txt, "[^a-zA-Z\s\-\\" & biaodianCN & biaodianEN & "]")
    End If
    p = P_REG_REPLACE(p, "(\n)+", Chr(10))
    
    P_EN = p
End Function
Function P_NUM(ByVal txt As String, Optional index_num As Integer = 0)
    Dim p As String
    p = ""
    
    p = P_REG_REPLACE(txt, "[^0-9\.\d\s\-\\]")
    p = P_REG_REPLACE(p, "(\n)+", Chr(10))
    
    If index_num <> 0 Then
        p = P_REG_FIND(txt, "(\d+)", 1)(index_num)
    End If
    
    P_NUM = p
End Function

Function P_CN(ByVal txt As String, Optional needNum As Integer = 0)
    Dim biaodianCN As String, biaodianEN
    biaodianCN = "\u3002|\uff1f|\uff01|\uff0c|\u3001|\uff1b|\uff1a|\u201c|\u201d|\u2018|\u2019|\uff08|\uff09|\u300a|\u300b|\u3008|\u3009|\u3010|\u3011|\u300e|\u300f|\u300c|\u300d|\ufe43|\ufe44|\u3014|\u3015|\u2026|\u2014|\uff5e|\ufe4f|\uffe5"
    biaodianEN = ",\.;\:"" '!"
    
    If needNum Then
        p = P_REG_REPLACE(txt, "[^\u4e00-\u9fa5\d" & biaodianCN & biaodianEN & "]")
    Else
        p = P_REG_REPLACE(txt, "[^\u4e00-\u9fa5" & biaodianCN & biaodianEN & "]")
    End If
    p = P_REG_REPLACE(p, "(\n)+", Chr(10))
    P_CN = p
End Function


Function P_REG_REPLACE(ByVal txt As String, ByVal reg As String, Optional ByVal targetTxt As String = "")
    Dim regx
    Dim p As String
    Set regx = CreateObject("vbscript.regexp")
    With regx
        .Global = True
        .IgnoreCase = True
        .Pattern = reg
    End With
        
    p = regx.Replace(txt, targetTxt)
    
    P_REG_REPLACE = p
    Set regx = Nothing
    
End Function
Function P_REG_FIND(ByVal txt As String, ByVal reg As String, Optional returnType As Integer = 0)
    Dim regx
    Dim p
    Set regx = CreateObject("vbscript.regexp")
    With regx
        .Global = True
        .IgnoreCase = True
        .Pattern = reg
    End With
        
    Set p = regx.Execute(txt)
    
    If p.Count > 0 Then
        If returnType = 0 Then
            P_REG_FIND = True
        ElseIf returnType = 1 Then
            P_REG_FIND = matchCollectionToArray(p, 1)
        ElseIf returnType = 2 Then
            P_REG_FIND = p.Count
        ElseIf returnType = 3 Then
            P_REG_FIND = matchCollectionToArray(p, 2)
        End If
    Else
        P_REG_FIND = False
    End If
    Set regx = Nothing
    
End Function

Private Function matchCollectionToArray2(ByVal mc)
    Dim i As Integer, find_result()
    i = 1
    For Each m In mc

        ReDim Preserve find_result(1 To mc.Count)
        find_result(i) = m.value
        i = i + 1
    Next
    matchCollectionToArray = find_result
End Function

Private Function matchCollectionToArray(ByVal mc, Optional ByVal typ = 0)
    Dim i As Integer, j As Integer, find_result(), p As String
    i = 1
    If typ = 0 Then
        For Each m In mc
            
            ReDim Preserve find_result(1 To i + 1)
            find_result(i) = m.value
            i = i + 1
        Next
    ElseIf typ = 1 Then
        For Each m In mc
            If m.submatches.Count > 0 Then
                For Each subm In m.submatches
                    ReDim Preserve find_result(1 To i + 1)
                    find_result(i) = subm
                    i = i + 1
                Next
            Else
                ReDim Preserve find_result(1 To i + 1)
                find_result(i) = m.value
                i = i + 1
            End If
        Next
    ElseIf typ = 2 Then
        For Each m In mc
            If m.submatches.Count > 0 Then
                For Each subm In m.submatches
                    p = p & subm
                Next
            Else
                p = p & m.value
            End If
        Next
    End If
    matchCollectionToArray = IIf(p <> "", p, find_result)
End Function


Function P_TEXTJOIN(ByVal range, ByVal splitter As String, Optional ignoreBlank As Boolean = False, Optional ignoreRepeat As Boolean = False, Optional actionForMerge As Integer = 0, Optional ByVal columnFirst As Boolean = True)
    Dim r As Long, c As Long, last_c As Long, last_r As Long, txt As String
    
    If TypeName(range) = "String" Then
        P_TEXTJOIN = range
        Exit Function
    End If
    
    Dim usedRange As range
    Set usedRange = ActiveSheet.usedRange
    Set range = Intersect(usedRange, range)

    Dim result As String
    result = ""
    
    r = 1
    c = 1

    last_c = range.Columns.Count
    last_r = range.Rows.Count

    If Not columnFirst Then
        While c <= last_c
            r = 1
            While r <= last_r
                txt = range.Columns(c).Cells(r)
                If txt = "" And actionForMerge = 1 And range.Rows(r).Cells(c).MergeCells Then
                    txt = range.Rows(r).Cells(c).MergeArea.Cells(1)
                End If
                If txt <> "" Or (txt = "" And ignoreBlank = False) Then
                    If result = "" Then
                        result = txt
                    Else
                        If ignoreRepeat Then
                            If Not result Like "*" & txt & "*" Then
                                result = result & splitter & txt
                            End If
                            
                        Else
                            result = result & splitter & txt
                        End If
                    End If
                End If
                r = r + 1
            Wend
            c = c + 1
        Wend
    Else
        While r <= last_r
            c = 1
            While c <= last_c
                txt = range.Rows(r).Cells(c)
                If txt = "" And actionForMerge = 1 And range.Rows(r).Cells(c).MergeCells Then
                    txt = range.Rows(r).Cells(c).MergeArea.Cells(1)
                End If
                If txt <> "" Or (txt = "" And ignoreBlank = False) Then
                    If result = "" Then
                        result = txt
                    Else
                        If ignoreRepeat Then
                            If Not result Like "*" & txt & "*" Then
                                result = result & splitter & txt
                            End If
                            
                        Else
                            result = result & splitter & txt
                        End If
                        
                    End If
                End If
                c = c + 1
            Wend
            r = r + 1
        Wend
    End If
    P_TEXTJOIN = result
End Function
Function P_SPLIT(ByVal txtRange, ByVal splitter As String, ByVal get_index As Integer, Optional ByVal returnType As Integer = 0)
    Dim txt As String, p As String, total As Integer

    txt = P_TEXTJOIN(txtRange, splitter, 1, 1)
    total = P_COUNTIF(txtRange, splitter)
    
    get_index = IIf(get_index < 0, total + get_index + 1, get_index)
    get_index = IIf(get_index >= total, total, get_index)
    
    
    If returnType = 0 Then
        p = Split(txt, splitter)(get_index - 1)
    ElseIf returnType = 1 Then
        For i = 1 To get_index
            If p = "" Then
                p = Split(txt, splitter)(i - 1)
            Else
                p = p & splitter & Split(txt, splitter)(i - 1)
            End If
        Next
    ElseIf returnType = -1 Then
        For i = get_index To total
            If p = "" Then
                p = Split(txt, splitter)(i - 1)
            Else
                p = p & splitter & Split(txt, splitter)(i - 1)
            End If
        Next
    End If
    
    P_SPLIT = p
    
End Function
Function P_SPLIT2(ByVal txtRange As range, ByVal splitter1 As String, ByVal splitter2 As String, ByVal get_index As Integer)
    Dim p As String, reg As String
    reg = ""

    If splitter1 = splitter2 Then
        P_SPLIT2 = P_SPLIT(txtRange, splitter1, get_index)
        GoTo exit_split
    End If
    
    reg = "\" & splitter1 & "([^\" & splitter2 & "]*)\" & splitter2
    If index_num <> 0 Then
        p = P_REG_REPLACE(txtRange, reg)
    Else
        p = P_REG_FIND(txtRange, reg, 1)(get_index)
    End If
    p = Replace(p, splitter1, "")
    p = Replace(p, splitter2, "")
    P_SPLIT2 = p
exit_split:
End Function

Function P_COUNTIF(ByVal txtRange, ByVal str As String)
    Dim p As String, txt As String
    txt = P_TEXTJOIN(txtRange, "", 1, 1)
    p = P_REG_FIND(txt, str, 2)
    
    P_COUNTIF = p
End Function



Function P_SHEETDATA(ByVal sheetIndex As String, ByVal dataCell As range, Optional ByVal offsetr As Long = 0, Optional ByVal keyColumn As range = Nothing)
    
    If IsNumeric(sheetIndex) Then
        P_SHEETDATA = Worksheets(CInt(sheetIndex)).range(dataCell.Address).Offset(offsetr)
    Else
        Dim i As Integer, j As Integer, last_r As Long, r_num As Long, first_r As Long, data_r As Long
        i = Split(sheetIndex, "-", 2)(0)
        j = Split(sheetIndex, "-", 2)(1)
        last_r = 0
        first_r = dataCell.Row
        
        For n = i To j
            r_num = last_r
            If Not keyColumn Is Nothing Then
                last_r = last_r + Worksheets(n).range(keyColumn.Address).End(xlDown).Row - first_r + 1
            Else
                last_r = last_r + Worksheets(n).usedRange.Rows.Count - first_r + 1
            End If
            If offsetr <= last_r Then
                    data_r = offsetr - r_num + first_r - 1
                Exit For
            End If
        Next
        P_SHEETDATA = Worksheets(n).Cells(data_r, dataCell.Column)
    End If
End Function

Function P_PHONE(phoneNo As String, Optional ByVal index As Integer = 1)
    P_PHONE = P_REG_FIND(phoneNo, "1[3578]\d{9}", 1)(index)
End Function

Function P_SHENFENZHENG(id As String, getType As Integer)
    Dim info As String
    Dim shengxiao()
    shengxiao = Array("Êó", "Å£", "»¢", "ÍÃ", "Áú", "Éß", "Âí", "Ñò", "ºï", "¼¦", "¹·", "Öí")
    
    Select Case getType
    
        Case 2
            info = CDate(WorksheetFunction.text(Mid(id, 7, 8), "0000-00-00"))
        Case 5
            info = IIf(Mid(id, 17, 1) Mod 2 = 0, "Å®", "ÄÐ")
        Case 3
            info = Year(Date) - Mid(id, 7, 4)
        Case 4
            info = shengxiao((Mid(id, 7, 4) - 1900) Mod 12)
        Case 1
            info = WorksheetFunction.WebService("http://www.ladeng6666.com/app/shenfenzheng/index.php?id=" & Left(id, 6))
    End Select
    P_SHENFENZHENG = info
End Function

Function P_SHOUZIMU(text As String)
    Dim p As String
    text = " " & text
    p = P_REG_FIND(text, "(\s[A-Za-z])", 3)
    p = Replace(p, " ", "")
    P_SHOUZIMU = p
End Function

Function P_PINYIN(text As String, Optional ByVal shouzimu As Boolean = True) As Variant
    On Error Resume Next
    If text = "" Then Exit Function
    If shouzimu Then
        Dim pinyin As String, hanzi As String, i As Integer
        i = 1
        hanzi = Left(text, 1)
next_hanzi:
        
        pinyin = pinyin & Application.WorksheetFunction.VLookup(hanzi, [{"ß¹","A";"°Ë","B";"àê","C";"´î","D";"¶ê","E";"·¢","F";"¸Á","G";"îþ","H";"»÷","J";"ßÇ","K";"À¬","L";"Âè","M";"ÄÃ","N";"àÞ","O";"Å¾","P";"Æß","Q";"È»","R";"Øí","S";"Ëû","T";"ÍÚ","W";"Ï¦","X";"Ñ¹","Y";"ÔÓ","Z"}], 2)
        
        If i < Len(text) Then
            i = i + 1
            hanzi = Mid(text, i, 1)
            GoTo next_hanzi
        End If
    Else
        pinyin = WorksheetFunction.WebService("http://www.ladeng6666.com/app/pinyin/index.php?hanzi=" & text)
    End If
    
    P_PINYIN = pinyin
End Function




Function IFS(ByVal condition, ByVal result, ParamArray vars())
    Dim p As String, var_len As Integer
    p = ""

    var_len = UBound(vars) + 1
    If Application.Evaluate(CBool(condition)) Then
        p = result
    Else
        If var_len = 1 Then
            p = vars(0)
        Else
            For i = 0 To var_len - 2 Step 2
                If Application.Evaluate(CBool(vars(i))) Then
                    p = vars(i + 1)
                    GoTo exit_ifs
                End If
            Next
            If var_len Mod 2 = 1 Then
                p = vars(var_len - 1)
            Else
                Application.Evaluate ("" * 1)
            End If
        End If
        
    End If
    
exit_ifs:
    IFS = p
End Function

Function SWITCH(ByVal value, ByVal condition, ByVal result, ParamArray vars())
    Dim p As String, var_len As Integer
    p = ""

    var_len = UBound(vars) + 1

    If Not P_REG_FIND(condition, "[<>=]") Then
        condition = "=" & condition
    End If
    
    If pEvaluate(value, condition) Then
        p = result
    Else
        For i = 0 To var_len - 2 Step 2
            condition = vars(i)
            If Not P_REG_FIND(condition, "[<>=]") Then
                condition = "=" & condition
            End If
            If pEvaluate(value, condition) Then
                p = vars(i + 1)
                GoTo exit_switch
            End If
        Next
        If var_len Mod 2 = 1 Then
            p = vars(var_len - 1)
        Else
            Application.Evaluate ("" * 1)
        End If
        
    End If
    
exit_switch:
    SWITCH = p
End Function

Private Function setTitleDict(ByVal dataSheet As Worksheet, ByVal titler As Integer, Optional titlec As Integer = 1)
    Dim c As Integer
    Set data_title = CreateObject("Scripting.Dictionary")
    data_title.RemoveAll
    c = titlec
    With dataSheet
        While .Cells(titler, c) <> ""
            data_title.Add .Cells(titler, c).value, c
            c = c + 1
        Wend
    End With
End Function


Private Function pEvaluate(ByVal value As String, ByVal condition As String)

    On Error GoTo err_handle  'Óöµ½´íÎóÏòÏÂÖ´ÐÐ

    pEvaluate = Application.Evaluate(value & condition)
    a = 1 + pEvaluate
    
    GoTo end_fun
    
err_handle:

    operation = Left(condition, 1)
    condition = Right(condition, Len(condition) - 1)
    If operation = "=" Then
            pEvaluate = value = condition
    ElseIf operation = ">" Then
            pEvaluate = value > condition
    ElseIf operation = "<" Then
            pEvaluate = value < condition

    End If
end_fun:

End Function

Function P_LOOKUP(ByVal lookup_value, ByVal lookup_array As range, ByVal lookup_rule, Optional ByVal return_array = Nothing)
    Dim offset_c, last_r, lookup_column
    If return_array Is Nothing Then Set return_array = lookup_array
    
    Set lookup_column = lookup_array.Columns(1)
    offset_c = return_array.Column - lookup_column.Column
    
    'Êý×Ö
    If IsNumeric(lookup_rule) Then
        Dim findResult
        Dim totalFind As Integer
        If lookup_rule > 0 Then
            If lookup_value = "0" Then
                totalFind = WorksheetFunction.CountIf(lookup_array, ">0")
                Set findResult = findByIndex(lookup_value, lookup_array, totalFind - lookup_rule + 1)
            Else
                Set findResult = findByIndex(lookup_value, lookup_array, lookup_rule)
            End If
        Else
            
            If lookup_value = "0" Then
                Set findResult = findByIndex(lookup_value, lookup_array, -lookup_rule)
            Else
                totalFind = WorksheetFunction.CountIf(lookup_array, lookup_value)
                Set findResult = findByIndex(lookup_value, lookup_array, totalFind + lookup_rule + 1)
            End If
        End If
        
        P_LOOKUP = findResult.Offset(0, offset_c)
        
    ElseIf lookup_value.Cells.Count > 0 Then
        Dim lookValues, lookValueFirst As range, valuesCount, lookValueFirst_c As Integer, lookValueFirst_total As Integer
        Dim ruleTitles, ruleTitleFirst As range, i As Integer
        Dim dataTitles, dataFirstColumn As range
        Dim firstValueTotal As Integer, first_c As Integer
        
        Set lookValues = lookup_value.Rows(1)
        Set lookValueFirst = lookValues.Cells(1)
        
        Set ruleTitles = lookup_rule.Rows(1)
        Set rultTitleFirst = ruleTitles.Cells(1)
        
        Set dataTitles = lookup_array.Rows(1)

        lookValueFirst_c = WorksheetFunction.Match(rultTitleFirst, dataTitles, 0)
        Set dataFirstColumn = lookup_array.Columns(lookValueFirst_c)
        
        lookValueFirst_total = WorksheetFunction.CountIf(dataFirstColumn, lookValueFirst)
        
        i = 1
        While i <= lookValueFirst_total
            isFind = True
            Set findResult = findByIndex(lookValueFirst, dataFirstColumn, i)
            For c = 1 To ruleTitles.Cells.Count
                data_c = WorksheetFunction.Match(ruleTitles.Cells(c), dataTitles, 0) - lookValueFirst_c
                If lookValues.Cells(c) <> findResult.Offset(0, data_c) Then
                    isFind = False
                End If
            Next
            If isFind Then
                GoTo exit_mulFind
            Else
                Set findResult = Nothing
            End If
            i = i + 1
        Wend
exit_mulFind:
'        MsgBox findResult.Address & "," & offset_c
        P_LOOKUP = findResult.Offset(0, offset_c)

    End If
End Function

Function P_ROW(Optional ByVal cell As range, Optional ByVal ignoreEmpty As Boolean = False)
    Dim thisCell As range, thisr As Long, thisc As Long, actualr As Long

    If Not cell Is Nothing Then
        Set thisCell = cell
    Else
        Set thisCell = Application.thisCell
    End If
    
    thisr = thisCell.Row
    thisc = thisCell.Column
    actualr = 0

    If thisr = 1 Then
        actualr = 1
        GoTo exit_row
    End If
    If ignoreEmpty Then
        actualr = WorksheetFunction.CountA(range(Cells(1, thisc), thisCell.Offset(-1))) + 1
        GoTo exit_row
    End If
    
    For r = 1 To thisr
        actualr = actualr + 1
        If Cells(r, thisc).MergeCells Then
            r = r + Cells(r, thisc).MergeArea.Rows.Count - 1
        End If
    Next
exit_row:
    P_ROW = actualr
End Function

Function P_COLUMN(Optional ByVal cell As range, Optional ByVal ignoreEmpty As Boolean = False)
    Dim thisCell As range, thisr As Long, thisc As Long, actualc As Long

    If Not cell Is Nothing Then
        Set thisCell = cell
    Else
        Set thisCell = Application.thisCell
    End If
    
    thisr = thisCell.Row
    thisc = thisCell.Column
    actualc = 0

    If thisc = 1 Then
        actualc = 1
        GoTo exit_row
    End If
    If ignoreEmpty Then
        actualc = WorksheetFunction.CountA(range(Cells(thisr, 1), thisCell.Offset(, -1))) + 1
        GoTo exit_row
    End If
    
    For c = 1 To thisc
        actualc = actualc + 1
        If Cells(thisr, c).MergeCells Then
            c = c + Cells(thisr, c).MergeArea.Columns.Count - 1
        End If
    Next
exit_row:
    P_COLUMN = actualc
End Function

Function P_OFFSET(ByVal reference As range, ByVal index As Long, ByVal gap As Integer, Optional ByVal direction As Integer = 0)
    index = index - 1
    If direction = 0 Then
        P_OFFSET = reference.Offset(index * gap)
    Else
        P_OFFSET = reference.Offset(, index * gap)
    End If
End Function

Function P_INDEX(ByVal reference1 As range, ByVal reference2 As range, ByVal index As Long)
    Dim gap As Long, direction As String, r As Long, c As Long
    r = reference1.Row
    c = reference1.Column
    
    index = index - 1
    If reference1.Row = reference2.Row Then
        direction = "H"
        gap = Abs(reference1.Column - reference2.Column)
        P_INDEX = reference1.Parent.Cells(r, c + gap * index)

    ElseIf reference1.Column = reference2.Column Then
        direction = "V"
        gap = Abs(reference1.Row - reference2.Row)

        P_INDEX = reference1.Parent.Cells(r + gap * index, c)
    End If
End Function
Private Function findByIndex(ByVal look_value, ByVal look_array As range, Optional ByVal no = 1)
    Dim k As Long, i As Integer
    i = 1
    If look_value = "0" Then
        While i <= no
            k = WorksheetFunction.Match(100000, look_array, 1)
            If k > 0 Then i = i + 1
            If i <= no Then
                Set look_array = look_array.Resize(k - 1)
            End If
        Wend
    Else
        While i <= no
            k = WorksheetFunction.Match(look_value, look_array, 0)
            If k > 0 Then i = i + 1
            If i <= no Then
                Set look_array = look_array.Resize(look_array.Cells.Count - k)
                Set look_array = look_array.Offset(k)
            End If
        Wend
    End If
    If k > 0 Then
        Set findByIndex = look_array.Cells(k)
    End If
End Function

Function P_RANDOM(rnd_type As Integer)
    Dim first3, rand3 As Long, other As Long, rnd_length As Integer
    rnd_length = 7
    first3 = Array(139, 135, 170, 173, 150, 131, 181, 151)
    rand3 = CInt(Rnd * UBound(first3))
    other = CLng(Rnd * 9 * 10 ^ rnd_length + 10 ^ rnd_length)
    P_RANDOM = first3(rand3) & other

End Function


Function P_UNIQUE(rng As range, index As Integer)
    Dim D As Object, cell As range, i As Integer
    i = 1

    For Each cell In Intersect(rng, ActiveSheet.usedRange)
        If WorksheetFunction.CountIf(rng, rng.Cells(i)) = 1 Then
            i = i + 1
        End If
        
        If i >= index Then
            result = rng.Cells(i)
            Exit For
        End If
    Next
    P_UNIQUE = result
End Function
