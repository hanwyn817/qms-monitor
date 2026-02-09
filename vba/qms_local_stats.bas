Attribute VB_Name = "qms_local_stats"
Option Explicit

' QMS local overdue statistics and extraction (VBA-only)
' - No LLM
' - No PDF
' - Reads Config sheet and ledger Excel files directly
' - Writes overdue detail, summary and warnings

Private Const SHEET_CONFIG As String = "Config"
Private Const SHEET_OVERDUE As String = "Overdue_Events"
Private Const SHEET_SUMMARY As String = "Summary"
Private Const SHEET_LOG As String = "Log"
Private Const SHEET_PRECHECK As String = "Precheck_Report"

Private Const HEADER_HINTS As String = "申请时间|发起日期|计划完成日期|完成日期|状态|编号|内容|责任人|责任部门|分管"
Private Const OPEN_WORKBOOK_SECURITY_FORCE_DISABLE As Long = 3
Private Const PRECHECK_MAX_DATE_SAMPLES As Long = 5

Public Sub RunQmsLocalStats()
    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Dim reportDate As Date
    If Not PromptReportDate(reportDate) Then GoTo CleanExit

    Dim wsConfig As Worksheet
    Set wsConfig = GetConfigSheet()

    Dim wsOverdue As Worksheet
    Dim wsSummary As Worksheet
    Dim wsLog As Worksheet
    Set wsOverdue = EnsureSheet(SHEET_OVERDUE)
    Set wsSummary = EnsureSheet(SHEET_SUMMARY)
    Set wsLog = EnsureSheet(SHEET_LOG)

    Dim warnings As Collection
    Set warnings = New Collection
    Dim errors As Collection
    Set errors = New Collection

    Dim validationIssues As Object
    Set validationIssues = CollectConfigValidationIssues(wsConfig, True, True)
    AppendIssuesToCollection validationIssues, errors

    If errors.Count > 0 Then
        WriteValidationLog wsLog, warnings, errors, "配置校验失败，未开始统计。"
        MsgBox "配置校验失败（" & CStr(errors.Count) & "项），未开始统计。" & vbCrLf & _
               "示例: " & PreviewCollection(errors, 3) & vbCrLf & _
               "请查看 Log 工作表完整明细。", vbCritical
        GoTo CleanExit
    End If

    Dim configs As Collection
    Set configs = LoadConfigs(wsConfig, warnings)
    If configs.Count = 0 Then
        warnings.Add "配置文件中没有可用配置。"
        WriteLogSheet wsLog, warnings
        MsgBox "没有可用配置，请检查 Config 工作表。", vbExclamation
        GoTo CleanExit
    End If

    Dim openStatusRules As Object
    On Error Resume Next
    Set openStatusRules = BuildOpenStatusRules(configs)
    If Err.Number <> 0 Then
        errors.Add Err.Description
        Err.Clear
    End If
    On Error GoTo CleanFail

    If errors.Count > 0 Then
        WriteValidationLog wsLog, warnings, errors, "状态规则校验失败，未开始统计。"
        MsgBox "状态规则校验失败（" & CStr(errors.Count) & "项）。" & vbCrLf & _
               "示例: " & PreviewCollection(errors, 3) & vbCrLf & _
               "请查看 Log 工作表完整明细。", vbCritical
        GoTo CleanExit
    End If

    Dim stats As Object
    Set stats = InitStats()

    Dim overdueEvents As Collection
    Set overdueEvents = New Collection

    Dim processedFiles As Long
    Dim skippedFiles As Long
    processedFiles = 0
    skippedFiles = 0

    Dim cfg As Object
    For Each cfg In configs
        ProcessOneConfig cfg, reportDate, openStatusRules, stats, overdueEvents, warnings, processedFiles, skippedFiles
    Next cfg

    WriteOverdueSheet wsOverdue, overdueEvents
    WriteSummarySheet wsSummary, reportDate, stats, processedFiles, skippedFiles, warnings.Count
    WriteLogSheet wsLog, warnings

    MsgBox "统计完成。" & vbCrLf & _
           "超期事件: " & CStr(overdueEvents.Count) & vbCrLf & _
           "成功读取台账: " & CStr(processedFiles) & vbCrLf & _
           "跳过台账: " & CStr(skippedFiles), vbInformation

CleanExit:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Exit Sub

CleanFail:
    WriteErrorToLog wsLog, "运行失败: " & Err.Description
    MsgBox "运行失败: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Public Sub RunQmsPrecheckReport()
    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Dim wsConfig As Worksheet
    Dim wsPrecheck As Worksheet
    Dim wsLog As Worksheet
    Set wsConfig = GetConfigSheet()
    Set wsPrecheck = EnsureSheet(SHEET_PRECHECK)
    Set wsLog = EnsureSheet(SHEET_LOG)

    Dim warnings As Collection
    Set warnings = New Collection
    Dim errors As Collection
    Set errors = New Collection

    Dim validationIssues As Object
    Set validationIssues = CollectConfigValidationIssues(wsConfig, True, True)
    AppendIssuesToCollection validationIssues, errors

    Dim reportRows As Collection
    Set reportRows = BuildPrecheckRows(wsConfig, validationIssues)

    WritePrecheckSheet wsPrecheck, reportRows
    WriteValidationLog wsLog, warnings, errors, "已生成预检查报告（未执行统计）。"

    Dim errCount As Long
    errCount = errors.Count
    MsgBox "预检查完成（未执行统计）。" & vbCrLf & _
           "检查项: " & CStr(reportRows.Count) & vbCrLf & _
           "校验错误: " & CStr(errCount) & vbCrLf & _
           "告警: " & CStr(warnings.Count) & vbCrLf & _
           "详细结果请查看 Precheck_Report 与 Log。", vbInformation

CleanExit:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Exit Sub

CleanFail:
    WriteErrorToLog wsLog, "预检查失败: " & Err.Description
    MsgBox "预检查失败: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Function PromptReportDate(ByRef reportDate As Date) As Boolean
    Dim raw As String
    raw = InputBox("请输入统计基准日期（YYYY-MM-DD）：", "QMS Local Stats", Format$(Date, "yyyy-mm-dd"))
    If Len(raw) = 0 Then Exit Function

    Dim parsed As Date
    If Not TryParseDate(raw, parsed) Then
        MsgBox "日期格式无效，请输入 YYYY-MM-DD。", vbExclamation
        Exit Function
    End If

    reportDate = DateSerial(Year(parsed), Month(parsed), Day(parsed))
    PromptReportDate = True
End Function

Private Function GetConfigSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(Trim$(ws.Name), SHEET_CONFIG, vbTextCompare) = 0 Then
            Set GetConfigSheet = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 2001, "GetConfigSheet", "未找到配置工作表: " & SHEET_CONFIG
End Function

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set EnsureSheet = ws
End Function

Private Function LoadConfigs(ByVal ws As Worksheet, ByVal warnings As Collection) As Collection
    Dim configs As Collection
    Set configs = New Collection

    Dim lastRow As Long
    lastRow = GetConfigLastRow(ws)
    If lastRow < 2 Then
        Set LoadConfigs = configs
        Exit Function
    End If

    Dim r As Long
    For r = 2 To lastRow
        Dim topic As String
        Dim moduleName As String
        Dim yearValue As String
        Dim filePath As String
        Dim sheetNameRaw As Variant

        topic = Trim$(ToText(ws.Cells(r, 2).Value2))
        moduleName = Trim$(ToText(ws.Cells(r, 3).Value2))
        yearValue = ParseYear(ws.Cells(r, 4).Value2)
        filePath = Trim$(ToText(ws.Cells(r, 5).Value2))
        sheetNameRaw = NormalizeSheetName(ws.Cells(r, 6).Value2)

        If Len(moduleName) = 0 And Len(filePath) = 0 Then GoTo ContinueRow

        Dim idCol As Long, contentCol As Long, initiatedCol As Long
        idCol = ColToIndex(ws.Cells(r, 7).Value2)
        contentCol = ColToIndex(ws.Cells(r, 8).Value2)
        initiatedCol = ColToIndex(ws.Cells(r, 9).Value2)

        If Len(moduleName) = 0 Then
            warnings.Add "Config第" & CStr(r) & "行缺失质量模块，已跳过。"
            GoTo ContinueRow
        End If
        If Len(filePath) = 0 Then
            warnings.Add "Config第" & CStr(r) & "行缺失文件路径，已跳过: 模块=" & moduleName
            GoTo ContinueRow
        End If
        If idCol <= 0 Or contentCol <= 0 Or initiatedCol <= 0 Then
            warnings.Add "Config第" & CStr(r) & "行核心列(G/H/I)缺失或非法，已跳过: 模块=" & moduleName
            GoTo ContinueRow
        End If

        Dim plannedCol As Long, plannedDueDays As Long
        Dim hasPlannedCol As Boolean, hasPlannedDueDays As Boolean
        ParsePlannedRule ws.Cells(r, 10).Value2, r, moduleName, warnings, plannedCol, plannedDueDays, hasPlannedCol, hasPlannedDueDays

        Dim cfg As Object
        Set cfg = CreateObject("Scripting.Dictionary")
        cfg("row_no") = r
        cfg("topic") = topic
        cfg("module") = moduleName
        cfg("year") = yearValue
        cfg("file_path") = filePath
        cfg("sheet_name") = sheetNameRaw
        cfg("id_col") = idCol
        cfg("content_col") = contentCol
        cfg("initiated_col") = initiatedCol
        cfg("has_planned_col") = hasPlannedCol
        cfg("planned_col") = plannedCol
        cfg("has_planned_due_days") = hasPlannedDueDays
        cfg("planned_due_days") = plannedDueDays
        cfg("status_col") = ColToIndexOptional(ws.Cells(r, 11).Value2)
        cfg("owner_dept_col") = ColToIndexOptional(ws.Cells(r, 12).Value2)
        cfg("owner_col") = ColToIndexOptional(ws.Cells(r, 13).Value2)
        cfg("qa_col") = ColToIndexOptional(ws.Cells(r, 14).Value2)
        cfg("qa_manager_col") = ColToIndexOptional(ws.Cells(r, 15).Value2)
        cfg("open_status_value") = Trim$(ToText(ws.Cells(r, 16).Value2))
        cfg("data_start_row") = ParseDataStartRow(ws.Cells(r, 17).Value2, r, moduleName, warnings)

        configs.Add cfg

ContinueRow:
    Next r

    Set LoadConfigs = configs
End Function

Private Function BuildOpenStatusRules(ByVal configs As Collection) As Object
    Dim rules As Object
    Set rules = CreateObject("Scripting.Dictionary")

    Dim errors As Collection
    Set errors = New Collection

    Dim cfg As Object
    For Each cfg In configs
        Dim moduleName As String
        Dim openStatusValue As String
        moduleName = Trim$(ToText(cfg("module")))
        openStatusValue = Trim$(ToText(cfg("open_status_value")))

        If Len(moduleName) = 0 Then GoTo ContinueRule

        If Len(openStatusValue) = 0 Then
            errors.Add "Config第" & CStr(cfg("row_no")) & "行 模块[" & moduleName & "]缺少未完成状态值。"
            GoTo ContinueRule
        End If

        If rules.Exists(moduleName) Then
            If CStr(rules(moduleName)) <> openStatusValue Then
                errors.Add "模块[" & moduleName & "]存在多个未完成状态值: [" & CStr(rules(moduleName)) & "] 与 [" & openStatusValue & "]。"
            End If
        Else
            rules(moduleName) = openStatusValue
        End If

ContinueRule:
    Next cfg

    If errors.Count > 0 Then
        Err.Raise vbObjectError + 2002, "BuildOpenStatusRules", "未完成状态值配置错误: " & JoinCollection(errors, "; ")
    End If

    Set BuildOpenStatusRules = rules
End Function

Private Sub ProcessOneConfig( _
    ByVal cfg As Object, _
    ByVal reportDate As Date, _
    ByVal openStatusRules As Object, _
    ByVal stats As Object, _
    ByVal overdueEvents As Collection, _
    ByVal warnings As Collection, _
    ByRef processedFiles As Long, _
    ByRef skippedFiles As Long)

    Dim sourcePath As String
    sourcePath = ResolvePath(CStr(cfg("file_path")))

    Dim values As Variant
    Dim sheetNameResolved As String
    Dim errMsg As String
    If Not ReadSheetValues(sourcePath, cfg("sheet_name"), values, sheetNameResolved, errMsg) Then
        warnings.Add "模块[" & CStr(cfg("module")) & "] 文件读取失败，已跳过: " & sourcePath & " (" & errMsg & ")"
        skippedFiles = skippedFiles + 1
        Exit Sub
    End If
    processedFiles = processedFiles + 1

    Dim rowCount As Long, colCount As Long
    rowCount = UBound(values, 1)
    colCount = UBound(values, 2)

    If rowCount <= 1 Then
        warnings.Add "模块[" & CStr(cfg("module")) & "] 表内容为空或只有表头: " & sourcePath & " / " & CStr(cfg("sheet_name"))
        Exit Sub
    End If

    Dim startRow As Long
    startRow = CLng(cfg("data_start_row"))
    If startRow < 2 Then startRow = 2
    If startRow > rowCount Then
        warnings.Add "模块[" & CStr(cfg("module")) & "] 数据起始行[" & CStr(startRow) & "]超出范围: " & sourcePath & " / " & CStr(cfg("sheet_name"))
        Exit Sub
    End If

    Dim moduleName As String
    moduleName = CStr(cfg("module"))
    If Not openStatusRules.Exists(moduleName) Then
        warnings.Add "模块[" & moduleName & "]未配置未完成状态值，已跳过该配置行。"
        skippedFiles = skippedFiles + 1
        Exit Sub
    End If
    Dim openStatusValue As String
    openStatusValue = Trim$(CStr(openStatusRules(moduleName)))

    Dim r As Long
    For r = startRow To rowCount
        Dim eventId As String, content As String, initiatedRaw As String
        eventId = GetCellText(values, r, CLng(cfg("id_col")))
        content = GetCellText(values, r, CLng(cfg("content_col")))
        initiatedRaw = GetCellText(values, r, CLng(cfg("initiated_col")))

        If Len(eventId) = 0 And Len(content) = 0 And Len(initiatedRaw) = 0 Then GoTo ContinueEvent

        Dim initiatedDate As Date
        Dim hasInitiatedDate As Boolean
        hasInitiatedDate = TryParseDate(initiatedRaw, initiatedDate)

        Dim plannedDate As Date
        Dim hasPlannedDate As Boolean
        hasPlannedDate = False

        If CBool(cfg("has_planned_due_days")) Then
            If hasInitiatedDate Then
                plannedDate = DateAdd("d", CLng(cfg("planned_due_days")), initiatedDate)
                hasPlannedDate = True
            End If
        ElseIf CBool(cfg("has_planned_col")) Then
            Dim plannedRaw As String
            plannedRaw = GetCellText(values, r, CLng(cfg("planned_col")))
            hasPlannedDate = TryParseDate(plannedRaw, plannedDate)
        ElseIf hasInitiatedDate Then
            plannedDate = AddOneMonthCompat(initiatedDate)
            hasPlannedDate = True
        End If

        Dim statusValue As String
        Dim ownerDept As String
        Dim ownerName As String
        Dim qaName As String
        Dim qaManagerName As String
        statusValue = GetCellText(values, r, CLng(cfg("status_col")))
        ownerDept = GetCellText(values, r, CLng(cfg("owner_dept_col")))
        ownerName = GetCellText(values, r, CLng(cfg("owner_col")))
        qaName = GetCellText(values, r, CLng(cfg("qa_col")))
        qaManagerName = GetCellText(values, r, CLng(cfg("qa_manager_col")))

        If Not hasInitiatedDate And IsHeaderLikeRow(eventId, content, initiatedRaw) Then GoTo ContinueEvent

        If Not hasInitiatedDate And Len(initiatedRaw) > 0 Then
            warnings.Add "模块[" & moduleName & "] 行" & CStr(r) & "发起日期解析失败: '" & initiatedRaw & "' (" & sourcePath & "/" & sheetNameResolved & ")"
        End If

        stats("total_count") = CLng(stats("total_count")) + 1
        AddCount stats("by_year_total"), CStr(cfg("year")), 1
        AddCount stats("by_module_total"), moduleName, 1

        Dim isOpen As Boolean
        isOpen = (Trim$(statusValue) = openStatusValue)

        Dim isOverdue As Boolean
        isOverdue = (hasPlannedDate And plannedDate < reportDate And isOpen)
        If isOverdue Then
            stats("overdue_count") = CLng(stats("overdue_count")) + 1
            AddCount stats("by_year_overdue"), CStr(cfg("year")), 1
            AddCount stats("by_module_overdue"), moduleName, 1
            AddCount stats("by_qa_overdue"), qaName, 1
            AddCount stats("by_qa_manager_overdue"), qaManagerName, 1
            AddCount stats("by_owner_dept_overdue"), ownerDept, 1

            Dim ev As Object
            Set ev = CreateObject("Scripting.Dictionary")
            ev("topic") = CStr(cfg("topic"))
            ev("module") = moduleName
            ev("year") = CStr(cfg("year"))
            ev("event_id") = eventId
            ev("content") = content
            ev("initiated_date") = IIf(hasInitiatedDate, Format$(initiatedDate, "yyyy-mm-dd"), "")
            ev("planned_date") = IIf(hasPlannedDate, Format$(plannedDate, "yyyy-mm-dd"), "")
            ev("status") = statusValue
            ev("owner_dept") = ownerDept
            ev("owner") = ownerName
            ev("qa") = qaName
            ev("qa_manager") = qaManagerName
            ev("source_file") = sourcePath
            ev("source_sheet") = sheetNameResolved
            ev("source_row") = r
            overdueEvents.Add ev
        End If

ContinueEvent:
    Next r
End Sub

Private Function InitStats() As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s("total_count") = 0&
    s("overdue_count") = 0&
    Set s("by_year_total") = CreateObject("Scripting.Dictionary")
    Set s("by_year_overdue") = CreateObject("Scripting.Dictionary")
    Set s("by_module_total") = CreateObject("Scripting.Dictionary")
    Set s("by_module_overdue") = CreateObject("Scripting.Dictionary")
    Set s("by_qa_overdue") = CreateObject("Scripting.Dictionary")
    Set s("by_qa_manager_overdue") = CreateObject("Scripting.Dictionary")
    Set s("by_owner_dept_overdue") = CreateObject("Scripting.Dictionary")
    Set InitStats = s
End Function

Private Sub AddCount(ByVal d As Object, ByVal key As String, ByVal delta As Long)
    Dim k As String
    k = Trim$(key)
    If Len(k) = 0 Then Exit Sub
    If d.Exists(k) Then
        d(k) = CLng(d(k)) + delta
    Else
        d(k) = delta
    End If
End Sub

Private Sub WriteOverdueSheet(ByVal ws As Worksheet, ByVal overdueEvents As Collection)
    ws.Cells.Clear

    Dim headers As Variant
    headers = Array("主题", "质量模块", "年份", "编号", "内容", "发起日期", "计划完成日期", "状态", "责任部门", "责任人", "分管QA", "分管QA中层", "来源文件", "来源Sheet", "来源行")

    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    ws.Rows(1).Font.Bold = True

    If overdueEvents.Count = 0 Then
        ws.Columns("A:O").AutoFit
        Exit Sub
    End If

    Dim arr() As Variant
    ReDim arr(1 To overdueEvents.Count, 1 To 15)

    For i = 1 To overdueEvents.Count
        Dim ev As Object
        Set ev = overdueEvents(i)
        arr(i, 1) = ev("topic")
        arr(i, 2) = ev("module")
        arr(i, 3) = ev("year")
        arr(i, 4) = ev("event_id")
        arr(i, 5) = ev("content")
        arr(i, 6) = ev("initiated_date")
        arr(i, 7) = ev("planned_date")
        arr(i, 8) = ev("status")
        arr(i, 9) = ev("owner_dept")
        arr(i, 10) = ev("owner")
        arr(i, 11) = ev("qa")
        arr(i, 12) = ev("qa_manager")
        arr(i, 13) = ev("source_file")
        arr(i, 14) = ev("source_sheet")
        arr(i, 15) = ev("source_row")
    Next i

    ws.Range("A2").Resize(overdueEvents.Count, 15).Value = arr
    ws.Columns("A:O").AutoFit

    On Error Resume Next
    ws.Range("A1").CurrentRegion.Sort Key1:=ws.Range("B2"), Order1:=xlAscending, _
                                       Key2:=ws.Range("G2"), Order2:=xlAscending, _
                                       Key3:=ws.Range("D2"), Order3:=xlAscending, _
                                       Header:=xlYes
    On Error GoTo 0
End Sub

Private Sub WriteSummarySheet( _
    ByVal ws As Worksheet, _
    ByVal reportDate As Date, _
    ByVal stats As Object, _
    ByVal processedFiles As Long, _
    ByVal skippedFiles As Long, _
    ByVal warningCount As Long)

    ws.Cells.Clear
    Dim r As Long
    r = 1

    ws.Cells(r, 1).Value = "报告日期"
    ws.Cells(r, 2).Value = Format$(reportDate, "yyyy-mm-dd")
    r = r + 1
    ws.Cells(r, 1).Value = "成功读取台账"
    ws.Cells(r, 2).Value = processedFiles
    r = r + 1
    ws.Cells(r, 1).Value = "跳过台账"
    ws.Cells(r, 2).Value = skippedFiles
    r = r + 1
    ws.Cells(r, 1).Value = "告警数"
    ws.Cells(r, 2).Value = warningCount
    r = r + 2

    Dim totalCount As Long
    Dim overdueCount As Long
    totalCount = CLng(stats("total_count"))
    overdueCount = CLng(stats("overdue_count"))

    ws.Cells(r, 1).Value = "总体统计"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Cells(r, 1).Value = "总事件数"
    ws.Cells(r, 2).Value = totalCount
    r = r + 1
    ws.Cells(r, 1).Value = "超期事件数"
    ws.Cells(r, 2).Value = overdueCount
    r = r + 1
    ws.Cells(r, 1).Value = "超期占比"
    ws.Cells(r, 2).Value = PercentText(overdueCount, totalCount)
    r = r + 2

    r = WriteYearSection(ws, r, stats("by_year_total"), stats("by_year_overdue"))
    r = WriteModuleSection(ws, r, stats("by_module_total"), stats("by_module_overdue"))
    r = WriteRankSection(ws, r, "超期按分管QA统计（降序）", "分管QA", stats("by_qa_overdue"))
    r = WriteRankSection(ws, r, "超期按分管QA中层统计（降序）", "分管QA中层", stats("by_qa_manager_overdue"))
    r = WriteRankSection(ws, r, "超期按责任部门统计（降序）", "责任部门", stats("by_owner_dept_overdue"))

    ws.Columns("A:E").AutoFit
End Sub

Private Function WriteYearSection(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalDict As Object, ByVal overdueDict As Object) As Long
    Dim r As Long
    r = startRow

    ws.Cells(r, 1).Value = "按年度统计"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    ws.Cells(r, 1).Value = "年份"
    ws.Cells(r, 2).Value = "总事件数"
    ws.Cells(r, 3).Value = "超期事件数"
    ws.Cells(r, 4).Value = "超期占比"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Font.Bold = True
    r = r + 1

    Dim keys As Variant
    keys = SortedKeysAsc(totalDict)
    If IsEmpty(keys) Then
        ws.Cells(r, 1).Value = "无数据"
        WriteYearSection = r + 2
        Exit Function
    End If

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim y As String
        Dim totalCount As Long
        Dim overdueCount As Long
        y = CStr(keys(i))
        totalCount = CLng(totalDict(y))
        If overdueDict.Exists(y) Then overdueCount = CLng(overdueDict(y))

        ws.Cells(r, 1).Value = y
        ws.Cells(r, 2).Value = totalCount
        ws.Cells(r, 3).Value = overdueCount
        ws.Cells(r, 4).Value = PercentText(overdueCount, totalCount)
        r = r + 1
    Next i

    WriteYearSection = r + 1
End Function

Private Function WriteModuleSection(ByVal ws As Worksheet, ByVal startRow As Long, ByVal totalDict As Object, ByVal overdueDict As Object) As Long
    Dim r As Long
    r = startRow

    ws.Cells(r, 1).Value = "按质量模块统计"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    ws.Cells(r, 1).Value = "质量模块"
    ws.Cells(r, 2).Value = "总事件数"
    ws.Cells(r, 3).Value = "超期事件数"
    ws.Cells(r, 4).Value = "超期占比"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Font.Bold = True
    r = r + 1

    Dim keys As Variant
    keys = SortedKeysAsc(totalDict)
    If IsEmpty(keys) Then
        ws.Cells(r, 1).Value = "无数据"
        WriteModuleSection = r + 2
        Exit Function
    End If

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim moduleName As String
        Dim totalCount As Long
        Dim overdueCount As Long
        moduleName = CStr(keys(i))
        totalCount = CLng(totalDict(moduleName))
        If overdueDict.Exists(moduleName) Then overdueCount = CLng(overdueDict(moduleName))

        ws.Cells(r, 1).Value = moduleName
        ws.Cells(r, 2).Value = totalCount
        ws.Cells(r, 3).Value = overdueCount
        ws.Cells(r, 4).Value = PercentText(overdueCount, totalCount)
        r = r + 1
    Next i

    WriteModuleSection = r + 1
End Function

Private Function WriteRankSection( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal title As String, _
    ByVal keyHeader As String, _
    ByVal rankDict As Object) As Long

    Dim r As Long
    r = startRow

    ws.Cells(r, 1).Value = title
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    ws.Cells(r, 1).Value = keyHeader
    ws.Cells(r, 2).Value = "超期数"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Bold = True
    r = r + 1

    Dim pairs As Variant
    pairs = SortedPairsByCountDesc(rankDict)
    If IsEmpty(pairs) Then
        ws.Cells(r, 1).Value = "无数据"
        WriteRankSection = r + 2
        Exit Function
    End If

    Dim i As Long
    For i = 1 To UBound(pairs, 1)
        ws.Cells(r, 1).Value = CStr(pairs(i, 1))
        ws.Cells(r, 2).Value = CLng(pairs(i, 2))
        r = r + 1
    Next i

    WriteRankSection = r + 1
End Function

Private Sub WriteLogSheet(ByVal ws As Worksheet, ByVal warnings As Collection)
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "时间"
    ws.Cells(1, 2).Value = "告警信息"
    ws.Rows(1).Font.Bold = True

    If warnings.Count = 0 Then
        ws.Cells(2, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        ws.Cells(2, 2).Value = "无告警"
        ws.Columns("A:B").AutoFit
        Exit Sub
    End If

    Dim i As Long
    For i = 1 To warnings.Count
        ws.Cells(i + 1, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        ws.Cells(i + 1, 2).Value = CStr(warnings(i))
    Next i

    ws.Columns("A:B").AutoFit
End Sub

Private Sub WriteErrorToLog(ByVal ws As Worksheet, ByVal message As String)
    If ws Is Nothing Then Exit Sub
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "时间"
    ws.Cells(1, 2).Value = "错误"
    ws.Rows(1).Font.Bold = True
    ws.Cells(2, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(2, 2).Value = message
    ws.Columns("A:B").AutoFit
End Sub

Private Sub WriteValidationLog(ByVal ws As Worksheet, ByVal warnings As Collection, ByVal errors As Collection, ByVal summaryText As String)
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "时间"
    ws.Cells(1, 2).Value = "级别"
    ws.Cells(1, 3).Value = "信息"
    ws.Rows(1).Font.Bold = True

    Dim rowNo As Long
    rowNo = 2

    ws.Cells(rowNo, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(rowNo, 2).Value = "INFO"
    ws.Cells(rowNo, 3).Value = summaryText
    rowNo = rowNo + 1

    Dim i As Long
    For i = 1 To errors.Count
        ws.Cells(rowNo, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        ws.Cells(rowNo, 2).Value = "ERROR"
        ws.Cells(rowNo, 3).Value = CStr(errors(i))
        rowNo = rowNo + 1
    Next i

    For i = 1 To warnings.Count
        ws.Cells(rowNo, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        ws.Cells(rowNo, 2).Value = "WARN"
        ws.Cells(rowNo, 3).Value = CStr(warnings(i))
        rowNo = rowNo + 1
    Next i

    ws.Columns("A:C").AutoFit
End Sub

Private Function CollectConfigValidationIssues( _
    ByVal wsConfig As Worksheet, _
    ByVal checkPathReadable As Boolean, _
    ByVal checkSheetReadable As Boolean) As Object

    Dim issues As Object
    Set issues = CreateObject("Scripting.Dictionary")

    ValidateConfigHeaders wsConfig, issues

    Dim moduleStatus As Object
    Set moduleStatus = CreateObject("Scripting.Dictionary")
    Dim moduleStatusRow As Object
    Set moduleStatusRow = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = GetConfigLastRow(wsConfig)
    If lastRow < 2 Then
        AddIssue issues, 0, "Config 未发现有效数据行（第2行起为空）。"
        Set CollectConfigValidationIssues = issues
        Exit Function
    End If

    Dim r As Long
    For r = 2 To lastRow
        Dim moduleName As String
        Dim filePath As String
        Dim sheetRef As Variant
        moduleName = Trim$(ToText(wsConfig.Cells(r, 3).Value2))
        filePath = Trim$(ToText(wsConfig.Cells(r, 5).Value2))
        sheetRef = NormalizeSheetName(wsConfig.Cells(r, 6).Value2)

        If Len(moduleName) = 0 And Len(filePath) = 0 Then GoTo ContinueRow

        If Len(moduleName) = 0 Then AddIssue issues, r, "缺失质量模块（C列）。"
        If Len(filePath) = 0 Then AddIssue issues, r, "缺失文件路径（E列）。"

        ValidateColumnRule wsConfig.Cells(r, 7).Value2, r, "G", "编号列", True, issues
        ValidateColumnRule wsConfig.Cells(r, 8).Value2, r, "H", "内容列", True, issues
        ValidateColumnRule wsConfig.Cells(r, 9).Value2, r, "I", "发起日期列", True, issues
        ValidatePlannedRule wsConfig.Cells(r, 10).Value2, r, moduleName, issues
        ValidateColumnRule wsConfig.Cells(r, 11).Value2, r, "K", "状态列", False, issues
        ValidateColumnRule wsConfig.Cells(r, 12).Value2, r, "L", "责任部门列", False, issues
        ValidateColumnRule wsConfig.Cells(r, 13).Value2, r, "M", "责任人列", False, issues
        ValidateColumnRule wsConfig.Cells(r, 14).Value2, r, "N", "分管QA列", False, issues
        ValidateColumnRule wsConfig.Cells(r, 15).Value2, r, "O", "分管QA中层列", False, issues
        ValidateDataStartRule wsConfig.Cells(r, 17).Value2, r, moduleName, issues

        Dim openStatus As String
        openStatus = Trim$(ToText(wsConfig.Cells(r, 16).Value2))
        If Len(moduleName) > 0 Then
            If Len(openStatus) = 0 Then
                AddIssue issues, r, "缺失未完成状态值（P列）。"
            Else
                If moduleStatus.Exists(moduleName) Then
                    If CStr(moduleStatus(moduleName)) <> openStatus Then
                        AddIssue issues, r, "同模块未完成状态值冲突（P列）。"
                        AddIssue issues, CLng(moduleStatusRow(moduleName)), "同模块未完成状态值冲突（P列）。"
                    End If
                Else
                    moduleStatus(moduleName) = openStatus
                    moduleStatusRow(moduleName) = r
                End If
            End If
        End If

        If checkPathReadable And Len(filePath) > 0 Then
            Dim resolvedPath As String
            resolvedPath = ResolvePath(filePath)
            If Not IsAbsolutePath(filePath) And Len(ThisWorkbook.Path) = 0 Then
                AddIssue issues, r, "当前统计工作簿尚未保存，无法可靠解析相对路径[" & filePath & "]。"
                GoTo ContinueRow
            End If
            Dim errMsg As String
            If Not CanOpenWorkbookSheet(resolvedPath, sheetRef, checkSheetReadable, errMsg) Then
                AddIssue issues, r, "文件/Sheet不可读: " & errMsg
            End If
        End If

ContinueRow:
    Next r

    Set CollectConfigValidationIssues = issues
End Function

Private Sub ValidateConfigHeaders(ByVal ws As Worksheet, ByVal issues As Object)
    Dim c As Long
    For c = 1 To 17
        If Len(Trim$(ToText(ws.Cells(1, c).Value2))) = 0 Then
            AddIssue issues, 0, "Config第1行缺失表头: " & ColumnLabel(c) & "列。"
        End If
    Next c
End Sub

Private Sub ValidateColumnRule( _
    ByVal rawValue As Variant, _
    ByVal rowNo As Long, _
    ByVal colLabel As String, _
    ByVal fieldName As String, _
    ByVal requiredField As Boolean, _
    ByVal issues As Object)

    Dim s As String
    s = Trim$(ToText(rawValue))
    If Len(s) = 0 Then
        If requiredField Then AddIssue issues, rowNo, fieldName & "（" & colLabel & "列）缺失。"
        Exit Sub
    End If

    If ColToIndex(s) <= 0 Then
        AddIssue issues, rowNo, fieldName & "（" & colLabel & "列）非法[" & s & "]，应为列标字母或正整数。"
    End If
End Sub

Private Sub ValidatePlannedRule( _
    ByVal rawValue As Variant, _
    ByVal rowNo As Long, _
    ByVal moduleName As String, _
    ByVal issues As Object)

    Dim s As String
    s = Trim$(ToText(rawValue))
    If Len(s) = 0 Then Exit Sub

    If IsNumeric(s) Then Exit Sub
    If IsLettersOnly(s) Then Exit Sub

    AddIssue issues, rowNo, "计划规则（J列）非法[" & s & "]，应为列标字母或数字天数。"
End Sub

Private Sub ValidateDataStartRule( _
    ByVal rawValue As Variant, _
    ByVal rowNo As Long, _
    ByVal moduleName As String, _
    ByVal issues As Object)

    Dim s As String
    s = Trim$(ToText(rawValue))
    If Len(s) = 0 Then Exit Sub
    If Not IsNumeric(s) Then
        AddIssue issues, rowNo, "数据起始行（Q列）非法[" & s & "]。"
        Exit Sub
    End If
    If CLng(CDbl(s)) < 2 Then
        AddIssue issues, rowNo, "数据起始行（Q列）小于2。"
    End If
End Sub

Private Function CanOpenWorkbookSheet( _
    ByVal filePath As String, _
    ByVal sheetRef As Variant, _
    ByVal checkSheetReadable As Boolean, _
    ByRef errMsg As String) As Boolean

    Dim wb As Workbook
    Dim ws As Worksheet
    On Error GoTo HandleErr

    If Not OpenWorkbookReadOnly(filePath, wb, errMsg) Then
        CanOpenWorkbookSheet = False
        Exit Function
    End If

    If checkSheetReadable Then
        If IsNumeric(sheetRef) Then
            Set ws = wb.Worksheets(CLng(sheetRef))
        Else
            Set ws = wb.Worksheets(CStr(sheetRef))
        End If
    End If

    CloseWorkbookQuietly wb
    CanOpenWorkbookSheet = True
    Exit Function

HandleErr:
    If Len(errMsg) = 0 Then errMsg = Err.Description
    If InStr(1, errMsg, "Subscript out of range", vbTextCompare) > 0 Then
        errMsg = "Sheet不存在或索引越界: " & CStr(sheetRef)
    End If
    CloseWorkbookQuietly wb
    CanOpenWorkbookSheet = False
End Function

Private Sub AddIssue(ByVal issues As Object, ByVal rowNo As Long, ByVal message As String)
    Dim key As String
    key = CStr(rowNo)

    If issues.Exists(key) Then
        If InStr(1, CStr(issues(key)), message, vbTextCompare) = 0 Then
            issues(key) = CStr(issues(key)) & "; " & message
        End If
    Else
        issues(key) = message
    End If
End Sub

Private Sub AppendIssuesToCollection(ByVal issues As Object, ByVal target As Collection)
    Dim keys As Variant
    keys = SortedIssueKeys(issues)
    If IsEmpty(keys) Then Exit Sub

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim key As String
        key = CStr(keys(i))
        If key = "0" Then
            target.Add "[Config全局] " & CStr(issues(key))
        Else
            target.Add "[Config第" & key & "行] " & CStr(issues(key))
        End If
    Next i
End Sub

Private Function SortedIssueKeys(ByVal issues As Object) As Variant
    If issues Is Nothing Or issues.Count = 0 Then
        SortedIssueKeys = Empty
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(0 To issues.Count - 1)

    Dim i As Long
    Dim k As Variant
    i = 0
    For Each k In issues.Keys
        arr(i) = CStr(k)
        i = i + 1
    Next k

    Dim j As Long
    Dim temp As String
    For i = 1 To UBound(arr)
        temp = arr(i)
        j = i - 1
        Do While j >= 0 And CLng(arr(j)) > CLng(temp)
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = temp
    Next i

    SortedIssueKeys = arr
End Function

Private Function BuildPrecheckRows( _
    ByVal wsConfig As Worksheet, _
    ByVal issues As Object) As Collection

    Dim rows As Collection
    Set rows = New Collection

    Dim lastRow As Long
    lastRow = GetConfigLastRow(wsConfig)
    If lastRow < 2 Then
        Set BuildPrecheckRows = rows
        Exit Function
    End If

    Dim r As Long
    For r = 2 To lastRow
        Dim topic As String
        Dim moduleName As String
        Dim filePath As String
        Dim sheetRef As Variant
        topic = Trim$(ToText(wsConfig.Cells(r, 2).Value2))
        moduleName = Trim$(ToText(wsConfig.Cells(r, 3).Value2))
        filePath = Trim$(ToText(wsConfig.Cells(r, 5).Value2))
        sheetRef = NormalizeSheetName(wsConfig.Cells(r, 6).Value2)

        If Len(moduleName) = 0 And Len(filePath) = 0 Then GoTo ContinueRow

        Dim rec As Object
        Set rec = CreateObject("Scripting.Dictionary")
        rec("row_no") = r
        rec("topic") = topic
        rec("module") = moduleName
        rec("source_file") = ResolvePath(filePath)
        rec("source_sheet") = CStr(sheetRef)
        rec("config_check") = IIf(issues.Exists(CStr(r)), "失败", "通过")
        rec("readable_check") = "未扫描"
        rec("content_check") = "未扫描"
        rec("date_anomaly_count") = 0
        rec("date_anomaly_samples") = ""
        rec("notes") = ""

        If issues.Exists(CStr(r)) Then rec("notes") = CStr(issues(CStr(r)))

        Dim idCol As Long, contentCol As Long, initiatedCol As Long
        idCol = ColToIndex(wsConfig.Cells(r, 7).Value2)
        contentCol = ColToIndex(wsConfig.Cells(r, 8).Value2)
        initiatedCol = ColToIndex(wsConfig.Cells(r, 9).Value2)

        If Len(filePath) = 0 Or idCol <= 0 Or contentCol <= 0 Or initiatedCol <= 0 Then
            rows.Add rec
            GoTo ContinueRow
        End If

        Dim values As Variant
        Dim sheetNameResolved As String
        Dim errMsg As String
        If Not ReadSheetValues(CStr(rec("source_file")), sheetRef, values, sheetNameResolved, errMsg) Then
            rec("readable_check") = "不可读"
            rec("content_check") = "-"
            rec("notes") = AppendText(CStr(rec("notes")), "读取失败: " & errMsg)
            rows.Add rec
            GoTo ContinueRow
        End If

        rec("readable_check") = "可读"
        rec("source_sheet") = sheetNameResolved

        Dim rowCount As Long
        rowCount = UBound(values, 1)
        If rowCount <= 1 Then
            rec("content_check") = "空表/仅表头"
            rows.Add rec
            GoTo ContinueRow
        End If

        Dim localWarnings As Collection
        Set localWarnings = New Collection

        Dim startRow As Long
        startRow = ParseDataStartRow(wsConfig.Cells(r, 17).Value2, r, moduleName, localWarnings)
        If startRow > rowCount Then
            rec("content_check") = "数据起始行越界"
            rows.Add rec
            GoTo ContinueRow
        End If

        rec("content_check") = "正常"

        Dim plannedCol As Long, plannedDueDays As Long
        Dim hasPlannedCol As Boolean, hasPlannedDueDays As Boolean
        ParsePlannedRule wsConfig.Cells(r, 10).Value2, r, moduleName, localWarnings, plannedCol, plannedDueDays, hasPlannedCol, hasPlannedDueDays
        If localWarnings.Count > 0 Then
            rec("notes") = AppendText(CStr(rec("notes")), JoinCollection(localWarnings, "; "))
        End If

        Dim anomalyCount As Long
        Dim anomalySamples As String
        ScanDateAnomalies values, startRow, idCol, contentCol, initiatedCol, hasPlannedCol, plannedCol, anomalyCount, anomalySamples
        rec("date_anomaly_count") = anomalyCount
        rec("date_anomaly_samples") = anomalySamples
        If anomalyCount > 0 Then
            rec("content_check") = "存在日期异常"
        End If

        rows.Add rec

ContinueRow:
    Next r

    Set BuildPrecheckRows = rows
End Function

Private Sub ScanDateAnomalies( _
    ByVal values As Variant, _
    ByVal startRow As Long, _
    ByVal idCol As Long, _
    ByVal contentCol As Long, _
    ByVal initiatedCol As Long, _
    ByVal hasPlannedCol As Boolean, _
    ByVal plannedCol As Long, _
    ByRef anomalyCount As Long, _
    ByRef anomalySamples As String)

    Dim rowCount As Long
    rowCount = UBound(values, 1)

    Dim r As Long
    For r = startRow To rowCount
        Dim eventId As String, content As String, initiatedRaw As String
        eventId = GetCellText(values, r, idCol)
        content = GetCellText(values, r, contentCol)
        initiatedRaw = GetCellText(values, r, initiatedCol)

        If Len(eventId) = 0 And Len(content) = 0 And Len(initiatedRaw) = 0 Then GoTo ContinueRow

        Dim d As Date
        If Len(initiatedRaw) > 0 And Not TryParseDate(initiatedRaw, d) Then
            If Not IsHeaderLikeRow(eventId, content, initiatedRaw) Then
                anomalyCount = anomalyCount + 1
                anomalySamples = AppendDateSample(anomalySamples, r, "发起日期", initiatedRaw)
            End If
        End If

        If hasPlannedCol Then
            Dim plannedRaw As String
            plannedRaw = GetCellText(values, r, plannedCol)
            If Len(plannedRaw) > 0 And Not TryParseDate(plannedRaw, d) Then
                anomalyCount = anomalyCount + 1
                anomalySamples = AppendDateSample(anomalySamples, r, "计划日期", plannedRaw)
            End If
        End If

ContinueRow:
    Next r
End Sub

Private Function AppendDateSample(ByVal raw As String, ByVal rowNo As Long, ByVal fieldName As String, ByVal fieldValue As String) As String
    Dim sample As String
    sample = "R" & CStr(rowNo) & " " & fieldName & "=""" & fieldValue & """"

    If Len(raw) = 0 Then
        AppendDateSample = sample
    Else
        Dim existingCount As Long
        existingCount = UBound(Split(raw, " | ")) + 1
        If existingCount >= PRECHECK_MAX_DATE_SAMPLES Then
            AppendDateSample = raw
        Else
            AppendDateSample = raw & " | " & sample
        End If
    End If
End Function

Private Sub WritePrecheckSheet(ByVal ws As Worksheet, ByVal reportRows As Collection)
    ws.Cells.Clear

    Dim headers As Variant
    headers = Array("配置行", "主题", "质量模块", "文件路径", "Sheet", "配置校验", "文件/Sheet可读", "表内容检查", "日期异常数", "日期异常示例", "备注")

    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    ws.Rows(1).Font.Bold = True

    If reportRows.Count = 0 Then
        ws.Cells(2, 1).Value = "无可检查配置"
        ws.Columns("A:K").AutoFit
        Exit Sub
    End If

    Dim arr() As Variant
    ReDim arr(1 To reportRows.Count, 1 To 11)

    For i = 1 To reportRows.Count
        Dim rec As Object
        Set rec = reportRows(i)
        arr(i, 1) = rec("row_no")
        arr(i, 2) = rec("topic")
        arr(i, 3) = rec("module")
        arr(i, 4) = rec("source_file")
        arr(i, 5) = rec("source_sheet")
        arr(i, 6) = rec("config_check")
        arr(i, 7) = rec("readable_check")
        arr(i, 8) = rec("content_check")
        arr(i, 9) = rec("date_anomaly_count")
        arr(i, 10) = rec("date_anomaly_samples")
        arr(i, 11) = rec("notes")
    Next i

    ws.Range("A2").Resize(reportRows.Count, 11).Value = arr
    ws.Columns("A:K").AutoFit
End Sub

Private Function GetConfigLastRow(ByVal ws As Worksheet) As Long
    Dim rA As Long, rC As Long, rE As Long
    rA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    rC = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    rE = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    GetConfigLastRow = Application.WorksheetFunction.Max(2, rA, rC, rE)
End Function

Private Function AppendText(ByVal lhs As String, ByVal rhs As String) As String
    If Len(Trim$(rhs)) = 0 Then
        AppendText = lhs
    ElseIf Len(Trim$(lhs)) = 0 Then
        AppendText = rhs
    Else
        AppendText = lhs & "; " & rhs
    End If
End Function

Private Function ColumnLabel(ByVal colIndex As Long) As String
    Dim n As Long
    n = colIndex
    If n <= 0 Then
        ColumnLabel = "A"
        Exit Function
    End If

    Dim s As String
    Do While n > 0
        Dim r As Long
        r = (n - 1) Mod 26
        s = Chr$(Asc("A") + r) & s
        n = (n - 1) \ 26
    Loop
    ColumnLabel = s
End Function

Private Function ResolvePath(ByVal rawPath As String) As String
    Dim p As String
    p = Trim$(rawPath)
    If Len(p) = 0 Then
        ResolvePath = p
        Exit Function
    End If

    If IsAbsolutePath(p) Then
        ResolvePath = p
    Else
        If Len(ThisWorkbook.Path) = 0 Then
            ' Workbook not saved yet. Keep original relative path so caller can report a clear error.
            ResolvePath = p
        Else
            ResolvePath = ThisWorkbook.Path & Application.PathSeparator & p
        End If
    End If
End Function

Private Function IsAbsolutePath(ByVal p As String) As Boolean
    Dim s As String
    s = Trim$(p)
    If Len(s) = 0 Then Exit Function

    ' Windows drive path, e.g. C:\...
    If Len(s) >= 2 And Mid$(s, 2, 1) = ":" Then
        IsAbsolutePath = True
        Exit Function
    End If
    ' UNC path, e.g. \\server\share\...
    If Left$(s, 2) = "\\" Then
        IsAbsolutePath = True
        Exit Function
    End If
    ' Unix/macOS absolute path, e.g. /Users/...
    If Left$(s, 1) = "/" Then
        IsAbsolutePath = True
        Exit Function
    End If
End Function

Private Function OpenWorkbookReadOnly(ByVal filePath As String, ByRef wb As Workbook, ByRef errMsg As String) As Boolean
    Dim oldSecurity As Variant
    Dim hasAutomationSecurity As Boolean
    hasAutomationSecurity = TrySetAutomationSecurityForceDisable(oldSecurity)

    On Error GoTo OpenFallback
    Set wb = Application.Workbooks.Open( _
        Filename:=filePath, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, _
        AddToMru:=False)
    OpenWorkbookReadOnly = True
    GoTo FinallyExit

OpenFallback:
    Dim firstErr As String
    firstErr = Err.Description
    Err.Clear

    On Error GoTo OpenFail
    Set wb = Application.Workbooks.Open(filePath, 0, True)
    OpenWorkbookReadOnly = True
    GoTo FinallyExit

OpenFail:
    errMsg = BuildOpenErrorMessage(filePath, firstErr, Err.Description)
    OpenWorkbookReadOnly = False

FinallyExit:
    RestoreAutomationSecurity hasAutomationSecurity, oldSecurity
End Function

Private Function TrySetAutomationSecurityForceDisable(ByRef oldValue As Variant) As Boolean
    On Error Resume Next
    oldValue = Application.AutomationSecurity
    If Err.Number = 0 Then
        Application.AutomationSecurity = OPEN_WORKBOOK_SECURITY_FORCE_DISABLE
        TrySetAutomationSecurityForceDisable = (Err.Number = 0)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub RestoreAutomationSecurity(ByVal canRestore As Boolean, ByVal oldValue As Variant)
    If Not canRestore Then Exit Sub
    On Error Resume Next
    Application.AutomationSecurity = oldValue
    Err.Clear
    On Error GoTo 0
End Sub

Private Function BuildOpenErrorMessage(ByVal filePath As String, ByVal firstErr As String, ByVal finalErr As String) As String
    Dim msg As String
    msg = Trim$(finalErr)
    If Len(msg) = 0 Then msg = Trim$(firstErr)
    If Len(msg) = 0 Then msg = "未知错误"

    If Len(Dir$(filePath, vbNormal)) = 0 Then
        BuildOpenErrorMessage = "文件不存在或路径不可达: " & filePath
    Else
        BuildOpenErrorMessage = msg
    End If
End Function

Private Sub CloseWorkbookQuietly(ByRef wb As Workbook)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wb = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Function ReadSheetValues( _
    ByVal filePath As String, _
    ByVal sheetRef As Variant, _
    ByRef values As Variant, _
    ByRef sheetNameResolved As String, _
    ByRef errMsg As String) As Boolean

    Dim wb As Workbook
    Dim ws As Worksheet
    On Error GoTo HandleErr

    If Not OpenWorkbookReadOnly(filePath, wb, errMsg) Then
        ReadSheetValues = False
        Exit Function
    End If

    If IsNumeric(sheetRef) Then
        Set ws = wb.Worksheets(CLng(sheetRef))
    Else
        Set ws = wb.Worksheets(CStr(sheetRef))
    End If
    sheetNameResolved = ws.Name

    Dim lastRow As Long, lastCol As Long
    FindLastCell ws, lastRow, lastCol

    Dim raw As Variant
    raw = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2
    values = NormalizeTo2D(raw)

    CloseWorkbookQuietly wb
    ReadSheetValues = True
    Exit Function

HandleErr:
    If Len(errMsg) = 0 Then errMsg = Err.Description
    If InStr(1, errMsg, "Subscript out of range", vbTextCompare) > 0 Then
        errMsg = "Sheet不存在或索引越界: " & CStr(sheetRef)
    End If
    CloseWorkbookQuietly wb
    ReadSheetValues = False
End Function

Private Sub FindLastCell(ByVal ws As Worksheet, ByRef lastRow As Long, ByRef lastCol As Long)
    Dim cRow As Range
    Dim cCol As Range

    Set cRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    Set cCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

    If cRow Is Nothing Or cCol Is Nothing Then
        lastRow = 1
        lastCol = 1
    Else
        lastRow = cRow.Row
        lastCol = cCol.Column
    End If
End Sub

Private Function NormalizeTo2D(ByVal raw As Variant) As Variant
    If IsArray(raw) Then
        NormalizeTo2D = raw
    Else
        Dim arr(1 To 1, 1 To 1) As Variant
        arr(1, 1) = raw
        NormalizeTo2D = arr
    End If
End Function

Private Function GetCellText(ByVal values As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    If colIndex <= 0 Then Exit Function
    If rowIndex < LBound(values, 1) Or rowIndex > UBound(values, 1) Then Exit Function
    If colIndex < LBound(values, 2) Or colIndex > UBound(values, 2) Then Exit Function
    If IsError(values(rowIndex, colIndex)) Then Exit Function
    GetCellText = Trim$(ToText(values(rowIndex, colIndex)))
End Function

Private Function IsHeaderLikeRow(ByVal eventId As String, ByVal content As String, ByVal initiatedRaw As String) As Boolean
    Dim initiatedText As String
    initiatedText = Trim$(initiatedRaw)
    If Len(initiatedText) = 0 Then Exit Function

    Dim parts As Variant
    parts = Split(HEADER_HINTS, "|")

    Dim i As Long
    Dim headerHit As Boolean
    For i = LBound(parts) To UBound(parts)
        If InStr(1, initiatedText, CStr(parts(i)), vbTextCompare) > 0 Then
            headerHit = True
            Exit For
        End If
    Next i
    If Not headerHit Then Exit Function

    Dim eventIdLower As String
    Dim contentLower As String
    eventIdLower = LCase$(Trim$(eventId))
    contentLower = LCase$(Trim$(content))

    Dim idLike As Boolean
    idLike = (InStr(1, eventId, "编号", vbTextCompare) > 0) Or _
             (InStr(1, eventId, "序号", vbTextCompare) > 0) Or _
             (eventIdLower = "id") Or (eventIdLower = "编号") Or (eventIdLower = "序号")

    Dim contentLike As Boolean
    contentLike = (InStr(1, content, "内容", vbTextCompare) > 0) Or _
                  (contentLower = "content") Or (contentLower = "事项")

    IsHeaderLikeRow = (idLike Or contentLike)
End Function

Private Function ParseDataStartRow(ByVal rawValue As Variant, ByVal rowNo As Long, ByVal moduleName As String, ByVal warnings As Collection) As Long
    Dim s As String
    s = Trim$(ToText(rawValue))
    If Len(s) = 0 Then
        ParseDataStartRow = 2
        Exit Function
    End If

    If Not IsNumeric(s) Then
        warnings.Add "Config第" & CStr(rowNo) & "行 模块[" & moduleName & "]数据起始行非法[" & s & "]，已回退为2。"
        ParseDataStartRow = 2
        Exit Function
    End If

    Dim n As Long
    n = CLng(CDbl(s))
    If n < 2 Then
        warnings.Add "Config第" & CStr(rowNo) & "行 模块[" & moduleName & "]数据起始行[" & CStr(n) & "]小于2，已回退为2。"
        n = 2
    End If
    ParseDataStartRow = n
End Function

Private Sub ParsePlannedRule( _
    ByVal rawValue As Variant, _
    ByVal rowNo As Long, _
    ByVal moduleName As String, _
    ByVal warnings As Collection, _
    ByRef plannedCol As Long, _
    ByRef plannedDueDays As Long, _
    ByRef hasPlannedCol As Boolean, _
    ByRef hasPlannedDueDays As Boolean)

    Dim s As String
    s = Trim$(ToText(rawValue))
    If Len(s) = 0 Then Exit Sub

    If IsNumeric(s) Then
        plannedDueDays = CLng(CDbl(s))
        hasPlannedDueDays = True
        Exit Sub
    End If

    If IsLettersOnly(s) Then
        plannedCol = ColToIndex(s)
        If plannedCol > 0 Then
            hasPlannedCol = True
            Exit Sub
        End If
    End If

    warnings.Add "Config第" & CStr(rowNo) & "行 模块[" & moduleName & "]计划完成规则非法[" & s & "]，应为列标字母(如J/AA)或数字天数。"
End Sub

Private Function ParseYear(ByVal rawValue As Variant) As String
    Dim s As String
    s = Trim$(ToText(rawValue))
    If Len(s) = 0 Then
        ParseYear = "未知"
        Exit Function
    End If
    If IsNumeric(s) Then
        ParseYear = CStr(CLng(CDbl(s)))
    Else
        ParseYear = s
    End If
End Function

Private Function ColToIndex(ByVal rawValue As Variant) As Long
    Dim s As String
    s = Trim$(ToText(rawValue))
    If Len(s) = 0 Then Exit Function

    If IsNumeric(s) Then
        Dim n As Long
        n = CLng(CDbl(s))
        If n > 0 Then ColToIndex = n
        Exit Function
    End If

    Dim letters As String
    letters = UCase$(s)
    If Not IsLettersOnly(letters) Then Exit Function

    Dim i As Long, result As Long
    For i = 1 To Len(letters)
        result = result * 26 + (Asc(Mid$(letters, i, 1)) - Asc("A") + 1)
    Next i
    ColToIndex = result
End Function

Private Function ColToIndexOptional(ByVal rawValue As Variant) As Long
    ColToIndexOptional = ColToIndex(rawValue)
End Function

Private Function NormalizeSheetName(ByVal rawValue As Variant) As Variant
    Dim s As String
    s = Trim$(ToText(rawValue))
    If Len(s) = 0 Then
        NormalizeSheetName = 1
        Exit Function
    End If
    If IsDigitsOnly(s) Then
        NormalizeSheetName = CLng(s)
    Else
        NormalizeSheetName = s
    End If
End Function

Private Function TryParseDate(ByVal rawValue As Variant, ByRef outDate As Date) As Boolean
    Dim s As String
    s = Trim$(ToText(rawValue))
    If Len(s) = 0 Then Exit Function

    If IsNumeric(s) Then
        Dim serial As Double
        serial = CDbl(s)
        If serial > 0 Then
            outDate = DateSerial(1899, 12, 30) + serial
            TryParseDate = True
            Exit Function
        End If
    End If

    Dim normalized As String
    normalized = s
    normalized = Replace(normalized, "年", "-")
    normalized = Replace(normalized, "月", "-")
    normalized = Replace(normalized, "日", "")
    normalized = Replace(normalized, "/", "-")
    normalized = Replace(normalized, ".", "-")
    normalized = CollapseSpaces(normalized)

    If TryParseYmd(normalized, outDate) Then
        TryParseDate = True
        Exit Function
    End If
    If TryParseYm(normalized, outDate) Then
        TryParseDate = True
        Exit Function
    End If

    On Error Resume Next
    outDate = CDate(normalized)
    TryParseDate = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function TryParseYmd(ByVal textValue As String, ByRef outDate As Date) As Boolean
    Dim y As Long
    Dim mo As Long
    Dim d As Long

    If Not ExtractFirstYmd(textValue, y, mo, d) Then Exit Function
    If Not IsDatePartValid(y, mo, d) Then Exit Function

    outDate = DateSerial(y, mo, d)
    TryParseYmd = True
End Function

Private Function TryParseYm(ByVal textValue As String, ByRef outDate As Date) As Boolean
    Dim s As String
    s = Trim$(textValue)
    If Len(s) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(s, "-")
    If UBound(parts) <> 1 Then Exit Function

    If Not IsDigitsOnly(parts(0)) Then Exit Function
    If Not IsDigitsOnly(parts(1)) Then Exit Function
    If Len(parts(0)) <> 4 Then Exit Function

    Dim y As Long
    Dim mo As Long
    y = CLng(parts(0))
    mo = CLng(parts(1))
    If mo < 1 Or mo > 12 Then Exit Function

    outDate = DateSerial(y, mo, 1)
    TryParseYm = True
End Function

Private Function IsDatePartValid(ByVal y As Long, ByVal mo As Long, ByVal d As Long) As Boolean
    If mo < 1 Or mo > 12 Then Exit Function
    If d < 1 Then Exit Function
    Dim maxDay As Long
    maxDay = Day(DateSerial(y, mo + 1, 0))
    IsDatePartValid = (d <= maxDay)
End Function

Private Function ExtractFirstYmd(ByVal textValue As String, ByRef y As Long, ByRef mo As Long, ByRef d As Long) As Boolean
    Dim s As String
    s = Trim$(textValue)
    If Len(s) < 8 Then Exit Function

    Dim i As Long
    For i = 1 To Len(s) - 7
        If IsDigitsOnly(Mid$(s, i, 4)) Then
            If Mid$(s, i + 4, 1) = "-" Then
                Dim monthStart As Long
                monthStart = i + 5

                Dim ml As Long
                For ml = 1 To 2
                    Dim monthEnd As Long
                    monthEnd = monthStart + ml - 1
                    If monthEnd > Len(s) Then Exit For

                    Dim monthPart As String
                    monthPart = Mid$(s, monthStart, ml)
                    If Not IsDigitsOnly(monthPart) Then GoTo NextMonthLen

                    If monthEnd + 1 > Len(s) Then GoTo NextMonthLen
                    If Mid$(s, monthEnd + 1, 1) <> "-" Then GoTo NextMonthLen

                    Dim dayStart As Long
                    dayStart = monthEnd + 2

                    Dim dl As Long
                    For dl = 1 To 2
                        Dim dayEnd As Long
                        dayEnd = dayStart + dl - 1
                        If dayEnd > Len(s) Then Exit For

                        Dim dayPart As String
                        dayPart = Mid$(s, dayStart, dl)
                        If Not IsDigitsOnly(dayPart) Then GoTo NextDayLen

                        Dim tailPos As Long
                        tailPos = dayEnd + 1
                        If tailPos <= Len(s) Then
                            Dim tailChar As String
                            tailChar = Mid$(s, tailPos, 1)
                            If tailChar >= "0" And tailChar <= "9" Then GoTo NextDayLen
                        End If

                        y = CLng(Mid$(s, i, 4))
                        mo = CLng(monthPart)
                        d = CLng(dayPart)
                        ExtractFirstYmd = True
                        Exit Function
NextDayLen:
                    Next dl
NextMonthLen:
                Next ml
            End If
        End If
    Next i
End Function

Private Function AddOneMonthCompat(ByVal d As Date) As Date
    Dim y As Long, mo As Long, dayNo As Long
    y = Year(d)
    mo = Month(d)
    dayNo = Day(d)

    If mo = 12 Then
        y = y + 1
        mo = 1
    Else
        mo = mo + 1
    End If

    Dim maxDay As Long
    maxDay = Day(DateSerial(y, mo + 1, 0))
    If dayNo > maxDay Then dayNo = maxDay
    AddOneMonthCompat = DateSerial(y, mo, dayNo)
End Function

Private Function SortedKeysAsc(ByVal d As Object) As Variant
    If d Is Nothing Or d.Count = 0 Then
        SortedKeysAsc = Empty
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(0 To d.Count - 1)

    Dim i As Long
    Dim k As Variant
    i = 0
    For Each k In d.Keys
        arr(i) = CStr(k)
        i = i + 1
    Next k

    Dim j As Long
    Dim temp As String
    For i = 1 To UBound(arr)
        temp = arr(i)
        j = i - 1
        Do While j >= 0 And StrComp(arr(j), temp, vbTextCompare) > 0
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = temp
    Next i

    SortedKeysAsc = arr
End Function

Private Function SortedPairsByCountDesc(ByVal d As Object) As Variant
    If d Is Nothing Or d.Count = 0 Then
        SortedPairsByCountDesc = Empty
        Exit Function
    End If

    Dim arr() As Variant
    ReDim arr(1 To d.Count, 1 To 2)

    Dim i As Long
    Dim k As Variant
    i = 1
    For Each k In d.Keys
        arr(i, 1) = CStr(k)
        arr(i, 2) = CLng(d(k))
        i = i + 1
    Next k

    Dim j As Long
    Dim keyName As String
    Dim keyCount As Long
    For i = 2 To UBound(arr, 1)
        keyName = CStr(arr(i, 1))
        keyCount = CLng(arr(i, 2))
        j = i - 1
        Do While j >= 1 And ShouldMoveDown(CStr(arr(j, 1)), CLng(arr(j, 2)), keyName, keyCount)
            arr(j + 1, 1) = arr(j, 1)
            arr(j + 1, 2) = arr(j, 2)
            j = j - 1
        Loop
        arr(j + 1, 1) = keyName
        arr(j + 1, 2) = keyCount
    Next i

    SortedPairsByCountDesc = arr
End Function

Private Function ShouldMoveDown(ByVal existingName As String, ByVal existingCount As Long, ByVal newName As String, ByVal newCount As Long) As Boolean
    If existingCount < newCount Then
        ShouldMoveDown = True
    ElseIf existingCount > newCount Then
        ShouldMoveDown = False
    Else
        ShouldMoveDown = (StrComp(existingName, newName, vbTextCompare) > 0)
    End If
End Function

Private Function PercentText(ByVal num As Long, ByVal den As Long) As String
    If den <= 0 Then
        PercentText = "0%"
    Else
        PercentText = Format$(Round((CDbl(num) * 100#) / CDbl(den), 2), "0.00") & "%"
    End If
End Function

Private Function ToText(ByVal v As Variant) As String
    If IsError(v) Then
        ToText = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        ToText = ""
    Else
        ToText = CStr(v)
    End If
End Function

Private Function CollapseSpaces(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    CollapseSpaces = t
End Function

Private Function IsDigitsOnly(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function

    Dim i As Long
    For i = 1 To Len(t)
        Dim ch As String
        ch = Mid$(t, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsDigitsOnly = True
End Function

Private Function IsLettersOnly(ByVal s As String) As Boolean
    Dim t As String
    t = UCase$(Trim$(s))
    If Len(t) = 0 Then Exit Function

    Dim i As Long
    For i = 1 To Len(t)
        Dim ch As String
        ch = Mid$(t, i, 1)
        If ch < "A" Or ch > "Z" Then Exit Function
    Next i
    IsLettersOnly = True
End Function

Private Function JoinCollection(ByVal c As Collection, ByVal delimiter As String) As String
    Dim i As Long
    For i = 1 To c.Count
        If i > 1 Then JoinCollection = JoinCollection & delimiter
        JoinCollection = JoinCollection & CStr(c(i))
    Next i
End Function

Private Function PreviewCollection(ByVal c As Collection, ByVal maxItems As Long) As String
    If c Is Nothing Or c.Count = 0 Then
        PreviewCollection = "无"
        Exit Function
    End If

    Dim i As Long
    Dim limit As Long
    limit = maxItems
    If limit < 1 Then limit = 1
    If limit > c.Count Then limit = c.Count

    For i = 1 To limit
        If i > 1 Then PreviewCollection = PreviewCollection & " | "
        PreviewCollection = PreviewCollection & CStr(c(i))
    Next i

    If c.Count > limit Then
        PreviewCollection = PreviewCollection & " ...共" & CStr(c.Count) & "项"
    End If
End Function
