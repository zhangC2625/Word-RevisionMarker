Attribute VB_Name = "RevisionMarker"
Option Explicit

'============================================================================
' 诊断宏：用 NextRevision 导航，检查文档修订
'============================================================================
Sub DiagnoseRevisions()
    Dim insCount As Long, delCount As Long, otherCount As Long
    Dim rev As Revision
    Dim rvType As Long
    Dim msg As String

    insCount = 0: delCount = 0: otherCount = 0

    ' 回到文档开头
    Selection.HomeKey Unit:=wdStory

    Do
        On Error Resume Next
        Selection.NextRevision Wrap:=wdFindStop
        On Error GoTo 0

        If Selection.Range.Revisions.Count = 0 Then Exit Do

        rvType = Selection.Range.Revisions(1).Type

        Select Case rvType
            Case 1:  ' wdRevisionInsert
                insCount = insCount + 1
            Case 2:  ' wdRevisionDelete
                delCount = delCount + 1
            Case Else
                otherCount = otherCount + 1
        End Select

        ' 移到当前修订末尾，避免重复找到同一条
        Selection.Collapse Direction:=wdCollapseEnd
    Loop

    msg = "修订诊断结果：" & vbCrLf & vbCrLf & _
          "插入: " & insCount & " 条" & vbCrLf & _
          "删除: " & delCount & " 条" & vbCrLf & _
          "其他: " & otherCount & " 条"

    MsgBox msg, vbInformation, "诊断结果"
End Sub


'============================================================================
' 主宏：逐条审阅修订标记，转为可视格式
' 使用 NextRevision 导航，不碰 Revisions 集合
'============================================================================
Sub ReviewAndMarkRevisions_Simple()
    Dim rev As Revision
    Dim rvType As Long
    Dim revText As String
    Dim userInput As String
    Dim rng As Range
    Dim insRng As Range
    Dim delCount As Long, insCount As Long
    Dim processedDel As Long, processedIns As Long
    Dim markedDel As Long, markedIns As Long
    Dim autoMarkAll As Boolean

    ' ── 第一遍：统计 ──
    delCount = 0: insCount = 0
    Selection.HomeKey Unit:=wdStory
    Do
        On Error Resume Next
        Selection.NextRevision Wrap:=wdFindStop
        On Error GoTo 0
        If Selection.Range.Revisions.Count = 0 Then Exit Do

        rvType = Selection.Range.Revisions(1).Type
        If rvType = 1 Then insCount = insCount + 1
        If rvType = 2 Then delCount = delCount + 1

        Selection.Collapse Direction:=wdCollapseEnd
    Loop

    If delCount + insCount = 0 Then
        MsgBox "文档中没有找到插入或删除类型的修订标记。", vbInformation, "提示"
        Exit Sub
    End If

    autoMarkAll = False
    processedDel = 0: processedIns = 0
    markedDel = 0: markedIns = 0

    Application.ScreenUpdating = False

    ' ══════════════════════════════════════════
    ' 第一阶段：逐条审阅"删除"的文字
    ' ══════════════════════════════════════════
    If delCount > 0 Then
        Do
            Selection.HomeKey Unit:=wdStory

            Do
                On Error Resume Next
                Selection.NextRevision Wrap:=wdFindStop
                On Error GoTo 0
                If Selection.Range.Revisions.Count = 0 Then Exit Do

                rvType = Selection.Range.Revisions(1).Type

                ' 跳过非删除类型
                If rvType <> 2 Then
                    Selection.Collapse Direction:=wdCollapseEnd
                    GoTo ContinueDeleteLoop
                End If

                processedDel = processedDel + 1
                Set rev = Selection.Range.Revisions(1)
                revText = Trim(rev.Range.Text)

                If autoMarkAll Then
                    Set rng = rev.Range.Duplicate
                    rev.Reject
                    rng.Font.ColorIndex = wdRed
                    rng.Font.Strikethrough = True
                    markedDel = markedDel + 1
                    GoTo ContinueDeleteLoop
                End If

                userInput = InputBox( _
                    "━━━ 【删除的文字】(" & processedDel & "/" & delCount & ") ━━━" & vbCrLf & _
                    vbCrLf & _
                    """" & Left(revText, 200) & """" & vbCrLf & _
                    vbCrLf & _
                    "标记效果：红色 + 删除线" & vbCrLf & _
                    vbCrLf & _
                    "已标记: " & markedDel & " 条" & vbCrLf & _
                    vbCrLf & _
                    "输入数字选择：" & vbCrLf & _
                    "  1 → 标记（红色+删除线）" & vbCrLf & _
                    "  2 → 跳过这一条" & vbCrLf & _
                    "  3 → 全部标记（不再询问）" & vbCrLf & _
                    "  0 → 停止退出", _
                    "修订审核 — 删除", "1")

                If StrPtr(userInput) = 0 Then GoTo Cleanup

                Select Case Trim(userInput)
                    Case "1"
                        Set rng = rev.Range.Duplicate
                        rev.Reject
                        rng.Font.ColorIndex = wdRed
                        rng.Font.Strikethrough = True
                        markedDel = markedDel + 1

                    Case "2"
                        rev.Reject   ' 移除修订标记，不加格式

                    Case "3"
                        autoMarkAll = True
                        Set rng = rev.Range.Duplicate
                        rev.Reject
                        rng.Font.ColorIndex = wdRed
                        rng.Font.Strikethrough = True
                        markedDel = markedDel + 1

                    Case "0"
                        GoTo Cleanup

                    Case Else
                        MsgBox "请输入 0、1、2 或 3。", vbExclamation
                        processedDel = processedDel - 1
                End Select

ContinueDeleteLoop:
            Loop

            ' 只有当所有删除都处理完才退出
            If processedDel >= delCount Then Exit Do
        Loop
    End If

    ' ══════════════════════════════════════════
    ' 第二阶段：逐条审阅"插入"的文字
    ' ══════════════════════════════════════════
    If insCount > 0 Then
        Do
            Selection.HomeKey Unit:=wdStory

            Do
                On Error Resume Next
                Selection.NextRevision Wrap:=wdFindStop
                On Error GoTo 0
                If Selection.Range.Revisions.Count = 0 Then Exit Do

                rvType = Selection.Range.Revisions(1).Type

                ' 跳过非插入类型
                If rvType <> 1 Then
                    Selection.Collapse Direction:=wdCollapseEnd
                    GoTo ContinueInsertLoop
                End If

                processedIns = processedIns + 1
                Set rev = Selection.Range.Revisions(1)
                revText = Trim(rev.Range.Text)

                If autoMarkAll Then
                    Set insRng = rev.Range.Duplicate
                    rev.Accept
                    insRng.Font.ColorIndex = wdBlue
                    insRng.Font.Underline = wdUnderlineSingle
                    markedIns = markedIns + 1
                    GoTo ContinueInsertLoop
                End If

                userInput = InputBox( _
                    "━━━ 【插入的文字】(" & processedIns & "/" & insCount & ") ━━━" & vbCrLf & _
                    vbCrLf & _
                    """" & Left(revText, 200) & """" & vbCrLf & _
                    vbCrLf & _
                    "标记效果：蓝色 + 下划线" & vbCrLf & _
                    vbCrLf & _
                    "已标记: " & markedIns & " 条" & vbCrLf & _
                    vbCrLf & _
                    "输入数字选择：" & vbCrLf & _
                    "  1 → 标记（蓝色+下划线）" & vbCrLf & _
                    "  2 → 跳过这一条" & vbCrLf & _
                    "  3 → 全部标记（不再询问）" & vbCrLf & _
                    "  0 → 停止退出", _
                    "修订审核 — 插入", "1")

                If StrPtr(userInput) = 0 Then GoTo Cleanup

                Select Case Trim(userInput)
                    Case "1"
                        Set insRng = rev.Range.Duplicate
                        rev.Accept
                        insRng.Font.ColorIndex = wdBlue
                        insRng.Font.Underline = wdUnderlineSingle
                        markedIns = markedIns + 1

                    Case "2"
                        rev.Accept   ' 接受但不加格式

                    Case "3"
                        autoMarkAll = True
                        Set insRng = rev.Range.Duplicate
                        rev.Accept
                        insRng.Font.ColorIndex = wdBlue
                        insRng.Font.Underline = wdUnderlineSingle
                        markedIns = markedIns + 1

                    Case "0"
                        GoTo Cleanup

                    Case Else
                        MsgBox "请输入 0、1、2 或 3。", vbExclamation
                        processedIns = processedIns - 1
                End Select

ContinueInsertLoop:
            Loop

            If processedIns >= insCount Then Exit Do
        Loop
    End If

Cleanup:
    Application.ScreenUpdating = True

    MsgBox "审阅结束！" & vbCrLf & vbCrLf & _
           "删除: 审阅 " & processedDel & "/" & delCount & " 条，标记 " & markedDel & " 条" & vbCrLf & _
           "插入: 审阅 " & processedIns & "/" & insCount & " 条，标记 " & markedIns & " 条", _
           vbInformation, "审阅完成"
End Sub


'============================================================================
' 清除由本工具添加的蓝色/红色格式
'============================================================================
Sub ClearRevisionFormatting()
    Dim result As VbMsgBoxResult

    result = MsgBox("此操作将清除所有 蓝色下划线 和 红色删除线 格式。" & vbCrLf & _
                    "此操作不可撤销（修订标记已被接受）。" & vbCrLf & vbCrLf & _
                    "是否继续？", vbYesNo + vbExclamation, "警告")

    If result = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    With ActiveDocument.Content.Find
        .ClearFormatting
        .Font.ColorIndex = wdBlue
        .Font.Underline = wdUnderlineSingle
        With .Replacement
            .ClearFormatting
            .Font.ColorIndex = wdAuto
            .Font.Underline = wdUnderlineNone
        End With
        .Execute Replace:=wdReplaceAll
    End With

    With ActiveDocument.Content.Find
        .ClearFormatting
        .Font.ColorIndex = wdRed
        .Font.Strikethrough = True
        With .Replacement
            .ClearFormatting
            .Font.ColorIndex = wdAuto
            .Font.Strikethrough = False
        End With
        .Execute Replace:=wdReplaceAll
    End With

    Application.ScreenUpdating = True
    MsgBox "格式已清除。", vbInformation, "完成"
End Sub


'============================================================================
' 帮助
'============================================================================
Sub ShowHelp()
    MsgBox "修订标记转换工具" & vbCrLf & vbCrLf & _
           "1. DiagnoseRevisions" & vbCrLf & _
           "   诊断文档中修订的类型和数量" & vbCrLf & vbCrLf & _
           "2. ReviewAndMarkRevisions_Simple" & vbCrLf & _
           "   逐条审阅，输入数字选择：" & vbCrLf & _
           "   1=标记  2=跳过  3=全部标记  0=停止" & vbCrLf & vbCrLf & _
           "效果：" & vbCrLf & _
           "  删除文字 → 红色 + 删除线" & vbCrLf & _
           "  插入文字 → 蓝色 + 下划线", _
           vbInformation, "帮助"
End Sub
