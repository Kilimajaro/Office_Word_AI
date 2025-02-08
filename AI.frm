Private Sub UserForm_Initialize()
    ' 初始化对话框标题
    Me.Caption = "Word AI助手"
    TextBox5.text = "qwen-max-0125"
End Sub
Private Sub Button1_Click()
    Dim selectedText As String
    Dim systemContent As String
    Dim http As Object
    Dim jsonBody As String
    Dim response As String
    Dim translation As String
    Dim regex As Object
    Dim matches As Object
    Dim new_content As String
    Dim dialogHistory As String

    ' 获取 Word 文档中选中的内容
    If Selection.Type = wdSelectionIP Then
        MsgBox "请先选中要咨询的文本。", vbExclamation
        Exit Sub
    End If
    selectedText = Selection.text
    selectedModel = TextBox5.text

    ' 获取用户输入的系统提示
    systemContent = TextBox1.text
    systemContent = Replace(systemContent, "\", "\\")       ' 转义反斜杠
    systemContent = Replace(systemContent, """", """""")    ' 转义双引号
    systemContent = Replace(systemContent, vbLf, "\n")      ' 转义换行符（LF）
    systemContent = Replace(systemContent, vbCr, "\r")      ' 转义回车符（CR）

    If selectedText = "" Then
        MsgBox "未检测到选中的文本，请先选中要咨询的内容。", vbExclamation
        Exit Sub
    End If

    ' 处理选中文本中的特殊字符，避免破坏 JSON 格式
    selectedText = Replace(selectedText, "\", "\\")        ' 转义反斜杠
    selectedText = Replace(selectedText, """", """""")     ' 转义双引号
    selectedText = Replace(selectedText, vbLf, "\n")       ' 转义换行符（LF）
    selectedText = Replace(selectedText, vbCr, "\r")       ' 转义回车符（CR）

    ' 创建HTTP对象
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    If http Is Nothing Then
        MsgBox "无法创建 HTTP 对象，请检查您的环境设置！", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' 构造JSON请求体，包含历史消息和当前消息
    dialogHistory = TextBox3.text
    dialogHistory = Replace(dialogHistory, "\", "\\")       ' 转义反斜杠
    dialogHistory = Replace(dialogHistory, """", """""")    ' 转义双引号
    dialogHistory = Replace(dialogHistory, vbLf, "\n")      ' 转义换行符（LF）
    dialogHistory = Replace(dialogHistory, vbCr, "\r")      ' 转义回车符（CR）
    new_content = dialogHistory + "以上是历史消息，以下是最新命令：" + systemContent
    jsonBody = "{""model"": """ & selectedModel & """,""messages"": [{""role"": ""system"",""content"": """ & new_content & """},{""role"": ""user"",""content"": """ & selectedText & """}]}"
 
    ' 发送POST请求
    ' 设置API密钥和URL
    Dim apiKey As String
    Dim apiUrl As String
    apiKey = "sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" ' 替换为您的API密钥
    apiUrl = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions" ' 建议使用阿里云百炼平台的请求
    With http
        .Open "POST", apiUrl, False
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .setRequestHeader "Content-Type", "application/json"
        .Send jsonBody
        If .readyState <> 4 Or .Status <> 200 Then
            MsgBox "请求失败: " & .Status & " - " & .responseText, vbCritical
            Exit Sub
        End If
        response = .responseText
    End With

    ' 使用正则表达式提取翻译结果
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = """content"":""(.*?)""}"
    regex.IgnoreCase = True
    regex.Global = True

    Set matches = regex.Execute(response)
    If matches.Count > 0 Then
        translation = matches(0).SubMatches(0)  ' 提取第一个匹配的内容

        ' 处理 JSON 数据中的转义字符
        translation = Replace(translation, "\""", """")   ' 替换转义的双引号
        translation = Replace(translation, "", """")      ' 替换单引号（如果需要）
        translation = Replace(translation, "\n", vbCrLf)  ' 将 \n 替换为换行符
        translation = Replace(translation, "\r", vbCrLf)  ' 将 \r 替换为换行符
        translation = Replace(translation, "\\", "\")      ' 取消反斜杠的转义

        ' 追加消息到TextBox3，并自动换行
        If Len(TextBox3.text) > 0 Then
            TextBox3.text = TextBox3.text & vbCrLf
        End If
        TextBox3.text = TextBox3.text & "用户: " & systemContent & vbCrLf & "回复: " & translation & vbCrLf & vbCrLf
    Else
        MsgBox "未能提取回复结果，请检查返回的数据格式。", vbCritical
    End If
End Sub
Private Sub Button2_Click()
    Dim insertText As String
    Dim lines() As String
    Dim i As Integer
    Dim markdownRegex As Object
    Dim regexPattern As String
    Dim matches As Object
    Dim mtch As Object
    Dim lineRange As range
    Dim insertPoint As Long ' 用于保存插入点的位置

    ' 获取用户在 TextBox3 中选中的文本
    insertText = TextBox3.SelText
    If insertText = "" Then
        MsgBox "请在 TextBox3 中选中要插入的文本。", vbExclamation
        Exit Sub
    End If

    ' 检查是否有选中的内容
    If Selection.Type = wdSelectionIP Then
        MsgBox "请先在 Word 文档中选中插入位置。", vbExclamation
        Exit Sub
    End If

    ' 按行分割文本
    lines = Split(insertText, vbCrLf)

    ' 初始化正则表达式对象
    Set markdownRegex = CreateObject("VBScript.RegExp")
    With markdownRegex
        .Global = True
        .IgnoreCase = True
    End With

    ' 获取当前选中范围的结束位置
    insertPoint = Selection.range.End

    ' 插入文本并处理 Markdown 格式
    For i = LBound(lines) To UBound(lines)
        Dim formattedLine As String
        formattedLine = lines(i)

        ' 处理加粗和斜体（***...***）
        regexPattern = "\*\*\*(.*?)\*\*\*"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(formattedLine)
        For Each mtch In matches
            formattedLine = Replace(formattedLine, mtch.value, mtch.SubMatches(0))
        Next mtch

        ' 处理加粗（**...**）
        regexPattern = "\*\*(.*?)\*\*"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(formattedLine)
        For Each mtch In matches
            formattedLine = Replace(formattedLine, mtch.value, mtch.SubMatches(0))
        Next mtch

        ' 处理斜体（*...*）
        regexPattern = "\*(.*?)\*"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(formattedLine)
        For Each mtch In matches
            formattedLine = Replace(formattedLine, mtch.value, mtch.SubMatches(0))
        Next mtch

        ' 处理标题
        regexPattern = "(#{1,4})\s(.*?(?=##|\*\*|\n|$))"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(formattedLine)
        For Each mtch In matches
            formattedLine = Replace(formattedLine, mtch.value, mtch.SubMatches(1))
        Next mtch

        ' 插入处理后的文本
        Set lineRange = ActiveDocument.range(insertPoint, insertPoint)
        lineRange.text = formattedLine & vbCrLf

        ' 格式化加粗和斜体（***...***）
        regexPattern = "\*\*\*(.*?)\*\*\*"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(lines(i))
        For Each mtch In matches
            Dim boldItalicRange As range
            Set boldItalicRange = lineRange.Duplicate
            boldItalicRange.Start = lineRange.Start + InStr(formattedLine, mtch.SubMatches(0))
            boldItalicRange.End = boldItalicRange.Start + Len(mtch.SubMatches(0))
            With boldItalicRange.Font
                .Bold = True
                .Italic = True
            End With
        Next mtch

        ' 格式化加粗（**...**）
        regexPattern = "\*\*(.*?)\*\*"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(lines(i))
        For Each mtch In matches
            Dim boldRange As range
            Set boldRange = lineRange.Duplicate
            boldRange.Start = lineRange.Start + InStr(formattedLine, mtch.SubMatches(0)) - 1
            boldRange.End = boldRange.Start + Len(mtch.SubMatches(0))
            boldRange.Font.Bold = True
        Next mtch

        ' 格式化斜体（*...*）
        regexPattern = "\*(.*?)\*"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(lines(i))
        For Each mtch In matches
            Dim italicRange As range
            Set italicRange = lineRange.Duplicate
            italicRange.Start = lineRange.Start + InStr(formattedLine, mtch.SubMatches(0)) - 1
            italicRange.End = italicRange.Start + Len(mtch.SubMatches(0))
            italicRange.Font.Italic = True
        Next mtch

        ' 格式化标题
        regexPattern = "(#{1,4})\s(.*?(?=##|\*\*|\n|$))"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(lines(i))
        For Each mtch In matches
            Dim headerRange As range
            Set headerRange = lineRange.Duplicate
            headerRange.Start = lineRange.Start + InStr(formattedLine, mtch.SubMatches(1)) - 1
            headerRange.End = headerRange.Start + Len(mtch.SubMatches(1))
            headerRange.Font.Bold = True
        Next mtch

        ' 更新插入点位置
        insertPoint = insertPoint + Len(formattedLine) + 1
    Next i
End Sub
Private Sub Button3_Click()
    Dim http As Object
    Dim url As String
    Dim json As String
    Dim response As String

    ' 设置 URL
    url = "http://127.0.0.1:8020/query"
    
    ' 设置请求体（JSON 格式）
    json = "{""query"": ""什么是先诉抗辩权？"", ""mode"": ""hybrid""}"
    
    ' 创建 XMLHTTP 对象
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' 打开 POST 请求
    http.Open "POST", url, False
    
    ' 设置请求头
    http.setRequestHeader "Content-Type", "application/json"
    
    ' 发送请求
    http.Send json
    
    ' 获取响应结果
    response = http.responseText
    
    ' 将响应结果写入 TextBox4
    TextBox4.value = response
    
    ' 释放对象
    Set http = Nothing
End Sub
