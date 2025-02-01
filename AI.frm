VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AI 
   Caption         =   "UserForm1"
   ClientHeight    =   7056
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   10800
   OleObjectBlob   =   "AI.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "AI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' ��ʼ���Ի������
    Me.Caption = "Word AI����"
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

    ' ��ȡ Word �ĵ���ѡ�е�����
    If Selection.Type = wdSelectionIP Then
        MsgBox "����ѡ��Ҫ��ѯ���ı���", vbExclamation
        Exit Sub
    End If
    selectedText = Selection.text

    ' ��ȡ�û������ϵͳ��ʾ
    systemContent = TextBox1.text
    systemContent = Replace(systemContent, "\", "\\")       ' ת�巴б��
    systemContent = Replace(systemContent, """", """""")    ' ת��˫����
    systemContent = Replace(systemContent, vbLf, "\n")      ' ת�廻�з���LF��
    systemContent = Replace(systemContent, vbCr, "\r")      ' ת��س�����CR��

    If selectedText = "" Then
        MsgBox "δ��⵽ѡ�е��ı�������ѡ��Ҫ��ѯ�����ݡ�", vbExclamation
        Exit Sub
    End If

    ' ����ѡ���ı��е������ַ��������ƻ� JSON ��ʽ
    selectedText = Replace(selectedText, "\", "\\")        ' ת�巴б��
    selectedText = Replace(selectedText, """", """""")     ' ת��˫����
    selectedText = Replace(selectedText, vbLf, "\n")       ' ת�廻�з���LF��
    selectedText = Replace(selectedText, vbCr, "\r")       ' ת��س�����CR��

    ' ����HTTP����
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    If http Is Nothing Then
        MsgBox "�޷����� HTTP �����������Ļ������ã�", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' ����JSON�����壬������ʷ��Ϣ�͵�ǰ��Ϣ
    dialogHistory = TextBox3.text
    dialogHistory = Replace(dialogHistory, "\", "\\")       ' ת�巴б��
    dialogHistory = Replace(dialogHistory, """", """""")    ' ת��˫����
    dialogHistory = Replace(dialogHistory, vbLf, "\n")      ' ת�廻�з���LF��
    dialogHistory = Replace(dialogHistory, vbCr, "\r")      ' ת��س�����CR��
    new_content = dialogHistory + "��������ʷ��Ϣ���������������" + systemContent
    jsonBody = "{""model"": ""qwen-plus"",""messages"": [{""role"": ""system"",""content"": """ & new_content & """},{""role"": ""user"",""content"": """ & selectedText & """}]}"

    ' ����POST����
    ' ����API��Կ��URL
    Dim apiKey As String
    Dim apiUrl As String
    apiKey = "sk-ad3f81c2a0934b3289501e9d4e3d6452" ' �滻Ϊ����API��Կ
    apiUrl = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    With http
        .Open "POST", apiUrl, False
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .setRequestHeader "Content-Type", "application/json"
        .send jsonBody
        If .readyState <> 4 Or .Status <> 200 Then
            MsgBox "����ʧ��: " & .Status & " - " & .responseText, vbCritical
            Exit Sub
        End If
        response = .responseText
    End With

    ' ʹ��������ʽ��ȡ������
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = """content"":""(.*?)""}"
    regex.IgnoreCase = True
    regex.Global = True

    Set matches = regex.Execute(response)
    If matches.Count > 0 Then
        translation = matches(0).SubMatches(0)  ' ��ȡ��һ��ƥ�������

        ' ���� JSON �����е�ת���ַ�
        translation = Replace(translation, "\""", """")   ' �滻ת���˫����
        translation = Replace(translation, "", """")      ' �滻�����ţ������Ҫ��
        translation = Replace(translation, "\n", vbCrLf)  ' �� \n �滻Ϊ���з�
        translation = Replace(translation, "\r", vbCrLf)  ' �� \r �滻Ϊ���з�
        translation = Replace(translation, "\\", "\")      ' ȡ����б�ܵ�ת��

        ' ׷����Ϣ��TextBox3�����Զ�����
        If Len(TextBox3.text) > 0 Then
            TextBox3.text = TextBox3.text & vbCrLf
        End If
        TextBox3.text = TextBox3.text & "�û�: " & systemContent & vbCrLf & "�ظ�: " & translation & vbCrLf & vbCrLf
    Else
        MsgBox "δ����ȡ�ظ���������鷵�ص����ݸ�ʽ��", vbCritical
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
    Dim insertPoint As Long ' ���ڱ��������λ��

    ' ��ȡ�û��� TextBox3 ��ѡ�е��ı�
    insertText = TextBox3.SelText
    If insertText = "" Then
        MsgBox "���� TextBox3 ��ѡ��Ҫ������ı���", vbExclamation
        Exit Sub
    End If

    ' ����Ƿ���ѡ�е�����
    If Selection.Type = wdSelectionIP Then
        MsgBox "������ Word �ĵ���ѡ�в���λ�á�", vbExclamation
        Exit Sub
    End If

    ' ���зָ��ı�
    lines = Split(insertText, vbCrLf)

    ' ��ʼ��������ʽ����
    Set markdownRegex = CreateObject("VBScript.RegExp")
    With markdownRegex
        .Global = True
        .IgnoreCase = True
    End With

    ' ��ȡ��ǰѡ�з�Χ�Ľ���λ��
    insertPoint = Selection.range.End

    ' �����ı������� Markdown ��ʽ
    For i = LBound(lines) To UBound(lines)
        Dim formattedLine As String
        formattedLine = lines(i)

        ' ����Ӵֺ�б�壨***...***��
        regexPattern = "\*\*\*(.*?)\*\*\*"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(formattedLine)
        For Each mtch In matches
            formattedLine = Replace(formattedLine, mtch.value, mtch.SubMatches(0))
        Next mtch

        ' ����Ӵ֣�**...**��
        regexPattern = "\*\*(.*?)\*\*"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(formattedLine)
        For Each mtch In matches
            formattedLine = Replace(formattedLine, mtch.value, mtch.SubMatches(0))
        Next mtch

        ' ����б�壨*...*��
        regexPattern = "\*(.*?)\*"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(formattedLine)
        For Each mtch In matches
            formattedLine = Replace(formattedLine, mtch.value, mtch.SubMatches(0))
        Next mtch

        ' �������
        regexPattern = "(#{1,4})\s(.*?(?=##|\n|$))"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(formattedLine)
        For Each mtch In matches
            formattedLine = Replace(formattedLine, mtch.value, mtch.SubMatches(1))
        Next mtch

        ' ���봦�����ı�
        Set lineRange = ActiveDocument.range(insertPoint, insertPoint)
        lineRange.text = formattedLine & vbCrLf

        ' ��ʽ���Ӵֺ�б�壨***...***��
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

        ' ��ʽ���Ӵ֣�**...**��
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

        ' ��ʽ��б�壨*...*��
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

        ' ��ʽ������
        regexPattern = "(#{1,4})\s(.*?(?=##|\n|$))"
        markdownRegex.pattern = regexPattern
        Set matches = markdownRegex.Execute(lines(i))
        For Each mtch In matches
            Dim headerRange As range
            Set headerRange = lineRange.Duplicate
            headerRange.Start = lineRange.Start + InStr(formattedLine, mtch.SubMatches(1)) - 1
            headerRange.End = headerRange.Start + Len(mtch.SubMatches(1))
            headerRange.Font.Bold = True
        Next mtch

        ' ���²����λ��
        insertPoint = insertPoint + Len(formattedLine) + 1
    Next i
End Sub
