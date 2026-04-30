Attribute VB_Name = "Module2"
' ==========================================
' [���� ��ũ��] Jira Ƽ�� �ڵ� ���� (��¥, ��������, Tab ���� �Ϻ� �ذẻ)
' ==========================================
Sub CreateJiraIssueFromEmail()
    Dim objMail As Outlook.MailItem
    Dim jiraUrl As String, projectKey As String, issueType As String
    Dim authKey As String, payload As String
    Dim xmlHttp As Object
    Dim safeSubject As String, safeBody As String
    Dim fullBody As String, latestBody As String
    Dim fromPos As Long
    
    Dim inputLabels As String, labelArray() As String, jsonLabels As String
    Dim inputStoryPoints As String, inputDueDate As String
    Dim startDate As String, storyPointFieldId As String, startDateFieldId As String
    Dim i As Integer

    ' 1. ���� ���� Ȯ��
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "�̸����� ���� �������ּ���.", vbExclamation
        Exit Sub
    End If

    If TypeOf Application.ActiveExplorer.Selection(1) Is MailItem Then
        Set objMail = Application.ActiveExplorer.Selection(1)
    Else
        MsgBox "������ �������ּ���.", vbExclamation
        Exit Sub
    End If

    ' ==========================================
    ' �� ����� ȯ�� ���� ��
    ' ==========================================
    authKey = "Bearer YOUR_JIRA_TOKEN_HERE"  ' 실제 토큰으로 교체 필요
    
    ' ���� ã���� ���丮 ����Ʈ�� ������ �ʵ� ID
    storyPointFieldId = "customfield_10006"
    startDateFieldId = "customfield_10220"
    ' ==========================================

    DoEvents

    ' ==========================================
    ' 2. ����� �Է� �� ��� ����
    ' ==========================================
    inputLabels = InputBox("��(Labels)�� �Է��ϼ���." & vbCrLf & "(���� ���� ��� ��ǥ�� ����, ��: ATDC, VTS, Test, Maintenance, I/EI-E2, T2-ICC..)", "Jira �� �Է�")
    If StrPtr(inputLabels) = 0 Then GoTo CancelProcess

    inputStoryPoints = InputBox("���丮 ����Ʈ(Story Points)�� ���ڷ� �Է��ϼ���. 0.25����" & vbCrLf & "(�Է����� ������ ��ĭ���� �����˴ϴ�.)", "Jira ����Ʈ �Է�")
    If StrPtr(inputStoryPoints) = 0 Then GoTo CancelProcess

    inputDueDate = InputBox("���� ������(Due Date)�� �Է��ϼ���." & vbCrLf & "(����: YYYY-MM-DD)", "Due Date ����", Format(Date + 7, "yyyy-mm-dd"))
    If StrPtr(inputDueDate) = 0 Then GoTo CancelProcess
    
    startDate = Format(Date, "yyyy-mm-dd")

    If Trim(inputLabels) <> "" Then
        labelArray = Split(inputLabels, ",")
        jsonLabels = ""
        For i = 0 To UBound(labelArray)
            jsonLabels = jsonLabels & """" & Trim(labelArray(i)) & """"
            If i < UBound(labelArray) Then jsonLabels = jsonLabels & ", "
        Next i
    End If

    ' ==========================================
    ' 3. ���� ���� (�ֽ� ���� ���� + ���� ����)
    ' ==========================================
    fullBody = objMail.body
    fromPos = InStr(1, fullBody, "From:", vbTextCompare)
    If fromPos > 0 Then latestBody = Left(fullBody, fromPos - 1) Else latestBody = fullBody
    
    ' ���� ����
    latestBody = RemoveSignature(latestBody)
    
    ' �� ������ ������ ���� �Լ� ���� (Tab ���� ���� ����)
    safeSubject = CleanJSONText(objMail.Subject)
    safeBody = CleanJSONText(latestBody)

    ' ==========================================
    ' 4. JSON ���̷ε� ����
    ' ==========================================
    payload = "{""fields"": {" & _
              """project"": {""key"": """ & projectKey & """}," & _
              """summary"": """ & safeSubject & """," & _
              """description"": """ & safeBody & """," & _
              """issuetype"": {""name"": """ & issueType & """}," & _
              """" & startDateFieldId & """: """ & startDate & """"
    
    If IsDate(inputDueDate) Then
        payload = payload & ", ""duedate"": """ & Format(inputDueDate, "yyyy-mm-dd") & """"
    End If
    
    If jsonLabels <> "" Then payload = payload & ", ""labels"": [" & jsonLabels & "]"
    If IsNumeric(inputStoryPoints) Then payload = payload & ", """ & storyPointFieldId & """: " & inputStoryPoints
    
    payload = payload & "}}"

    ' ==========================================
    ' 5. API ���� �� ��� ó��
    ' ==========================================
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error Resume Next
    xmlHttp.Open "POST", jiraUrl, False
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    xmlHttp.setRequestHeader "Authorization", authKey
    xmlHttp.Send payload
    
    If Err.Number <> 0 Then
        MsgBox "��Ʈ��ũ ����: " & Err.Description, vbCritical
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If xmlHttp.status = 201 Then
        Dim responseText As String, issueKey As String
        responseText = xmlHttp.responseText
        i = InStr(responseText, """key"":""") + 7
        issueKey = mid(responseText, i, InStr(i, responseText, """") - i)
        
        ' ÷������ Ȯ�� �� ���ε�
        If objMail.Attachments.Count > 0 Then
            MsgBox "Jira Ƽ��(" & issueKey & ")�� �����Ǿ����ϴ�!" & vbCrLf & _
                   "������: " & startDate & vbCrLf & "������: " & inputDueDate & vbCrLf & vbCrLf & _
                   "÷������ " & objMail.Attachments.Count & "�� ���ε带 �����մϴ�.", vbInformation
            
            ProcessAttachments objMail, issueKey, authKey
            MsgBox "��� ÷������ ���ε尡 �Ϸ�Ǿ����ϴ�!", vbInformation
        Else
            MsgBox "Jira Ƽ��(" & issueKey & ")�� �����Ǿ����ϴ�." & vbCrLf & "÷������ ����.", vbInformation
        End If
        
        CreateObject("WScript.Shell").Run "https://devstack.vwgroup.com/jira/browse/" & issueKey
    Else
        MsgBox "���� ���� (" & xmlHttp.status & "): " & xmlHttp.responseText, vbCritical
    End If
    Exit Sub

CancelProcess:
    MsgBox "����ڿ� ���� �۾��� ��ҵǾ����ϴ�.", vbInformation
End Sub

' ==========================================
' [���� �Լ� 1] ���� ����
' ==========================================
Function RemoveSignature(ByVal strBody As String) As String
    Dim signOffs As Variant, i As Integer, pos As Long, minPos As Long
    signOffs = Array(vbCrLf & "Best regards", vbCrLf & "Regards,", vbCrLf & "�����մϴ�", vbCrLf & "Thanks", vbCrLf & "--")
    minPos = Len(strBody) + 1
    For i = LBound(signOffs) To UBound(signOffs)
        pos = InStr(1, strBody, signOffs(i), vbTextCompare)
        If pos > 0 And pos < minPos Then minPos = pos
    Next i
    If minPos <= Len(strBody) Then RemoveSignature = Trim(Left(strBody, minPos - 1)) Else RemoveSignature = Trim(strBody)
End Function

' ==========================================
' [���� �Լ� 2] �ؽ�Ʈ ���� �� ǥ ��ȯ (���� ���� �ٽ�)
' ==========================================
Function CleanJSONText(ByVal strText As String) As String
    If Len(strText) = 0 Then
        CleanJSONText = ""
        Exit Function
    End If

    ' JSON ����ǥ �� ������ �̽�������
    strText = Replace(strText, "\", "\\")
    strText = Replace(strText, """", "\""")
    strText = Replace(strText, vbCrLf, vbLf)
    strText = Replace(strText, vbCr, vbLf)
    
    Dim lines() As String
    Dim i As Long
    Dim inTable As Boolean
    Dim resultLines() As String
    Dim resultCount As Long
    
    lines = Split(strText, vbLf)
    ReDim resultLines(UBound(lines))
    resultCount = 0
    inTable = False
    
    For i = 0 To UBound(lines)
        Dim currentLine As String
        currentLine = Trim(lines(i))
        
        ' Tab(Chr 9) ���ڰ� ������ ǥ(Table) �������� ��ȯ
        If InStr(lines(i), Chr(9)) > 0 Then
            If Not inTable Then
                resultLines(resultCount) = "||" & Replace(lines(i), Chr(9), "||") & "||"
                inTable = True
            Else
                resultLines(resultCount) = "|" & Replace(lines(i), Chr(9), "|") & "|"
            End If
            resultCount = resultCount + 1
        Else
            If inTable And currentLine <> "" Then
                resultLines(resultCount - 1) = resultLines(resultCount - 1) & "\\\\" & lines(i)
            ElseIf currentLine = "" Then
                resultLines(resultCount) = lines(i)
                resultCount = resultCount + 1
                inTable = False
            Else
                resultLines(resultCount) = lines(i)
                resultCount = resultCount + 1
                inTable = False
            End If
        End If
    Next i
    
    ReDim Preserve resultLines(resultCount - 1)
    strText = Join(resultLines, "\n") ' ����Ű�� \n���� ����
    
    ' Ȥ�ö� �����ִ� ǥ ���� Tab ���ڴ� ���� 4ĭ���� ġȯ�Ͽ� ���� ��õ ����
    strText = Replace(strText, Chr(9), "    ")
    
    ' ���� ����(0~31) ���� ����
    Dim j As Integer
    For j = 0 To 31
        strText = Replace(strText, Chr(j), "")
    Next j

    CleanJSONText = strText
End Function

' ==========================================
' [���� �Լ� 3] ÷������ ���μ��� ó��
' ==========================================
Sub ProcessAttachments(objMail As Object, issueKey As String, authKey As String)
    Dim att As Object, tempPath As String
    tempPath = Environ("TEMP") & "\"
    For Each att In objMail.Attachments
        att.SaveAsFile tempPath & att.fileName
        Call UploadAttachmentToJira(issueKey, tempPath & att.fileName, authKey)
        On Error Resume Next
        Kill tempPath & att.fileName
        On Error GoTo 0
    Next
End Sub

' ==========================================
' [���� �Լ� 4] ���� ���� ����
' ==========================================
Sub UploadAttachmentToJira(issueKey As String, filePath As String, authKey As String)
    Dim uploadUrl As String, boundary As String
    Dim xmlHttp As Object, objStream As Object, bodyStream As Object
    Dim fileContent As Variant, fileName As String
    
    uploadUrl = "https://devstack.vwgroup.com/jira/rest/api/2/issue/" & issueKey & "/attachments"
    fileName = mid(filePath, InStrRev(filePath, "\") + 1)
    boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
    
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1
    objStream.Open
    objStream.LoadFromFile filePath
    fileContent = objStream.Read
    objStream.Close
    
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Set bodyStream = CreateObject("ADODB.Stream")
    bodyStream.Type = 1
    bodyStream.Open
    
    StringToStream "--" & boundary & vbCrLf, bodyStream
    StringToStream "Content-Disposition: form-data; name=""file""; filename=""" & fileName & """" & vbCrLf, bodyStream
    StringToStream "Content-Type: application/octet-stream" & vbCrLf & vbCrLf, bodyStream
    bodyStream.Write fileContent
    StringToStream vbCrLf & "--" & boundary & "--" & vbCrLf, bodyStream
    bodyStream.Position = 0
    
    DoEvents
    On Error Resume Next
    xmlHttp.Open "POST", uploadUrl, False
    xmlHttp.setRequestHeader "Authorization", authKey
    xmlHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    xmlHttp.setRequestHeader "X-Atlassian-Token", "no-check"
    xmlHttp.Send bodyStream.Read
    On Error GoTo 0
End Sub

' ==========================================
' [���� �Լ� 5] ��Ʈ�� ���ڵ�
' ==========================================
Sub StringToStream(str As String, ByRef stream As Object)
    Dim tempStream As Object
    Set tempStream = CreateObject("ADODB.Stream")
    tempStream.Type = 2
    tempStream.Charset = "utf-8"
    tempStream.Open
    tempStream.WriteText str
    tempStream.Position = 0
    tempStream.Type = 1
    tempStream.Position = 3
    stream.Write tempStream.Read
    tempStream.Close
End Sub

