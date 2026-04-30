Attribute VB_Name = "Module1"
' ==============================================================================
' [���� ö�� �ϰẻ] Outlook �̸��� -> Confluence ������ ��ȯ ��ũ��
' - �̸��� ���� ����(.msg) �ڵ� ÷�� ��� ����
' - &nbsp; �ؽ�Ʈ ���� ����
' - <SPAN> ��ø ����, ����ǥ ���� ����('m'), ����Ʈ, �ߺ� class �� ��� ���� ��õ ����
' - (����) ���Խ� 5019 ��Ÿ�� ���� ���� ��Ÿ �ذ�
' ==============================================================================

Public Sub CreateConfluencePageFromEmail()

    Dim objMail As Outlook.MailItem
    Set objMail = GetCurrentMailItem()
    If objMail Is Nothing Then
        MsgBox "�̸����� ���� �����ϰų�(��Ͽ��� Ŭ��) �Ǵ� �̸��� â�� ����� ���¿��� �������ּ���.", vbExclamation
        Exit Sub
    End If

    ' ------------------------------
    ' �� ����� ȯ�� ���� ��
    ' ------------------------------
    Dim confluenceUrl As String, spaceKey As String, parentPageId As String
    Dim authKey As String

    confluenceUrl = "https://devstack.vwgroup.com/confluence/rest/api/content"
    spaceKey = "VWGKIC"
    parentPageId = "7207011679"
    authKey = "Bearer <your_confluence_api_token_here>"
    ' ------------------------------

    Dim rawHTML As String, safeHTML As String, safeTitle As String
    Dim payload As String
    Dim http As Object

    DoEvents

    ' 1) HTML ���� + XHTML ����ȭ + ���� �±� ���� �Ұ�
    rawHTML = CleanHTMLPhaseFinal(objMail.htmlBody)

    ' 2) Inline CID �̹��� ġȯ
    rawHTML = ReplaceCidImagesToConfluence(objMail, rawHTML)

    ' 3) �������� ���ϸ� ���� + ÷�� ����Ʈ �ڽ� ����
    Dim emailFileName As String
    emailFileName = BuildOriginalMailFileName(objMail)
    rawHTML = BuildAttachmentsHeaderBox(objMail, rawHTML, emailFileName)

    ' 4) JSON payload
    safeTitle = CleanTextForJSON(objMail.Subject & " (" & Format(Now, "yyyy-mm-dd hh:nn") & ")")
    safeHTML = EscapeForJSON(rawHTML)

    payload = "{""type"":""page""," & _
              """title"":""" & safeTitle & """," & _
              """space"":{""key"":""" & spaceKey & """},"

    If Len(parentPageId) > 0 Then
        payload = payload & """ancestors"":[{""id"":""" & parentPageId & """}],"
    End If

    payload = payload & """body"":{""storage"":{""value"":""" & safeHTML & """,""representation"":""storage""}}}"

    ' 5) Create page API ȣ��
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    On Error GoTo EH_NET
    http.Open "POST", confluenceUrl, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "Authorization", authKey
    http.Send payload
    On Error GoTo 0

    If Not (http.status = 200 Or http.status = 201) Then
        MsgBox "������ ���� ���� (Status: " & http.status & ")" & vbCrLf & http.responseText, vbCritical
        Exit Sub
    End If

    ' 6) pageId ����
    Dim pageId As String
    pageId = ExtractPageId(http.responseText)
    If Len(pageId) = 0 Then
        MsgBox "pageId ���� ����. ����:" & vbCrLf & http.responseText, vbCritical
        Exit Sub
    End If

    ' 7) ���ε�: �������� + ÷������
    Dim uploadedFilesList As String
    uploadedFilesList = UploadAllAttachments(confluenceUrl, authKey, pageId, objMail, emailFileName, rawHTML)

    MsgBox "������ ������ �Ϸ�Ǿ����ϴ�." & vbCrLf & vbCrLf & "[���� ó�� ���]" & vbCrLf & uploadedFilesList, vbInformation, "�۾� �Ϸ�"
    CreateObject("WScript.Shell").Run "https://devstack.vwgroup.com/confluence/pages/viewpage.action?pageId=" & pageId

    Set http = Nothing
    Exit Sub

EH_NET:
    MsgBox "��Ʈ��ũ/��û ����: " & Err.Description, vbCritical
    On Error GoTo 0
End Sub


' ==============================================================================
' Explorer ���� / Inspector ���� ���� ��� ����
' ==============================================================================
Private Function GetCurrentMailItem() As Outlook.MailItem
    On Error Resume Next
    If Not Application.ActiveInspector Is Nothing Then
        If TypeOf Application.ActiveInspector.CurrentItem Is Outlook.MailItem Then
            Set GetCurrentMailItem = Application.ActiveInspector.CurrentItem
            Exit Function
        End If
    End If
    If Not Application.ActiveExplorer Is Nothing Then
        If Application.ActiveExplorer.Selection.Count > 0 Then
            If TypeOf Application.ActiveExplorer.Selection(1) Is Outlook.MailItem Then
                Set GetCurrentMailItem = Application.ActiveExplorer.Selection(1)
                Exit Function
            End If
        End If
    End If
    Set GetCurrentMailItem = Nothing
End Function


' ==============================================================================
' [�ٽ�] HTML ���� �� ��� XML/XHTML ���� �Ϻ� ���� ����
' ==============================================================================
Private Function CleanHTMLPhaseFinal(ByVal html As String) As String

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Multiline = True

    ' 0) &nbsp; ���� ���� ��� (���� ���� ��� &nbsp;�� �Ϲ� �������� ���� ġȯ)
    html = Replace(html, "&nbsp;", " ", 1, -1, vbTextCompare)
    html = Replace(html, "&amp;nbsp;", " ", 1, -1, vbTextCompare)
    html = Replace(html, "&#160;", " ", 1, -1, vbTextCompare)

    ' 1) body ���θ� ����
    html = ExtractBodyInnerHtml(html)

    ' 2) �ּ� �� ���Ǻ� �ּ� ����
    re.Pattern = "": html = re.Replace(html, "")
    re.Pattern = "<!\[if[\s\S]*?endif\]>": html = re.Replace(html, "")

    ' 3) MSHTML �뷱�� (DOM Ʈ�� ����ȭ)
    html = BalanceByHtmlFile(html)

    ' 4) Office / ��ǥ�� XML ��� ���� ����
    re.Pattern = "<\?[^>]*>": html = re.Replace(html, "")
    re.Pattern = "<\/?xml:[^>]*>": html = re.Replace(html, "")
    re.Pattern = "<\/?(?:o|w|v|m|x):[^>]*>": html = re.Replace(html, "")
    re.Pattern = "\s+(?:o|w|v|m|x):[a-zA-Z0-9_-]+\s*=\s*(?:""[^""]*""|'[^']*'|[^\s>]+)": html = re.Replace(html, "")
    re.Pattern = "\s+(nowrap|checked|disabled|readonly|noshade|compact|multiple|noresize)\b(?!\s*=)": html = re.Replace(html, "")

    ' 5) ���ʿ��� ��Ÿ/��Ÿ�� �±� �Ұ�
    re.Pattern = "<\/?(?:meta|link|base|title|wbr)\b[^>]*>": html = re.Replace(html, "")
    re.Pattern = "<style[^>]*>[\s\S]*?<\/style>": html = re.Replace(html, "")

    ' 6) �ߺ� class ���� ����
    re.Pattern = "\s+class\s*=\s*(?:""[^""]*""|'[^']*'|[^\s>]+)": html = re.Replace(html, "")

    ' 7) [SPAN ��ø ���� �ذ�] �ζ��� �±� ������ ���� �Ұ� (��ҹ��� ���� ���� ����)
    re.Pattern = "<\/?(?:[sS][pP][aA][nN]|[fF][oO][nN][tT]|[cC][eE][nN][tT][eE][rR])\b[^>]*>"
    html = re.Replace(html, "")

    ' 8) ����Ʈ, ���̺� ��ҹ��� �ҹ��ڷ� ��ġ �� Confluence Class ����
    re.Pattern = "<[uU][lL]\b([^>]*)>": html = re.Replace(html, "<ul$1>")
    re.Pattern = "<\/[uU][lL]>": html = re.Replace(html, "</ul>")
    re.Pattern = "<[oO][lL]\b([^>]*)>": html = re.Replace(html, "<ol$1>")
    re.Pattern = "<\/[oO][lL]>": html = re.Replace(html, "</ol>")
    re.Pattern = "<[lL][iI]\b([^>]*)>": html = re.Replace(html, "<li$1>")
    re.Pattern = "<\/[lL][iI]>": html = re.Replace(html, "</li>")

    re.Pattern = "<[tT][aA][bB][lL][eE]\b([^>]*)>": html = re.Replace(html, "<table class=""confluenceTable""$1>")
    re.Pattern = "<\/[tT][aA][bB][lL][eE]>": html = re.Replace(html, "</table>")
    re.Pattern = "<[tT][rR]\b([^>]*)>": html = re.Replace(html, "<tr class=""confluenceTr""$1>")
    re.Pattern = "<\/[tT][rR]>": html = re.Replace(html, "</tr>")
    re.Pattern = "<[tT][dD]\b([^>]*)>": html = re.Replace(html, "<td class=""confluenceTd""$1>")
    re.Pattern = "<\/[tT][dD]>": html = re.Replace(html, "</td>")
    re.Pattern = "<[tT][hH]\b([^>]*)>": html = re.Replace(html, "<th class=""confluenceTh""$1>")
    re.Pattern = "<\/[tT][hH]>": html = re.Replace(html, "</th>")
    re.Pattern = "<\/?(?:thead|tbody|tfoot)\b[^>]*>": html = re.Replace(html, "")

    ' 9) ���� �±׸� DIV�� ġȯ
    re.Pattern = "<[pP]\b([^>]*)>": html = re.Replace(html, "<div$1>")
    re.Pattern = "<\/[pP]>": html = re.Replace(html, "</div>")

    ' 10) ['m' ���� �ذ�] ����ǥ ���� �Ӽ� ���� ����
    re.Pattern = "(\s[a-zA-Z0-9_\-:]+)=([^""'\s>]+)"
    html = re.Replace(html, "$1=""$2""")
    html = re.Replace(html, "$1=""$2""")
    html = re.Replace(html, "$1=""$2""")
    html = re.Replace(html, "$1=""$2""")

    ' 11) bare & ����
    html = FixBareAmpersands(html)

    ' 12) void tag �ڰ� ����(self-closing) ����ȭ
    html = NormalizeVoidTags(html)

    CleanHTMLPhaseFinal = html
    Set re = Nothing
End Function


' ==============================================================================
' <body ...>...</body> ���θ� ����
' ==============================================================================
Private Function ExtractBodyInnerHtml(ByVal html As String) As String
    Dim s As Long, e As Long
    s = InStr(1, html, "<body", vbTextCompare)
    If s > 0 Then
        s = InStr(s, html, ">") + 1
        e = InStr(1, html, "</body>", vbTextCompare)
        If e > s Then
            ExtractBodyInnerHtml = mid$(html, s, e - s)
            Exit Function
        End If
    End If
    ExtractBodyInnerHtml = html
End Function


' ==============================================================================
' MSHTML(htmlfile) �뷱��
' ==============================================================================
Private Function BalanceByHtmlFile(ByVal html As String) As String
    On Error GoTo FAILSAFE
    Dim doc As Object
    Set doc = CreateObject("htmlfile")
    doc.Open
    doc.Write "<!DOCTYPE html><meta http-equiv=""X-UA-Compatible"" content=""IE=edge""><body>" & html & "</body>"
    doc.Close
    BalanceByHtmlFile = doc.body.innerHTML
    Set doc = Nothing
    Exit Function
FAILSAFE:
    BalanceByHtmlFile = html
End Function


' ==============================================================================
' ��ƼƼ�� �ƴ� '&'�� '&amp;'�� ����
' ==============================================================================
Private Function FixBareAmpersands(ByVal html As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-f]+;)"
    FixBareAmpersands = re.Replace(html, "&amp;")
    Set re = Nothing
End Function


' ==============================================================================
' void tag self-closing ���� (���� 5019 �ذ�Ϸ�)
' ==============================================================================
Private Function NormalizeVoidTags(ByVal html As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True

    ' �̹� ���� �±� ���� (���Խ� ���� ������: [^>]* �� ����)
    re.Pattern = "<(hr|br|img|col|input)\b([^>]*)\/>"
    html = re.Replace(html, "<$1$2 />")
    ' ������ ���� �±� ���� ����
    re.Pattern = "<(hr|br|img|col|input)\b([^>]*)>"
    html = re.Replace(html, "<$1$2 />")
    
    html = Replace(html, "/ />", "/>")
    html = Replace(html, "  />", " />")

    NormalizeVoidTags = html
    Set re = Nothing
End Function


' ==============================================================================
' CID inline �̹��� -> Confluence attachment image ��ũ�� ġȯ
' ==============================================================================
Private Function ReplaceCidImagesToConfluence(ByVal mail As Outlook.MailItem, ByVal html As String) As String
    Dim cidProp As String: cidProp = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    Dim att As Outlook.Attachment
    Dim cidValue As String

    Dim reImg As Object
    Set reImg = CreateObject("VBScript.RegExp")
    reImg.IgnoreCase = True
    reImg.Global = True

    If mail.Attachments.Count > 0 Then
        For Each att In mail.Attachments
            If att.Type = 1 Then
                cidValue = ""
                On Error Resume Next
                cidValue = att.PropertyAccessor.GetProperty(cidProp)
                On Error GoTo 0

                If Len(cidValue) > 0 Then
                    Dim safeFn As String
                    safeFn = EscapeForXmlAttr(att.fileName)
                    reImg.Pattern = "<img[^>]*\s+src\s*=\s*(""|')?cid:" & EscapeRegex(cidValue) & "(\1)?[^>]*\/?>"
                    html = reImg.Replace(html, "<ac:image><ri:attachment ri:filename=""" & safeFn & """ /></ac:image>")
                End If
            End If
        Next att
    End If

    reImg.Pattern = "<img[^>]*\s+src\s*=\s*(""|')?cid:([^""'@>]+)(?:@[^""'>]+)?(\1)?[^>]*\/?>"
    html = reImg.Replace(html, "<ac:image><ri:attachment ri:filename=""$2"" /></ac:image>")

    ReplaceCidImagesToConfluence = html
    Set reImg = Nothing
End Function


' ==============================================================================
' ��� ÷�� �ڽ� ���� (�������� ��ũ + ÷������ ��ũ)
' ==============================================================================
Private Function BuildAttachmentsHeaderBox(ByVal mail As Outlook.MailItem, ByVal bodyHtml As String, ByVal emailFileName As String) As String
    Dim cidProp As String: cidProp = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    Dim att As Outlook.Attachment
    Dim cidValue As String
    Dim attHtml As String: attHtml = ""

    attHtml = attHtml & "<p style=""margin: 5px 0;""><strong>���� ����:</strong> " & _
              "<ac:link><ri:attachment ri:filename=""" & EscapeForXmlAttr(emailFileName) & """ />" & _
              "<ac:plain-text-link-body><![CDATA[" & emailFileName & "]]></ac:plain-text-link-body></ac:link></p>"

    Dim idx As Long: idx = 1
    If mail.Attachments.Count > 0 Then
        For Each att In mail.Attachments
            If att.Type = 1 Or att.Type = 5 Then
                cidValue = ""
                On Error Resume Next
                cidValue = att.PropertyAccessor.GetProperty(cidProp)
                On Error GoTo 0

                If (Len(cidValue) = 0) Or (InStr(1, bodyHtml, cidValue, vbTextCompare) = 0) Then
                    attHtml = attHtml & "<p style=""margin: 5px 0;""><strong>÷������" & idx & ":</strong> " & _
                              "<ac:link><ri:attachment ri:filename=""" & EscapeForXmlAttr(att.fileName) & """ />" & _
                              "<ac:plain-text-link-body><![CDATA[" & att.fileName & "]]></ac:plain-text-link-body></ac:link></p>"
                    idx = idx + 1
                End If
            End If
        Next att
    End If

    BuildAttachmentsHeaderBox = "<div style=""padding: 15px; margin-bottom: 20px; background-color: #f4f5f7; border-left: 4px solid #0052cc; border-radius: 3px;"">" & _
                               "<h4 style=""margin-top: 0;"">�̸��� �� ÷������ ���</h4>" & attHtml & _
                               "</div><hr/>" & bodyHtml
End Function


' ==============================================================================
' �������� ���ϸ� ����
' ==============================================================================
Private Function BuildOriginalMailFileName(ByVal mail As Outlook.MailItem) As String
    Dim baseName As String
    baseName = CleanFileName(mail.Subject)
    If Len(Trim$(baseName)) = 0 Then baseName = "Untitled_Email"
    If Len(baseName) > 120 Then baseName = Left$(baseName, 120)
    BuildOriginalMailFileName = baseName & ".msg"
End Function


' ==============================================================================
' ���ε�: ��������(.msg) + ÷������
' ==============================================================================
Private Function UploadAllAttachments(ByVal baseUrl As String, ByVal authKey As String, ByVal pageId As String, _
                                     ByVal mail As Outlook.MailItem, ByVal emailFileName As String, ByVal currentHtml As String) As String
    Dim tempFolder As String
    tempFolder = Environ$("TEMP") & "\"
    Dim result As String: result = ""

    ' ���� ���� ó��
    Dim emailPath As String
    emailPath = GetUniqueTempPath(tempFolder, emailFileName)
    On Error Resume Next
    mail.SaveAs emailPath, 3 ' 3 = olMSG
    If Err.Number <> 0 Then
        result = result & "- " & emailFileName & " (���� ���� ����: " & Err.Description & ")" & vbCrLf
        Err.Clear
    Else
        If UploadToConfluence(baseUrl, pageId, emailPath, authKey) Then
            result = result & "- " & emailFileName & " (����: ����)" & vbCrLf
        Else
            result = result & "- " & emailFileName & " (����: ����!)" & vbCrLf
        End If
        Kill emailPath
    End If
    On Error GoTo 0

    ' �Ϲ� ÷������ ó��
    Dim cidProp As String: cidProp = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    Dim att As Outlook.Attachment
    Dim cidValue As String, tempPath As String

    If mail.Attachments.Count > 0 Then
        For Each att In mail.Attachments
            If att.Type = 1 Or att.Type = 5 Then
                cidValue = ""
                On Error Resume Next
                cidValue = att.PropertyAccessor.GetProperty(cidProp)
                On Error GoTo 0

                If (Len(cidValue) = 0) Or (InStr(1, currentHtml, cidValue, vbTextCompare) = 0) Then
                    tempPath = GetUniqueTempPath(tempFolder, att.fileName)
                    On Error Resume Next
                    att.SaveAsFile tempPath
                    If Err.Number <> 0 Then
                        result = result & "- " & att.fileName & " (���� ����)" & vbCrLf
                        Err.Clear
                    Else
                        If UploadToConfluence(baseUrl, pageId, tempPath, authKey) Then
                            result = result & "- " & att.fileName & " (÷��: ����)" & vbCrLf
                        Else
                            result = result & "- " & att.fileName & " (÷��: ����!)" & vbCrLf
                        End If
                        Kill tempPath
                    End If
                    On Error GoTo 0
                End If
            End If
        Next att
    End If
    UploadAllAttachments = result
End Function


' ==============================================================================
' ���� JSON���� pageId ����
' ==============================================================================
Private Function ExtractPageId(ByVal json As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = """id""\s*:\s*""(\d+)"""
    If re.Test(json) Then
        Set m = re.Execute(json)
        ExtractPageId = m(0).SubMatches(0)
    Else
        ExtractPageId = ""
    End If
    Set re = Nothing
End Function


' ==============================================================================
' JSON escape
' ==============================================================================
Private Function EscapeForJSON(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    EscapeForJSON = s
End Function


' ==============================================================================
' ����(JSON)�� �ؽ�Ʈ ����
' ==============================================================================
Private Function CleanTextForJSON(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    CleanTextForJSON = s
End Function


' ==============================================================================
' ���� ���ε�(multipart/form-data)
' ==============================================================================
Private Function UploadToConfluence(ByVal baseUrl As String, ByVal pageId As String, ByVal filePath As String, ByVal authKey As String) As Boolean
    Dim uploadUrl As String
    uploadUrl = baseUrl & "/" & pageId & "/child/attachment"
    
    Dim fileName As String
    fileName = mid$(filePath, InStrRev(filePath, "\") + 1)
    
    Dim boundary As String
    boundary = "----WebKitFormBoundary" & Format(Now, "yyyymmddhhnnss") & CStr(Int((9999 * Rnd) + 1))
    
    Dim fileStream As Object
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1
    fileStream.Open
    fileStream.LoadFromFile filePath
    Dim fileBytes As Variant
    fileBytes = fileStream.Read
    fileStream.Close
    
    Dim body As Object
    Set body = CreateObject("ADODB.Stream")
    body.Type = 1
    body.Open
    
    StringToStreamUTF8 "--" & boundary & vbCrLf, body
    StringToStreamUTF8 "Content-Disposition: form-data; name=""file""; filename=""" & fileName & """" & vbCrLf, body
    StringToStreamUTF8 "Content-Type: application/octet-stream" & vbCrLf & vbCrLf, body
    body.Write fileBytes
    StringToStreamUTF8 vbCrLf & "--" & boundary & "--" & vbCrLf, body
    body.Position = 0
    
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", uploadUrl, False
    http.setRequestHeader "Authorization", authKey
    http.setRequestHeader "X-Atlassian-Token", "nocheck"
    http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.Send body.Read
    
    UploadToConfluence = (http.status = 200 Or http.status = 201)
    
    Set http = Nothing
    Set body = Nothing
    Set fileStream = Nothing
End Function


' ==============================================================================
' UTF-8 string -> binary stream
' ==============================================================================
Private Sub StringToStreamUTF8(ByVal s As String, ByRef stm As Object)
    Dim tmp As Object
    Set tmp = CreateObject("ADODB.Stream")
    tmp.Type = 2
    tmp.Charset = "utf-8"
    tmp.Open
    tmp.WriteText s
    tmp.Position = 0
    tmp.Type = 1
    tmp.Position = 3
    stm.Write tmp.Read
    tmp.Close
    Set tmp = Nothing
End Sub


' ==============================================================================
' ���ϸ� �������� ����
' ==============================================================================
Private Function CleanFileName(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "[\\/:*?""<>|]"
    CleanFileName = re.Replace(s, "_")
    Set re = Nothing
End Function


' ==============================================================================
' TEMP ���ϸ� �浹 ����
' ==============================================================================
Private Function GetUniqueTempPath(ByVal folderPath As String, ByVal fileName As String) As String
    Dim baseName As String, ext As String, p As Long
    Dim candidate As String, idx As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then
        baseName = Left$(fileName, p - 1)
        ext = mid$(fileName, p)
    Else
        baseName = fileName
        ext = ""
    End If
    candidate = folderPath & fileName
    idx = 1
    Do While Len(Dir$(candidate)) > 0
        candidate = folderPath & baseName & "_" & idx & ext
        idx = idx + 1
    Loop
    GetUniqueTempPath = candidate
End Function


' ==============================================================================
' XML attribute escape
' ==============================================================================
Private Function EscapeForXmlAttr(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    EscapeForXmlAttr = s
End Function


' ==============================================================================
' Regex escape
' ==============================================================================
Private Function EscapeRegex(ByVal s As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "([\\\^\$\.\|\?\*\+\(\)\[\]\{\}])"
    EscapeRegex = re.Replace(s, "\$1")
    Set re = Nothing
End Function

