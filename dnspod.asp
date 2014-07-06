<%
''
 ' DNSPod API ASP Web 示例
 ' http://www.likexian.com/
 '
 ' Copyright 2011-2014, Kexian Li
 ' Released under the Apache License, Version 2.0
 '
 ''

Class dnspod
    Public GradeList
    Public StatusList

    Private Sub Class_Initialize()
        Set GradeList = Server.CreateObject("Scripting.Dictionary")
        GradeList.Add "D_Free", "免费套餐"
        GradeList.Add "D_Plus", "豪华 VIP套餐"
        GradeList.Add "D_Extra", "企业I VIP套餐"
        GradeList.Add "D_Expert", "企业II VIP套餐"
        GradeList.Add "D_Ultra", "企业III VIP套餐"
        GradeList.Add "DP_Free", "新免费套餐"
        GradeList.Add "DP_Plus", "个人专业版"
        GradeList.Add "DP_Extra", "企业创业版"
        GradeList.Add "DP_Expert", "企业标准版"
        GradeList.Add "DP_Ultra", "企业旗舰版"

        Set StatusList = Server.CreateObject("Scripting.Dictionary")
        StatusList.Add "enable", "启用"
        StatusList.Add "pause", "暂停"
        StatusList.Add "spam", "封禁"
        StatusList.Add "lock", "锁定"
    End Sub

    Public Function ApiCall(strApi, strData)
    On Error Resume Next
        strApi = "https://dnsapi.cn/" & strApi
        strData = "login_email=" & Session("login_email") & "&login_password=" & Session("login_password") & "&login_code=" & Session("login_code") & "&format=xml&lang=cn&error_on_empty=no&" & strData

        strResult = PostData(strApi, strData, Session("cookies"))
        If strResult = "" Then
            Message "danger", "内部错误：调用失败", ""
        End If

        Set objRoot = GetRootNode(strResult)
        Set objNodes = objRoot.getElementsByTagName("dnspod/status")
        If objNodes(0).selectSingleNode("code").Text <> 1 And objNodes(0).selectSingleNode("code").Text <> 50 Then
            Message "danger", objNodes(0).selectSingleNode("message").Text, ""
        End If
        Set objNodes = Nothing

        Set ApiCall = objRoot
    End Function

    Public Function GetTemplate(strTemplate)
        Text = ReadText("template/" & strTemplate & ".html")
        GetTemplate = ReadText("template/index.html")
        GetTemplate = Replace(GetTemplate, "{{content}}", Text)
    End Function

    Public Sub Message(strStatus, strMessage, strUrl)
        If strStatus = "success" Then
            Status = "操作成功"
        Else
            Status = "操作失败"
        End If
        Text = GetTemplate("message")
        Text = Replace(Text, "{{title}}", Status)
        Text = Replace(Text, "{{status}}", strStatus)
        Text = Replace(Text, "{{message}}", strMessage)
        Text = Replace(Text, "{{url}}", strUrl)
        Response.Write(Text)
        Response.End
    End Sub

    Function ReadText(strFile)
        On Error Resume Next
        strFile = Server.MapPath(strFile)
        Set objStream = Server.CreateObject("ADODB.Stream")
        With objStream
            .Charset = "utf-8"
            .Open
            .LoadFromFile(strFile)
            ReadText = .ReadText()
        End With
        Set objStream = Nothing
    End Function

    Private Function GetRootNode(strData)
    On Error Resume Next
        Set GetRootNode = Server.CreateObject("Msxml2.DOMDocument")
        If Err.Number <> 0 Then
            Message "danger", "内部错误：服务器不支持Msxml2.DOMDocument", ""
        End If

        GetRootNode.Async = False
        GetRootNode.LoadXml(strData)
        If Err.Number <> 0 Then
            Message "danger", "内部错误：加载XML数据失败", ""
        End If
    End Function

    Private Function PostData(strUrl, strData, strCookies)
    On Error Resume Next
        Set objHttp = Server.CreateObject("Msxml2.XMLHTTP")
        If Err.Number <> 0 Then
            Message "danger", "内部错误：服务器不支持Msxml2.XMLHTTP", ""
        End If

        With objHttp
            .Open "post", strUrl, False, "", ""
            .SetRequestHeader "Content-Length", Len(strData)
            .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .SetRequestHeader "User-Agent", "DNSPod API ASP Web Client/1.0.0 (i@likexian.com)"
            If strCookies <> "" Then
                .SetRequestHeader "Cookie", strCookies
            End If
            .Send(strData)
            If .ReadyState <> 4 Then
                PostData = False
            Else
                PostData = BytesToStr(.ResponseBody)
                Cookies = ""
                Headers = .getAllResponseHeaders()
                Headers = Split(Headers, vbCrLf)
                For i = 0 To Ubound(Headers)
                    If Left(Headers(i), 13) = "Set-Cookie: t" Then
                        Cookies = Cookies & Mid(Headers(i), 12, InStr(Headers(i), ";") - 12) & "&"
                    End If
                Next
                If Cookies <> "" Then
                    Session("login_code") = ""
                    Session("cookies") = Left(Cookies, Len(Cookies) - 1)
                End If
            End If
        End With

        Set objHttp = Nothing
    End Function

    Private Function BytesToStr(bytBody)
    On Error Resume Next
        Set objStream = Server.CreateObject("ADODB.Stream")
        If Err.Number <> 0 Then
            Message "danger", "内部错误：服务器不支持ADODB.Stream", ""
        End If

        With objStream
            .Type = 1
            .Mode = 3
            .Open()
            .Write(bytBody)
            .Position = 0
            .Type = 2
            .Charset = "utf-8"
            BytesToStr = .ReadText()
            .Close()
        End With

        Set objStream = Nothing
    End Function
End Class
%>