<%
''
 ' DNSPod API ASP Web 示例
 ' http://www.zhetenga.com/
 '
 ' Copyright 2011, Kexian Li
 ' Released under the MIT, BSD, and GPL Licenses.
 '
 ''

Class dnspod
	Public Function ApiCall(strApi, strData)
	On Error Resume Next
		strApi = "https://dnsapi.cn/" & strApi
		strData = "login_email=" & Session("login_email") & "&login_password=" & Session("login_password") & "&format=xml&lang=cn&error_on_empty=no&" & strData

		strResult = PostData(strApi, strData)
		If strResult = "" Then
			Response.Write("内部错误：调用失败")
			Response.End()
		End If

		Set objRoot = GetRootNode(strResult)
		Set objNodes = objRoot.getElementsByTagName("dnspod/status")
		If objNodes(0).selectSingleNode("code").Text <> 1 Then
			Response.Write(objNodes(0).selectSingleNode("message").Text)
			Response.End()
		End If
		Set objNodes = Nothing

		Set ApiCall = objRoot
	End Function

	Private Function GetRootNode(strData)
	On Error Resume Next
		Set GetRootNode = Server.CreateObject("Msxml2.DOMDocument")
		If Err.Number <> 0 Then
			Response.Write("内部错误：服务器不支持Msxml2.DOMDocument")
			Response.End()
		End If

		GetRootNode.Async = False
		GetRootNode.LoadXml(strData)
		If Err.Number <> 0 Then
			Response.Write("内部错误：加载XML数据失败")
			Response.End()
		End If
	End Function

	Private Function PostData(strUrl, strData)
	On Error Resume Next
		Set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
		If Err.Number <> 0 Then
			Response.Write("内部错误：服务器不支持Microsoft.XMLHTTP")
			Response.End()
		End If

		With objHttp
			.Open "post", strUrl, False, "", ""
			.SetRequestHeader "Content-Length", Len(strData)
			.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			.SetRequestHeader "User-Agent", "DNSPod API ASP Web Client/0.1 (shallwedance@126.com)"
			.Send(strData)
			If .ReadyState <> 4 Then
				PostData = False
			Else
				PostData = BytesToStr(.ResponseBody)
			End If
		End With

		Set objHttp = Nothing
	End Function

	Private Function BytesToStr(bytBody)
	On Error Resume Next
		Set objStream = Server.CreateObject("ADODB.Stream")
		If Err.Number <> 0 Then
			Response.Write("内部错误：服务器不支持ADODB.Stream")
			Response.End()
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
