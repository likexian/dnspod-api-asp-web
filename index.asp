<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
''
 ' DNSPod API ASP Web 示例
 ' http://www.zhetenga.com/
 '
 ' Copyright 2011, Kexian Li
 ' Released under the MIT, BSD, and GPL Licenses.
 '
 ''

On Error Resume Next
Response.Charset = "UTF-8"
%>
<!--#include file="dnspod/dnspod.asp"-->
<!--#include file="dnspod/likexian.asp"-->
<%
Set dnspod_api = new dnspod
action = Trim(Request.QueryString("action"))

If action = "domainlist" Then
	If Request.Form("login_email") = "" Then
		If Session("login_email") = "" Then
			Response.Write("请输入登录账号。")
			Response.End()
		End If
	Else
		Session("login_email") = Trim(Request.Form("login_email"))
	End If

	If Request.Form("login_password") = "" Then
		If Session("login_password") = "" Then
			Response.Write("请输入登录密码。")
			Response.End()
		End If
	Else
		Session("login_password") = Trim(Request.Form("login_password"))
	End If

	Set objXML = dnspod_api.ApiCall("Domain.List", "")
	Set objNodes = objXML.getElementsByTagName("dnspod/domains/item")
	For i = 0 To objNodes.Length - 1
		list = list & "<tr><td>" & objNodes(i).selectSingleNode("id").Text & "</td><td>" & objNodes(i).selectSingleNode("name").Text & "</td><td>" & objNodes(i).selectSingleNode("grade").Text & "</td><td>" & objNodes(i).selectSingleNode("status").Text & "</td><td>" & objNodes(i).selectSingleNode("ext_status").Text & "</td><td>" & objNodes(i).selectSingleNode("records").Text & "</td><td>" & objNodes(i).selectSingleNode("is_mark").Text & "</td><td>" & objNodes(i).selectSingleNode("updated_on").Text & "</td><td><a href='index.asp?action=recordlist&domain_id=" & objNodes(i).selectSingleNode("id").Text & "'>记录</a> <a href='index.asp?action=domainremove&domain_id=" & objNodes(i).selectSingleNode("id").Text & "'>删除</a></td></tr>"
	Next

	Response.Write(Replace(domain_list, "{domain_list}", List))
ElseIf action = "domaincreate" Then
	If Request.Form("domain") = "" Then
		Response.Write("参数错误。")
		Response.End()
	End If

	Set objXML = dnspod_api.ApiCall("Domain.Create", "domain=" & Request.Form("domain"))

	Response.Write("添加成功，<a href=""index.asp?action=domainlist"">点击返回</a>。")
ElseIf action = "domainremove" Then
	If Request.QueryString("domain_id") = "" Then
		Response.Write("参数错误。")
		Response.End()
	End If

	Set objXML = dnspod_api.ApiCall("Domain.Remove", "domain_id=" & Request.QueryString("domain_id"))

	Response.Write("删除成功，<a href=""index.asp?action=domainlist"">点击返回</a>。")
ElseIf action = "recordlist" Then
	If Request.QueryString("domain_id") = "" Then
		Response.Write("参数错误。")
		Response.End()
	End If

	Set objXML = dnspod_api.ApiCall("Record.List", "domain_id=" & Request.QueryString("domain_id"))
	Set objNodes = objXML.getElementsByTagName("dnspod/records/item")
	For i = 0 To objNodes.Length - 1
		list = list & "<tr><td>" & objNodes(i).selectSingleNode("id").Text & "</td><td>" & objNodes(i).selectSingleNode("name").Text & "</td><td>" & objNodes(i).selectSingleNode("type").Text & "</td><td>" & objNodes(i).selectSingleNode("line").Text & "</td><td>" & objNodes(i).selectSingleNode("value").Text & "</td><td>" & objNodes(i).selectSingleNode("enabled").Text & "</td><td>" & objNodes(i).selectSingleNode("mx").Text & "</td><td>" & objNodes(i).selectSingleNode("ttl").Text & "</td><td></td></tr>"
	Next

	Response.Write(Replace(record_list, "{record_list}", List))
Else
	Response.Write(login_form)
End If
%>
