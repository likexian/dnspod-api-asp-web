<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
''
 ' DNSPod API ASP Web 示例
 ' http://www.likexian.com/
 '
 ' Copyright 2011-2014, Kexian Li
 ' Released under the Apache License, Version 2.0
 '
 ''

On Error Resume Next
Response.Charset = "UTF-8"
%>
<!--#include file="dnspod.asp"-->
<%
Set dnspod_api = new dnspod
action = Trim(Request.QueryString("action"))

If action = "domainlist" Then
	If Request.Form("login_code") = "" Then
		If Request.Form("login_email") = "" Then
			If Session("login_email") = "" Then
				dnspod_api.Message "danger", "请输入登录账号。", -1
			End If
		Else
			Session("login_email") = Request.Form("login_email")
		End If

		If Request.Form("login_password") = "" Then
			If Session("login_password") = "" Then
				dnspod_api.Message "danger", "请输入登录密码。", -1
			End If
		Else
			Session("login_password") = Request.Form("login_password")
		End If

		Session("login_code") = ""
	Else
		Session("login_code") = Request.Form("login_code")
	End If

	Set objXML = dnspod_api.ApiCall("Domain.List", "")
	Set objNodes = objXML.getElementsByTagName("dnspod/status")
	If objNodes(0).selectSingleNode("code").Text = "50" Then
		Response.Redirect("index.asp?action=logind")
		Response.End
	End If

	List = ""
	DomainSub = dnspod_api.ReadText("./template/domain_sub.html")
	Set objNodes = objXML.getElementsByTagName("dnspod/domains/item")
	For i = 0 To objNodes.Length - 1
		If objNodes(i).selectSingleNode("status").Text = "pause" Then
			status_new = "enable"
			status_text = "启用"
		Else
			status_new = "disable"
			status_text = "禁用"
		End If
		ListSub = Replace(DomainSub, "{{id}}", objNodes(i).selectSingleNode("id").Text)
		ListSub = Replace(ListSub, "{{domain}}", objNodes(i).selectSingleNode("name").Text)
		ListSub = Replace(ListSub, "{{grade}}", dnspod_api.GradeList.Item(objNodes(i).selectSingleNode("grade").Text))
		ListSub = Replace(ListSub, "{{status}}", dnspod_api.StatusList.Item(objNodes(i).selectSingleNode("status").Text))
		ListSub = Replace(ListSub, "{{status_new}}", status_new)
		ListSub = Replace(ListSub, "{{status_text}}", status_text)
		ListSub = Replace(ListSub, "{{records}}", objNodes(i).selectSingleNode("records").Text)
		ListSub = Replace(ListSub, "{{updated_on}}", objNodes(i).selectSingleNode("updated_on").Text)
		List = List & ListSub
	Next

	Text = dnspod_api.GetTemplate("domain")
	Text = Replace(Text, "{{title}}", "域名列表")
	Text = Replace(Text, "{{list}}", List)
ElseIf action = "domaincreate" Then
	If Request.Form("domain") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If

	Set objXML = dnspod_api.ApiCall("Domain.Create", "domain=" & Request.Form("domain"))

	dnspod_api.Message "success", "添加成功。", "index.asp?action=domainlist"
ElseIf action = "domainstatus" Then
	If Request.QueryString("domain_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If
	If Request.QueryString("status") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If

	Session("login_code") = Request.Form("login_code")
	Set objXML = dnspod_api.ApiCall("Domain.Status", "domain_id=" & Request.QueryString("domain_id") & "&status=" & Request.QueryString("status"))
	Set objNodes = objXML.getElementsByTagName("dnspod/status")
	If objNodes(0).selectSingleNode("code").Text = "50" Then
		Response.Redirect("index.asp?action=domainstatusd&domain_id=" & Request.QueryString("domain_id") & "&status=" & Request.QueryString("status"))
		Response.End
	End If

	If Request.QueryString("status") = "enable" Then
		Status = "启用"
	Else
		Status = "暂停"
	End If

	dnspod_api.Message "success", Status & "成功。", "index.asp?action=domainlist"
ElseIf action = "domainremove" Then
	If Request.QueryString("domain_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If

	Session("login_code") = Request.Form("login_code")
	Set objXML = dnspod_api.ApiCall("Domain.Remove", "domain_id=" & Request.QueryString("domain_id"))
	Set objNodes = objXML.getElementsByTagName("dnspod/status")
	If objNodes(0).selectSingleNode("code").Text = "50" Then
		Response.Redirect("index.asp?action=domainremoved&domain_id=" & Request.QueryString("domain_id"))
		Response.End
	End If

	dnspod_api.Message "success", "删除成功。", "index.asp?action=domainlist"
ElseIf action = "recordlist" Then
	If Request.QueryString("domain_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If

	Set objXML = dnspod_api.ApiCall("Record.List", "domain_id=" & Request.QueryString("domain_id"))

	List = ""
	RecordSub = dnspod_api.ReadText("./template/record_sub.html")
	Set objNodes = objXML.getElementsByTagName("dnspod/records/item")
	For i = 0 To objNodes.Length - 1
		If objNodes(i).selectSingleNode("enabled").Text = 1 Then
			Enabled = "启用"
			StatusNew = "disable"
			StatusText = "暂停"
		Else
			Enabled = "暂停"
			StatusNew = "enable"
			StatusText = "启用"
		End If
		If objNodes(i).selectSingleNode("mx").Text <> "" Then
			MX = objNodes(i).selectSingleNode("mx").Text
		Else
			MX = "-"
		End If
		ListSub = Replace(RecordSub, "{{domain_id}}", Request.QueryString("domain_id"))
		ListSub = Replace(ListSub, "{{id}}", objNodes(i).selectSingleNode("id").Text)
		ListSub = Replace(ListSub, "{{name}}", objNodes(i).selectSingleNode("name").Text)
		ListSub = Replace(ListSub, "{{value}}", objNodes(i).selectSingleNode("value").Text)
		ListSub = Replace(ListSub, "{{type}}", objNodes(i).selectSingleNode("type").Text)
		ListSub = Replace(ListSub, "{{line}}", objNodes(i).selectSingleNode("line").Text)
		ListSub = Replace(ListSub, "{{enabled}}", Enabled)
		ListSub = Replace(ListSub, "{{status_new}}", StatusNew)
		ListSub = Replace(ListSub, "{{status_text}}", StatusText)
		ListSub = Replace(ListSub, "{{mx}}", MX)
		ListSub = Replace(ListSub, "{{ttl}}", objNodes(i).selectSingleNode("ttl").Text)
		List = List & ListSub
	Next

	Set objNodes = objXML.getElementsByTagName("dnspod/domain")
	Text = dnspod_api.GetTemplate("record")
	Text = Replace(Text, "{{title}}", "记录列表 - " & objNodes(0).selectSingleNode("name").Text)
	Text = Replace(Text, "{{list}}", List)
	Text = Replace(Text, "{{domain_id}}", objNodes(0).selectSingleNode("id").Text)
	Text = Replace(Text, "{{grade}}", objNodes(0).selectSingleNode("grade").Text)
ElseIf action = "recordcreatef" Then
	If Request.QueryString("domain_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If

	If Session("type_" & Request.QueryString("grade")) = "" Then
		Set objXML = dnspod_api.ApiCall("Record.Type", "domain_grade=" & Request.QueryString("grade"))
		Set objNodes = objXML.getElementsByTagName("dnspod/types/item")
		TypeList = ""
		For i = 0 To objNodes.Length - 1
			TypeList = TypeList & objNodes(i).Text & ","
		Next
		Session("type_" & Request.QueryString("grade")) = Left(TypeList, Len(TypeList) - 1)
	End If

	If Session("type_" & Request.QueryString("grade")) <> "" Then
		Types = Split(Session("type_" & Request.QueryString("grade")), ",")
		TypeList = ""
		For i = 0 To UBound(Types)
			TypeList = TypeList & "<option value=""" & Types(i) & """>" & Types(i) & "</option>"
		Next
	End If

	If Session("line_" & Request.QueryString("grade")) = "" Then
		Set objXML = dnspod_api.ApiCall("Record.Line", "domain_grade=" & Request.QueryString("grade"))
		Set objNodes = objXML.getElementsByTagName("dnspod/lines/item")
		LineList = ""
		For i = 0 To objNodes.Length - 1
			LineList = LineList & objNodes(i).Text & ","
		Next
		Session("line_" & Request.QueryString("grade")) = Left(LineList, Len(LineList) - 1)
	End If

	If Session("line_" & Request.QueryString("grade")) <> "" Then
		Lines = Split(Session("line_" & Request.QueryString("grade")), ",")
		LineList = ""
		For i = 0 To UBound(Lines)
			LineList = LineList & "<option value=""" & Lines(i) & """>" & Lines(i) & "</option>"
		Next
	End If

	Text = dnspod_api.GetTemplate("recordcreatef")
	Text = Replace(Text, "{{title}}", "添加记录")
	Text = Replace(Text, "{{action}}", "recordcreate")
	Text = Replace(Text, "{{domain_id}}", Request.QueryString("domain_id"))
	Text = Replace(Text, "{{record_id}}", Request.QueryString("record_id"))
	Text = Replace(Text, "{{type_list}}", TypeList)
	Text = Replace(Text, "{{line_list}}", LineList)
	Text = Replace(Text, "{{sub_domain}}", "")
	Text = Replace(Text, "{{value}}", "")
	Text = Replace(Text, "{{mx}}", "10")
	Text = Replace(Text, "{{ttl}}", "600")
ElseIf action = "recordcreate" Then
	If Request.QueryString("domain_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If

	If Request.Form("sub_domain") = "" Then
		Request.Form("sub_domain") = "@"
	End If

	If Request.Form("value") = "" Then
		dnspod_api.Message "danger", "请输入记录值。", -1
	End If

	If Request.Form("type") = "MX" And Request.Form("mx") = "" Then
		Request.Form("mx") = 10
	End If

	If Request.Form("ttl") = "" Then
		Request.Form("ttl") = 600
	End If

	Set objXML = dnspod_api.ApiCall("Record.Create", "domain_id=" & Request.QueryString("domain_id") & "&sub_domain=" & Request.Form("sub_domain") & "&record_type=" & Request.Form("type") & "&record_line=" & Request.Form("line") & "&value=" & Request.Form("value") & "&mx=" & Request.Form("mx") & "&ttl=" & Request.Form("ttl"))

	dnspod_api.Message "success", "添加成功。", "index.asp?action=recordlist&domain_id=" & Request.QueryString("domain_id")
ElseIf action = "recordeditf" Then
	If Request.QueryString("domain_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If
	If Request.QueryString("record_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If

	Set objXML = dnspod_api.ApiCall("Record.Info", "domain_id=" & Request.QueryString("domain_id") & "&record_id=" & Request.QueryString("record_id"))
	Set objRecords = objXML.getElementsByTagName("dnspod/record")

	If Session("type_" & Request.QueryString("grade")) = "" Then
		Set objXML = dnspod_api.ApiCall("Record.Type", "domain_grade=" & Request.QueryString("grade"))
		Set objNodes = objXML.getElementsByTagName("dnspod/types/item")
		TypeList = ""
		For i = 0 To objNodes.Length - 1
			TypeList = TypeList & objNodes(i).Text & ","
		Next
		Session("type_" & Request.QueryString("grade")) = Left(TypeList, Len(TypeList) - 1)
	End If

	If Session("type_" & Request.QueryString("grade")) <> "" Then
		Types = Split(Session("type_" & Request.QueryString("grade")), ",")
		TypeList = ""
		For i = 0 To UBound(Types)
			If objRecords(0).selectSingleNode("record_type").Text = Types(i) Then
				Check = " selected=""selected"""
			Else
				Check = ""
			End If
			TypeList = TypeList & "<option value=""" & Types(i) & """" & Check & ">" & Types(i) & "</option>"
		Next
	End If

	If Session("line_" & Request.QueryString("grade")) = "" Then
		Set objXML = dnspod_api.ApiCall("Record.Line", "domain_grade=" & Request.QueryString("grade"))
		Set objNodes = objXML.getElementsByTagName("dnspod/lines/item")
		LineList = ""
		For i = 0 To objNodes.Length - 1
			LineList = LineList & objNodes(i).Text & ","
		Next
		Session("line_" & Request.QueryString("grade")) = Left(LineList, Len(LineList) - 1)
	End If

	If Session("line_" & Request.QueryString("grade")) <> "" Then
		Lines = Split(Session("line_" & Request.QueryString("grade")), ",")
		LineList = ""
		For i = 0 To UBound(Lines)
			If objRecords(0).selectSingleNode("record_line").Text = Lines(i) Then
				Check = " selected=""selected"""
			Else
				Check = ""
			End If
			LineList = LineList & "<option value=""" & Lines(i) & """" & Check & ">" & Lines(i) & "</option>"
		Next
	End If

	Text = dnspod_api.GetTemplate("recordcreatef")
	Text = Replace(Text, "{{title}}", "修改记录")
	Text = Replace(Text, "{{action}}", "recordedit")
	Text = Replace(Text, "{{domain_id}}", Request.QueryString("domain_id"))
	Text = Replace(Text, "{{record_id}}", Request.QueryString("record_id"))
	Text = Replace(Text, "{{type_list}}", TypeList)
	Text = Replace(Text, "{{line_list}}", LineList)
	Text = Replace(Text, "{{sub_domain}}", objRecords(0).selectSingleNode("sub_domain").Text)
	Text = Replace(Text, "{{value}}", objRecords(0).selectSingleNode("value").Text)
	Text = Replace(Text, "{{mx}}", objRecords(0).selectSingleNode("mx").Text)
	Text = Replace(Text, "{{ttl}}", objRecords(0).selectSingleNode("ttl").Text)
ElseIf action = "recordedit" Then
	If Request.QueryString("domain_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If
	If Request.QueryString("record_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If

	If Request.Form("sub_domain") = "" Then
		Request.Form("sub_domain") = "@"
	End If

	If Request.Form("value") = "" Then
		dnspod_api.Message "danger", "请输入记录值。", -1
	End If

	If Request.Form("type") = "MX" And Request.Form("mx") = "" Then
		Request.Form("mx") = 10
	End If

	If Request.Form("ttl") = "" Then
		Request.Form("ttl") = 600
	End If

	Set objXML = dnspod_api.ApiCall("Record.Modify", "domain_id=" & Request.QueryString("domain_id") & "&record_id=" & Request.QueryString("record_id") & "&sub_domain=" & Request.Form("sub_domain") & "&record_type=" & Request.Form("type") & "&record_line=" & Request.Form("line") & "&value=" & Request.Form("value") & "&mx=" & Request.Form("mx") & "&ttl=" & Request.Form("ttl"))

	dnspod_api.Message "success", "修改成功。", "index.asp?action=recordlist&domain_id=" & Request.QueryString("domain_id")
ElseIf action = "recordremove" Then
	If Request.QueryString("domain_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If
	If Request.QueryString("record_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If

	Set objXML = dnspod_api.ApiCall("Record.Remove", "domain_id=" & Request.QueryString("domain_id") & "&record_id=" & Request.QueryString("record_id"))

	dnspod_api.Message "success", "删除成功。", "index.asp?action=recordlist&domain_id=" & Request.QueryString("domain_id")
ElseIf action = "recordstatus" Then
	If Request.QueryString("domain_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If
	If Request.QueryString("record_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If
	If Request.QueryString("status") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If

	Set objXML = dnspod_api.ApiCall("Record.Status", "domain_id=" & Request.QueryString("domain_id") & "&record_id=" & Request.QueryString("record_id") & "&status=" & Request.QueryString("status"))

	If Request.QueryString("status") = "enable" Then
		Status = "启用"
	Else
		Status = "暂停"
	End If

	dnspod_api.Message "success",  Status & "成功。", "index.asp?action=recordlist&domain_id=" & Request.QueryString("domain_id")
ElseIf action = "domainstatusd" Then
	If Request.QueryString("domain_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If
	If Request.QueryString("status") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If
	If Request.QueryString("status") = "enable" Then
		Status = "启用"
	Else
		Status = "暂停"
	End If
	Text = dnspod_api.GetTemplate("logind")
	Text = Replace(Text, "{{title}}", "域名" & Status)
	Text = Replace(Text, "{{action}}", "domainstatus&domain_id=" & Request.QueryString("domain_id") & "&status=" & Request.QueryString("status"))
ElseIf action = "domainremoved" Then
	If Request.QueryString("domain_id") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If
	Text = dnspod_api.GetTemplate("logind")
	Text = Replace(Text, "{{title}}", "域名删除")
	Text = Replace(Text, "{{action}}", "domainremove&domain_id=" & Request.QueryString("domain_id"))
ElseIf action = "logind" Then
	If Session("login_email") = "" Or Session("login_password") = "" Then
		dnspod_api.Message "danger", "参数错误。", -1
	End If
	Text = dnspod_api.GetTemplate("logind")
	Text = Replace(Text, "{{title}}", "用户登录")
	Text = Replace(Text, "{{action}}", "domainlist")
Else
	Text = dnspod_api.GetTemplate("login")
	Text = Replace(Text, "{{title}}", "用户登录")
End If

Response.Write(Text)
%>