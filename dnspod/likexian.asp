<%
''
 ' DNSPod API ASP Web 示例
 ' http://www.zhetenga.com/
 '
 ' Copyright 2011, Kexian Li
 ' Released under the MIT, BSD, and GPL Licenses.
 '
 ''

head = "" &_
"<!DOCTYPE html>" & vbCRLF &_
"<html>" & vbCRLF &_
"<head>" & vbCRLF &_
"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" & vbCRLF &_
"<title>DNSPod API ASP Web 示例-李院长</title>" & vbCRLF &_
"<style type=""text/css"">" & vbCRLF &_
"body {" & vbCRLF &_
"	background: #fff;" & vbCRLF &_
"	color: #000;" & vbCRLF &_
"	font: font:13px 'Helvetica Neue',Arial,Sans-serif;" & vbCRLF &_
"	margin:30px;" & vbCRLF &_
"	padding:0;" & vbCRLF &_
"}" & vbCRLF &_
"a {" & vbCRLF &_
"	color: #133DB6;" & vbCRLF &_
"	text-decoration:none;" & vbCRLF &_
"}" & vbCRLF &_
"a:hover {" & vbCRLF &_
"	color: #133DB6;" & vbCRLF &_
"	text-decoration:underline;" & vbCRLF &_
"}" & vbCRLF &_
"#likexian_box {" & vbCRLF &_
"	margin: auto;" & vbCRLF &_
"	width: 800px;" & vbCRLF &_
"}" & vbCRLF &_
"</style>" & vbCRLF &_
"</head>" & vbCRLF &_
"<body>" & vbCRLF &_
"<div id=""likexian_box"">" & vbCRLF

foot = "" &_
"</div>" & vbCRLF &_
"</body>" & vbCRLF &_
"</html>" & vbCRLF

login_form = head &_
"<form name=""login"" method=""post"" action=""index.asp?action=domainlist"">" & vbCRLF &_
"<div>账号：<input type=""text"" name=""login_email"" /></div>" & vbCRLF &_
"<div>密码：<input type=""password"" name=""login_password"" /></div>" & vbCRLF &_
"<div><input type=""submit"" value=""登录"" /></div>" & vbCRLF &_
"</form>" & vbCRLF &_
foot

domain_list = head &_
"<form name=""login"" method=""post"" action=""index.asp?action=domaincreate"">" & vbCRLF &_
"<div>域名：<input type=""text"" name=""domain"" /><input type=""submit"" value=""添加"" /></div>" & vbCRLF &_
"</form>" & vbCRLF &_
"<table cellspacing=""0"" cellpadding=""5"" border=""1"" width=""100%"">" & vbCRLF &_
"	<tr>" & vbCRLF &_
"		<th>编号</th><th>域名</th><th>等级</th><th>状态</th><th>扩展状态</th><th>记录</th><th>星标</th><th>更新</th><th>操作</th>" & vbCRLF &_
"	</tr>" & vbCRLF &_
"	{domain_list}" & vbCRLF &_
"</table>" & vbCRLF &_
foot

record_list = head &_
"<div><a href=""index.asp?action=domainlist"">域名管理</a></div>" & vbCRLF &_
"<table cellspacing=""0"" cellpadding=""5"" border=""1"" width=""100%"">" & vbCRLF &_
"	<tr>" & vbCRLF &_
"		<th>编号</th><th>子域名</th><th>类型</th><th>线路</th><th>记录</th><th>状态</th><th>MX</th><th>TTL</th><th>操作</th>" & vbCRLF &_
"	</tr>" & vbCRLF &_
"	{record_list}" & vbCRLF &_
"</table>" & vbCRLF &_
foot
%>
