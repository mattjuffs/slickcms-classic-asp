<%@Language="VBScript"%>
<%Option Explicit%>
<!--#include virtual="/slickcms/slickcms.asp"-->
<%
	Dim objUser, strEmail, strPassword
	
	strEmail = Request.Form("email")
	strPassword = Request.Form("password")
	
	If Len(strEmail) > 0 And Len(strPassword) > 0 Then
		Call OpenDatabase()
		Set objUser = New User
		objUser.Email = strEmail
		objUser.Password = strPassword
		Call objUser.Login()
		Set objUser = Nothing
		Call CloseDatabase()
	End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Login | <%=Application("TitleTag")%></title>
<meta name="language" content="en-UK" />
<link href="<%=Application("CDN")%>css/admin_screen.css" rel="stylesheet" type="text/css" media="screen" />
<script src="<%=Application("CDN")%>scripts/slickcms.js" type="text/javascript"></script>
</head>

<body>
	<form id="loginregister" method="post" action="/admin/login.asp">
		<h1><%=Application("SiteName")%></h1>
		<%
			If Session("Message") <> "" Then
				Response.Write("<p class=""message"">" & Session("Message") & "</p>")
				Session("Message") = Empty
			End If
			
			If Session("Error") <> "" Then
				Response.Write("<p class=""error"">" & Session("Error") & "</p>")
				Session("Error") = Empty
			End If
		%>
		<p><label for="email">Email</label><br /><input type="text" name="email" id="email" maxlength="255" size="40" /></p>
		<p><label for="password">Password</label><br /><input type="password" name="password" id="password" maxlength="1024" size="40" /></p>
		<p><input type="submit" value="Login" /></p>
		<p><small><a href="/admin/forgotpassword.asp">Forgotten your password?</a></small></p>
		<p><small><a href="/admin/register.asp">New User? Register here!</a></small></p>
		<p><small><a href="<%=Application("SiteURL")%>">Return to the homepage</a></small></p>
	</form>
</body>
</html>