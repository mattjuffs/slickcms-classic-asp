<%@Language="VBScript"%>
<%Option Explicit%>
<!--#include virtual="/slickcms/slickcms.asp"-->
<%
	Dim objUser, strEmail
	
	strEmail = Request.Form("email")
	
	If Len(strEmail) > 0 Then
		Call OpenDatabase()
		Set objUser = New User
		objUser.Email = strEmail
		Call objUser.ResetPassword()
		Set objUser = Nothing
		Call CloseDatabase()
	Else
		Session("Message") = Application("ForgotPassword")
	End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Forgot Password | <%=Application("TitleTag")%></title>
<meta name="language" content="en-UK" />
<link href="<%=Application("CDN")%>css/admin_screen.css" rel="stylesheet" type="text/css" media="screen" />
<script src="<%=Application("CDN")%>scripts/slickcms.js" type="text/javascript"></script>
</head>

<body>
	<form id="loginregister" method="post" action="/admin/forgotpassword.asp">
		<h1><%=Application("SiteName")%></h1>
		<%
			If Session("Message") <> "" Then
				Response.Write("<p class=""message"">" & Session("Message") & "</p>")
				Session("Message") = ""
			End If
			
			If Session("Error") <> "" Then
				Response.Write("<p class=""error"">" & Session("Error") & "</p>")
				Session("Error") = ""
			End If
		%>
		<p><label for="email">Email</label><br /><input type="text" name="email" id="email" maxlength="255" size="40" /></p>
		<p><input type="submit" value="Reset Password" /></p>
		<p><small><a href="<%=Application("SiteURL")%>">Return to the homepage</a></small></p>
	</form>
</body>
</html>