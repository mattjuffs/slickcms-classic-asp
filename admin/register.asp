<%@Language="VBScript"%>
<%Option Explicit%>
<!--#include virtual="/slickcms/slickcms.asp"-->
<%
	Dim objUser
	
	If Len(Request.Form("email")) > 0 Then
		If Request.Form("password") <> Request.Form("confirmpassword") Then
			Session("Error") = "Your passwords don't match!"
		Else
			Call OpenDatabase()
			Set objUser = New User
	
			objUser.Name = Request.Form("name") & ""
			objUser.Email = Request.Form("email") & ""
			objUser.Password = Request.Form("password") & ""
			objUser.URL = Request.Form("url") & ""
			objUser.IP = Request.ServerVariables("REMOTE_ADDR")
	
			objUser.Register()
	
			Set objUser = Nothing
			Call CloseDatabase()
		End If
	End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Register | <%=Application("TitleTag")%></title>
<meta name="language" content="en-UK" />
<link href="<%=Application("CDN")%>css/admin_screen.css" rel="stylesheet" type="text/css" media="screen" />
<script src="<%=Application("CDN")%>scripts/slickcms.js" type="text/javascript"></script>
</head>

<body>
	<form id="loginregister" method="post" action="/admin/register.asp">
		<h1><%=Application("SiteName")%></h1>
		<p class="message">Fill in all of the fields below to register:</p>
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
		<p><label for="name">Name</label><br /><input type="text" name="name" id="name" maxlength="50" size="40" value="<%=Request.Form("Name")%>" /></p>
		<p><label for="email">Email</label><br /><input type="text" name="email" id="email" maxlength="255" size="40" value="<%=Request.Form("Email")%>" /></p>
		<p><label for="password">Password</label><br /><input type="password" name="password" id="password" maxlength="1024" size="40" /></p>
		<p><label for="confirmpassword">Confirm Password</label><br /><input type="password" name="confirmpassword" id="confirmpassword" maxlength="1024" size="40" /></p>
		<p><label for="url">Website URL</label><br /><input type="text" name="url" id="url" maxlength="1024" size="40" value="<%=Request.Form("URL")%>" /></p>
		<p><input type="submit" value="Register" /></p>
		<p><small><a href="<%=Application("SiteURL")%>">Return to the homepage</a></small></p>
	</form>
</body>
</html>