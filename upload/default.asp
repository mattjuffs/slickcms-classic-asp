<%
	'This passes the Session("LoggedOn") variable between Classic ASP and ASP.NET

	'short-term this should suffice, but is open to attacks
	'long-term the whole of SlickCMS will be migrated to ASP.NET (Version 4.0)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Check Login</title>
</head>

<body>

<form method="post" action="/upload/transfer.aspx" name="checklogin" id="checklogin">
	<input type="hidden" name="LoggedOn" id="LoggedOn" value="<%=Session("LoggedOn")%>" />
</form>

<script type="text/javascript">checklogin.submit();</script>

</body>

</html>