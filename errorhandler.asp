<%@ Language=VBScript %>
<%
	Option Explicit

	Dim ASPErr
	Dim strError
	Dim sv 'server variable
	Dim EmailError

	Set ASPErr = Server.GetLastError
	Response.Clear
	
	strError = ""
	
	Call BuildError()
	Call LogErrorToFile()
	Call LogErrorToEmail()
%>
<!--#include virtual="/slickcms/slickcms.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Error 500 | <%=Application("TitleTag")%></title>
<meta name="language" content="en-UK" />
<meta name="description" content="An error has occurred" />
<link href="<%=Application("CDN")%>css/screen.css" rel="stylesheet" type="text/css" media="screen" />
</head>
<body>
<div id="error">
	<h1>An error has occurred!</h1>
	<p>The webmaster has been notified. Please try another page or revisit soon.</p>
	<p><a href="/">Return to the homepage</a></p>
</div>
</body>
</html>
<%
	Sub BuildError()
		If Err.number = 0 Then
			strError = strError & "------------------------------------------------------" & vbCrLf
			strError = strError & "Error Time: " & Now & vbCrLf
			strError = strError & "Session ID: " & Session.SessionID & vbCrLf
            strError = strError & "HTTP_X_ORIGINAL_URL: " & Request.ServerVariables("HTTP_X_ORIGINAL_URL") & vbCrLf
            strError = strError & "URL: " & Request.ServerVariables("URL") & vbCrLf
			strError = strError & "---------------------------------------------ASP Error" & vbCrLf
			If ASPErr.ASPCode <> "" Then strError = strError & "* Error #: " & ASPErr.ASPCode & vbCrLf
			If ASPErr.Number <> 0 Then strError = strError & "* COM Error #: " & ASPErr.Number & " (" & Hex (ASPErr.Number) & ")" & vbCrLf
			If ASPErr.Source <> "" Then strError = strError & "* Source: " & ASPErr.Source & vbCrLf
			If ASPErr.Category <> "" Then strError = strError & "* Category: " & ASPErr.Category & vbCrLf
			If ASPErr.File <> "" Then strError = strError & "* File: " & "/" & Request.ServerVariables ("SERVER_NAME") & ASPErr.File & vbCrLf
			If ASPErr.Line <> 0 Then strError = strError & "* Line, Column:" & ASPErr.Line & ", " & ASPErr.Column & vbCrLf
			If ASPErr.Description <> "" Then strError = strError & "* Description: " & ASPErr.Description & vbCrLf
			If ASPErr.ASPDescription <> "" Then strError = strError & "* ASP Desc: " & ASPErr.ASPDescription & vbCrLf
			strError = strError & "------------------------------------------HTTP Headers" & vbCrLf
			strError = strError & Replace(Request.ServerVariables("ALL_HTTP"),vbLf,vbCrLf)
			strError = strError & "--------------------------------------Server Variables" & vbCrLf
			For Each sv In Request.ServerVariables()
				strError = strError & UCASE(sv) & ":" & Request.ServerVariables(sv) & vbCrLf
			Next
			strError = strError & "-----------------------------------------------Session" & vbCrLf
			For Each sv In Session.Contents()
				strError = strError & UCASE(sv) & ":" & Session.Contents(sv) & vbCrLf
			Next
			strError = strError & "------------------------------------------------------" & vbCrLf
		End If
	End Sub

	Sub LogErrorToFile()
		Set objSlickCMS = New SlickCMS
		Call objSlickCMS.Log(strError)
		Set objSlickCMS = Nothing
	End Sub
	
	Sub LogErrorToEmail()	
		EmailError = Replace(strError,vbCrLf,"<br />")	
		Set objSlickCMS = New SlickCMS
		Call objSlickCMS.SendEmail(Application("ErrorsEmail"), Application("Email"), "An Error Occurred on " & Application("SiteName"), EmailError)
		Set objSlickCMS = Nothing
	End Sub
%>