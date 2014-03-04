<%@Language="VBScript"%>
<%Option Explicit%>
<!--#include virtual="/slickcms/slickcms.asp"-->
<%
	Dim objUser
	Dim objRelationship
	Dim strSelected
    Dim intPostID

    Set objSlickCMS = New SlickCMS

	Call OpenDatabase()
	
	Set objUser = New User
	Call objUser.CheckLogin()
	Set objUser = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Admin | <%=Application("TitleTag")%></title>
<meta name="language" content="en-UK" />
<link href="<%=Application("CDN")%>css/admin_screen.css" rel="stylesheet" type="text/css" media="screen" />
<link href="<%=Application("CDN")%>wmd/wmd.css" rel="stylesheet" type="text/css" media="screen" />
<script src="<%=Application("CDN")%>scripts/slickcms.js" type="text/javascript"></script>
</head>

<body>
<div id="page">
	<div id="header">
		<ul>
			<li><a href="/admin/default.asp">Dashboard</a></li>
			<li>|</li>
			<li><a href="/admin/posts.asp">Posts</a></li>
			<li>|</li>
			<li><a href="/admin/comments.asp">Comments</a></li>
			<li>|</li>
			<li><a href="/admin/links.asp">Links</a></li>
			<li>|</li>
			<li><a href="/admin/categories.asp">Categories</a></li>
			<li>|</li>
			<li><a href="/admin/tags.asp">Tags</a></li>
			<li>|</li>
			<li><a href="/admin/relationships.asp">Relationships</a></li>
			<li>|</li>
			<li><a href="/admin/users.asp">Users</a></li>
			<li>|</li>
			<li><a href="/admin/feeds.asp">Feeds</a></li>
			<li>|</li>
			<li><a href="/">Site</a></li>
			<li>|</li>
			<li><a href="/admin/logout.asp">Logout <%=Session("Name")%></a></li>
		</ul>
	</div>
	<div id="content">
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