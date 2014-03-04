<%@Language="VBScript"%>
<%
	Option Explicit
	'(C) Copyright MMIX Matthew Juffs (Slickhouse.com) - released under the Microsoft Reciprocal License (Ms-RL)
%>
<!--#include virtual="/slickcms/slickcms.asp"-->
<%
	'open global connection to database
	Call OpenDatabase()
	
	Set objSlickCMS = New SlickCMS
	Set objPost = New Post
	Set objCaptcha = New Captcha
	Set objComment = New Comment
	Set objCategory = New Category
	Set objTag = New Tag
	Set objStatistic = New Statistic

    'determines which page we're on and sets the various variables used from here on
	Call UrlHandler()
	
	'search
    objPost.Keywords = Request.Form("keywords")
    objPost.SearchTemplate = "<p><a href=""[url]"">[title]</a> - [summary]</p>"

    'contact form post
	If objPost.Url = "send-message" Then
        If objCaptcha.Process <>  "pass" Then
            Session("EmailSent") = false
        Else
		    Call objSlickCMS.SendMessage(Request.Form("email"), Request.Form("message"), Request.Form("name"))
        End If
        
        'set Session so user can re-use their entered data
        Session("Name") = Request.Form("name")
        Session("Email") = Request.Form("email")
        Session("Comment") = Request.Form("message")
        
        Response.Redirect("/contact/")
        Response.End
	End If
	
	'retrieve post if single page
    Select Case objPost.UrlType
        Case "post","date" 'single item
        	Call objPost.GetPost()
	End Select
	
	'generate meta tag data
    Call objSlickCMS.Meta()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=strTitle%></title>
<meta name="language" content="en-UK" />
<meta name="description" content="<%=strDescription%>" />
<meta name="generator" content="SlickCMS <%=Application("SlickCMS_Version")%>" />
<meta name="verify-v1" content="<%=Application("GoogleVerifyTag")%>" />
<link href="<%=Application("CDN")%>css/screen.css" rel="stylesheet" type="text/css" media="screen" />
<%Call objSlickCMS.CSS()%>
<link rel="alternate" type="application/rss+xml" title="<%=Application("SiteName") & " Posts RSS 2.0 Feed"%>" href="<%=Application("SiteURL")%>rss2.asp?t=posts" />
<link rel="alternate" type="application/rss+xml" title="<%=Application("SiteName") & " Comments RSS 2.0 Feed"%>" href="<%=Application("SiteURL")%>rss2.asp?t=comments" />
</head>

<body>

<div id="page">
    <!--#include virtual="/header.asp"-->
    <!--#include virtual="/sidebar.asp"-->
    <!--#include virtual="/content.asp"-->
    <!--#include virtual="/footer.asp"-->
</div>

<script src="<%=Application("CDN")%>scripts/slickcms.js" type="text/javascript"></script>
<%=objSlickCMS.GoogleAnalytics()%>

</body>
</html>
<%
	'destroy objects
	Set objStatistic = Nothing
	Set objTag = Nothing
	Set objCategory = Nothing
    Set objComment = Nothing
	Set objCaptcha = Nothing
	Set objPost = Nothing
	Set objSlickCMS = Nothing
	
	'close global connection to database
	Call CloseDatabase()
%>