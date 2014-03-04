<%@ Language=VBScript %>
<%
    Option Explicit
    
    Response.ContentType = "application/rss+xml"
%>
<!--#include virtual="/slickcms/slickcms.asp"-->
<%
    Dim objFeed

    Set objFeed = New Feed
    
    objFeed.Version = "2.0"
    objFeed.FType = Request.QueryString("t")
    
    Call objFeed.RSS()
    
    Set objFeed = Nothing
%>