<%@ Language=VBScript %>
<%
    Option Explicit
    
    Response.ContentType = "text/xml"
%>
<!--#include virtual="/slickcms/slickcms.asp"-->
<%
    Dim objFeed

    Set objFeed = New Feed
    Call objFeed.Sitemap()
    Set objFeed = Nothing
%>