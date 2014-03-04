<%@Language="VBScript"%>
<%Option Explicit%>
<!--#include virtual="/slickcms/slickcms.asp"-->
<%
	Dim objUser
	Set objUser = New User
	Call objUser.Logout()
	Set objUser = Nothing
%>