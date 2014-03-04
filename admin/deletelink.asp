<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Delete Link</h2>
<%
	Set objLink = New Link
	objLink.ID = Request.QueryString("id")
	objLink.Delete()
	Set objLink = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->