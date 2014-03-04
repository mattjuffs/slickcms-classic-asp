<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Delete Tag</h2>
<%
	Set objTag = New Tag
	objTag.ID = Request.QueryString("id")
	objTag.Delete()
	Set objTag = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->