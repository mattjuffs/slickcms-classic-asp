<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Delete Comment</h2>
<%
	Set objComment = New Comment
	objComment.ID = Request.QueryString("id")
	objComment.Delete()
	Set objComment = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->