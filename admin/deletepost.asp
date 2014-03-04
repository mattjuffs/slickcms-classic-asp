<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Delete Post</h2>
<%
	Set objPost = New Post
	objPost.ID = Request.QueryString("id")
	objPost.Delete()
	Set objPost = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->