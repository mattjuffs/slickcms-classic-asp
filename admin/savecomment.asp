<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Save Comment</h2>
<%
	Set objComment = New Comment

	objComment.ID = Request.Form("id")
	objComment.PostID = Request.Form("postid")
	objComment.UserID = Request.Form("userid")
	objComment.Name = Request.Form("name")
	objComment.Email = Request.Form("email")
	objComment.URL = Request.Form("url")
	objComment.Content = Request.Form("wmd-input")
	objComment.Published = Request.Form("published")

	Response.Write(objComment.Save())

	Set objComment = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->