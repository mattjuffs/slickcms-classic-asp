<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Save Post</h2>
<%
	Set objPost = New Post
	
	objPost.ID = Request.Form("id")
	objPost.UserID = Request.Form("userid")
	objPost.Title = Request.Form("title")
	objPost.Summary = Request.Form("summary")
	objPost.Content = Request.Form("wmd-input")
	objPost.Search = Request.Form("search")
	objPost.Published = Request.Form("published")
	objPost.Pageable = Request.Form("pageable")
	objPost.URL = Request.Form("URL")

	objPost.Save()
    intPostID = objPost.ID 'get ID of new/updated Post

	Set objPost = Nothing
	
	'#369 - allow the User to associate Categories and Tags to the Post
	Response.Redirect("/admin/relationship.asp?postid=" & intPostID)
	Response.End
%>
<!--#include virtual="/admin/templatebottom.asp"-->