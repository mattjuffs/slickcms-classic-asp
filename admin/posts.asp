<!--#include virtual="/admin/templatetop.asp"-->
<%
    Set objPost = New Post
    
    If Request.QueryString("page") <> "" Then
        objPost.Pagination = Request.QueryString("page")
    End If
%>
<h2>Admin | Posts</h2>
<p><a href="/admin/post.asp">Add a new post</a>, or click on <em>edit</em> below to amend an existing one</p>
<table border="0" cellpadding="2px" cellspacing="0" class="records">
	<tr class="headings">
		<td>Author</td>
		<td>Title</td>
		<td>DateModified</td>
		<td>Published</td>
		<td>Pageable</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
<%
	strTemplate = "<tr>" & vbCrLf
	strTemplate = strTemplate & "<td>[author]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[title]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[datemodified]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[published]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[pageable]</td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/post.asp?id=[postid]"">edit</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/deletepost.asp?id=[postid]"">delete</a></td>" & vbCrLf
	strTemplate = strTemplate & "</tr>"

	objPost.PostsTemplate = strTemplate
	Call objPost.GetAdminPosts()
%>
</table>
<%=objSlickCMS.AdminPaging(objPost.AdminPostsCount,objPost.Pagination,"posts.asp")%>
<!--#include virtual="/admin/templatebottom.asp"-->
<%
    Set objPost = Nothing
%>