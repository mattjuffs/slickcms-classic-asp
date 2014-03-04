<!--#include virtual="/admin/templatetop.asp"-->
<%
    Set objComment = New Comment
    
    If Request.QueryString("page") <> "" Then
        objComment.Pagination = Request.QueryString("page")
    End If
%>
<h2>Admin | Comments</h2>
<p>Add a new comment by visiting the relevant post, or click on <em>edit</em> below to amend an existing one</p>
<table border="0" cellpadding="2px" cellspacing="0" class="records">
	<tr class="headings">
		<td>Commenter</td>
		<td>Post</td>
		<td>Content</td>
		<td>DateModified</td>
		<td>Published</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
<%
	strTemplate = "<tr>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/user.asp?id=[userid]"">[name]</a><br />[email]<br /><a href=""[url]"">[url]</a><br />[ip]</td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/post.asp?id=[postid]"">[posttitle]</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td>[content]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[datemodified]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[published]</td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/comment.asp?id=[commentid]"">edit</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/deletecomment.asp?id=[commentid]"">delete</a></td>" & vbCrLf
	strTemplate = strTemplate & "</tr>" & vbCrLf

	objComment.CommentsTemplate = strTemplate
	Call objComment.GetAdminComments()
%>
</table>
<%=objSlickCMS.AdminPaging(objComment.AdminLinksCount,objComment.Pagination,"comments.asp")%>
<!--#include virtual="/admin/templatebottom.asp"-->
<%
    Set objComment = Nothing
%>