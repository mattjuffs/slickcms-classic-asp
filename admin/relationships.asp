<!--#include virtual="/admin/templatetop.asp"-->
<%
    Set objRelationship = New Relationship
    
    If Request.QueryString("page") <> "" Then
        objRelationship.Pagination = Request.QueryString("page")
    End If
%>
<h2>Admin | Relationships</h2>
<p><a href="/admin/Relationship.asp">Add a new relationship</a>, or click on <em>edit</em> below to amend an existing one</p>
<table border="0" cellpadding="2px" cellspacing="0" class="records">
	<tr class="headings">
		<td>Category Name</td>
		<td>Tag Name</td>
		<td>Link Name</td>
		<td>Post Title</td>
		<td>User Name</td>
		<td>Order</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
<%
	strTemplate = "<tr>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/category.asp?id=[categoryid]"">[categoryname]</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/tag.asp?id=[tagid]"">[tagname]</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/link.asp?id=[linkid]"">[linkname]</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/post.asp?id=[postid]"">[posttitle]</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/user.asp?id=[userid]"">[username]</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td>[order]</td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/relationship.asp?id=[relationshipid]"">edit</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/deleteRelationship.asp?id=[relationshipid]"">delete</a></td>" & vbCrLf
	strTemplate = strTemplate & "</tr>"

	objRelationship.Template = strTemplate
	Call objRelationship.GetAdminRelationships()
%>
</table>
<%=objSlickCMS.AdminPaging(objRelationship.AdminRelationshipsCount,objRelationship.Pagination,"relationships.asp")%>
<!--#include virtual="/admin/templatebottom.asp"-->
<%
    Set objRelationship = Nothing
%>