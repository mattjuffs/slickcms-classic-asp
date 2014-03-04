<!--#include virtual="/admin/templatetop.asp"-->
<%
    Set objTag = New Tag
    
    If Request.QueryString("page") <> "" Then
        objTag.Pagination = Request.QueryString("page")
    End If
%>
<h2>Admin | Tags</h2>
<p><a href="/admin/tag.asp">Add a new tag</a>, or click on <em>edit</em> below to amend an existing one</p>
<table border="0" cellpadding="2px" cellspacing="0" class="records">
	<tr class="headings">
		<td>Name</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
<%
	strTemplate = "<tr>" & vbCrLf
	strTemplate = strTemplate & "<td>[name]</td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/tag.asp?id=[tagid]"">edit</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/deletetag.asp?id=[tagid]"">delete</a></td>" & vbCrLf
	strTemplate = strTemplate & "</tr>"

	objTag.Template = strTemplate
	Call objTag.GetAdminTags()
%>
</table>
<%=objSlickCMS.AdminPaging(objTag.AdminTagsCount,objTag.Pagination,"tags.asp")%>
<!--#include virtual="/admin/templatebottom.asp"-->
<%
	Set objTag = Nothing
%>