<!--#include virtual="/admin/templatetop.asp"-->
<%
    Set objLink = New Link

    If Request.QueryString("page") <> "" Then
        objLink.Pagination = Request.QueryString("page")
    End If    
%>
<h2>Admin | Links</h2>
<p><a href="/admin/link.asp">Add a new link</a>, or click on <em>edit</em> below to amend an existing one</p>
<table border="0" cellpadding="2px" cellspacing="0" class="records">
	<tr class="headings">
		<td>Name</td>
		<td>URL</td>
		<td>DateModified</td>
		<td>Published</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
<%
	strTemplate = "<tr>" & vbCrLf
	strTemplate = strTemplate & "<td>[name]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[url]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[datemodified]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[published]</td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/link.asp?id=[linkid]"">edit</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/deletelink.asp?id=[linkid]"">delete</a></td>" & vbCrLf
	strTemplate = strTemplate & "</tr>"

	objLink.Template = strTemplate
	Call objLink.GetAdminLinks()
%>
</table>
<%=objSlickCMS.AdminPaging(objLink.AdminLinksCount,objLink.Pagination,"links.asp")%>
<!--#include virtual="/admin/templatebottom.asp"-->
<%
    Set objLink = Nothing
%>