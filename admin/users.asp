<!--#include virtual="/admin/templatetop.asp"-->
<%
	Set objUser = New User
	
    If Request.QueryString("page") <> "" Then
        objUser.Pagination = Request.QueryString("page")
    End If
%>
<h2>Admin | Users</h2>
<p><a href="/admin/user.asp">Add a new user</a>, or click on <em>edit</em> below to amend an existing one</p>
<table border="0" cellpadding="2px" cellspacing="0" class="records">
	<tr class="headings">
		<td>Name</td>
		<td>Email</td>
		<td>URL</td>
		<td>DateModified</td>
		<td>Active</td>
		<td>LoginFails</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
<%	
	strTemplate = "<tr[class]>" & vbCrLf
	strTemplate = strTemplate & "<td>[name]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[email]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[url]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[datemodified]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[active]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[loginfails]</td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/user.asp?id=[userid]"">edit</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/deleteuser.asp?id=[userid]"">delete</a></td>" & vbCrLf
	strTemplate = strTemplate & "</tr>"

	objUser.Template = strTemplate
	Call objUser.GetAdminUsers()
%>
</table>
<%=objSlickCMS.AdminPaging(objUser.AdminUsersCount,objUser.Pagination,"users.asp")%>
<!--#include virtual="/admin/templatebottom.asp"-->
<%
	Set objUser = Nothing
%>