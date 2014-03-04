<!--#include virtual="/admin/templatetop.asp"-->
<%
	Set objLink = New Link

	objLink.ID = Request.QueryString("id")
	Call objLink.GetAdminLink()

	If objLink.ID <> 0 Then
%>
	<h2>Admin | Edit Link</h2>
	<p>Edit the fields below as necessary, then click save to update the link:</p>
<%Else%>
	<h2>Admin | Add Link</h2>
	<p>Fill in all fields below and click save to add your link:</p>
<%End If%>

<form method="post" action="/admin/savelink.asp">
	<table border="0" cellpadding="0" cellspacing="0" width="90%">
		<input type="hidden" name="id" value="<%=objLink.ID%>" />
		<tr>
			<td>Name</td>
			<td><input type="text" name="name" value="<%=objLink.Name%>" maxlength="255" size="50" /></td>
		</tr>
		<tr>
			<td>URL</td>
			<td><input type="text" name="url" value="<%=objLink.URL%>" maxlength="1024" size="50" /></td>
		</tr>
		<tr>
			<td valign="top">Description</td>
			<td><textarea name="description" rows="10" cols="50"><%=objLink.Description%></textarea></td>
		</tr>
		<tr><td colspan="2"><hr /><h3>Metadata</h3></td></tr>
		<tr>
			<td>DateCreated</td>
			<td><%=objLink.DateCreated%></td>
		</tr>
		<tr>
			<td>DateModified</td>
			<td><%=objLink.DateModified%></td>
		</tr>
		<tr>
			<td>Published</td>
			<td>
				<select name="published">
				<%
					Select Case cstr(objLink.Published)
						Case "0",""
				%>
					<option value="0" selected="selected">Unpublished</option>
					<option value="1">Published</option>
				<%
						Case "1"
				%>
					<option value="0">Unpublished</option>
					<option value="1" selected="selected">Published</option>
				<%
					End Select
				%>
				</select>
			</td>
		</tr>
		<tr>
			<td colspan="2"><input type="submit" name="submit" value="Save" /></td>
		</tr>
	</table>
</form>

<%
	Set objLink = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->