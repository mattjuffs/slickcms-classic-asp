<!--#include virtual="/admin/templatetop.asp"-->
<%
	Set objComment = New Comment

	objComment.ID = Request.QueryString("id")
	Call objComment.GetAdminComment()
%>
<h2>Admin | Edit Comment<%=objComment.UserID %></h2>
<p>Edit the fields below as necessary, then click save to update the comment:</p>
<form method="post" action="/admin/savecomment.asp">
	<table border="0" cellpadding="0" cellspacing="0" width="90%">
		<input type="hidden" name="id" value="<%=objComment.ID%>" />
		<input type="hidden" name="postid" value="<%=objComment.PostID%>" />
		<input type="hidden" name="userid" value="<%=objComment.UserID%>" />
		<%If objComment.UserID <> 0 Then%>
		<!--Registered User-->
		<tr>
		    <td>Name</td>
		    <td><%=objComment.Name%><input type="hidden" name="name" value="<%=objComment.Name%>" /></td>
		</tr>
		<tr>
		    <td>Email</td>
		    <td><%=objComment.Email%><input type="hidden" name="email" value="<%=objComment.Email%>" /></td>
		</tr>
		<tr>
		    <td>URL</td>
		    <td><%=objComment.URL%><input type="hidden" name="url" value="<%=objComment.URL%>" /></td>
		</tr>
		<%Else%>
		<!--Visitor-->
		<tr>
			<td>Name</td>
			<td><input type="text" name="name" value="<%=objComment.Name%>" size="50" maxlength="50" /></td>
		</tr>
		<tr>
			<td>Email</td>
			<td><input type="text" name="email" value="<%=objComment.Email%>" size="50" maxlength="255" /></td>
		</tr>
		<tr>
			<td>URL</td>
			<td><input type="text" name="url" value="<%=objComment.URL%>" size="50" maxlength="1024" /></td>
		</tr>
		<%End If%>
		<tr>
			<td valign="top">Content</td>
			<td>
				<div id="wmd-editor" class="wmd-panel">
					<div id="wmd-button-bar"></div>
					<textarea name="wmd-input" id="wmd-input" rows="1" cols="1"><%=objComment.Content%></textarea>
				</div>
				<script type="text/javascript" src="/wmd/showdown.js"></script>
				<script type="text/javascript" src="/wmd/wmd.js"></script>
				<p>Need to <a href="/upload/" target="_blank">upload an image/file</a>?</p>
			</td>
		</tr>
		<tr>
			<td valign="top">Preview</td>
			<td><div id="wmd-preview" class="wmd-preview"></div></td>
		</tr>
		<tr>
			<td valign="top">Output</td>
			<td><div id="wmd-output" class="wmd-output"></div></td>
		</tr>
		<tr><td colspan="2"><hr /><h3>Metadata</h3></td></tr>
		<tr>
			<td>IP</td>
			<td><%=objComment.IP%></td>
		</tr>
		<tr>
			<td>HTTP_USER_AGENT</td>
			<td><small><%=objComment.HTTP_USER_AGENT%></small></td>
		</tr>
		<tr>
			<td>DateCreated</td>
			<td><%=objComment.DateCreated%></td>
		</tr>
		<tr>
			<td>DateModified</td>
			<td><%=objComment.DateModified%></td>
		</tr>
		<tr>
			<td>Published</td>
			<td>
				<select name="published">
				<%
					Select Case cstr(objComment.Published)
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
	Set objComment = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->