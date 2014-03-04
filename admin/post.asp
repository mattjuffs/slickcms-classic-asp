<!--#include virtual="/admin/templatetop.asp"-->
<%
	Set objPost = New Post
	Set objUser = New User
	
	objUser.AuthorsTemplate = "<option value=""[id]"">[name]</option>"
	objUser.AuthorsTemplateSelected = "<option value=""[id]"" selected=""selected"">[name]</option>"

	objPost.ID = Request.QueryString("id")
	Call objPost.GetAdminPost()

	If objPost.ID <> 0 Then
%>
	<h2>Admin | Edit Post</h2>
	<p>Edit the fields below as necessary, then click save to update the post:</p>
<%Else%>
	<h2>Admin | Add Post</h2>
	<p>Fill in all fields below and click save to add your post:</p>
<%End If%>

<form method="post" action="/admin/savepost.asp">
	<table border="0" cellpadding="0" cellspacing="0" width="90%">
		<input type="hidden" name="id" value="<%=objPost.ID%>" />
		<tr>
			<td>Title</td>
			<td><input type="text" name="title" value="<%=objPost.Title%>" size="50" onblur="javascript:GenerateURL(this.value);" maxlength="255" /></td>
		</tr>
		<tr>
			<td>Permalink</td>
			<td><%=Application("SiteURL")%><input type="text" name="url" id="url" value="<%=objPost.URL%>" size="30" maxlength="255" />/ <small><em>(Url of the Post - excluding the date on Pageable Posts)</em></small></td>
		</tr>
		<tr>
			<td valign="top">Content</td>
			<td>
				<div id="wmd-editor" class="wmd-panel">
					<div id="wmd-button-bar"></div>
					<textarea name="wmd-input" id="wmd-input" rows="1" cols="1"><%=objPost.Content%></textarea>
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
			<td>Summary</td>
			<td><input type="text" name="summary" value="<%=objPost.Summary%> " size="50" maxlength="150" /> <small><em>(between 25 and 150 characters, as it's used for the meta description)</em></small></td>
		</tr>
		<tr>
			<td>Author</td>
			<td>
				<select name="userid">
					<option value="0">Choose...</option>
					<%
						objUser.ID = objPost.UserID
						If objUser.ID = 0 Then objUser.ID = Session("UserID") '#363
						Call objUser.GetAuthors()
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>Search</td>
			<td><input type="text" name="search" value="<%=objPost.Search%> " maxlength="255" /> <small><em>(specify additional keywords that are not present within your content, for search)</em></small></td>
		</tr>
		<%If objPost.ID <> 0 Then%>
		<tr>
			<td>DateCreated</td>
			<td><%=objPost.DateCreated%></td>
		</tr>
		<tr>
			<td>DateModified</td>
			<td><%=objPost.DateModified%></td>
		</tr>
		<%End If%>
		<tr>
			<td>Published</td>
			<td>
				<select name="published">
				<%
					Select Case cstr(objPost.Published)
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
				</select> <small><em>(Publish your Post if you want it to be visible to visitors)</em></small>
			</td>
		</tr>
		<tr>
			<td>Pageable</td>
			<td>
				<select name="pageable">
				<%
					Select Case cstr(objPost.Pageable)
						Case "0"
				%>
					<option value="0" selected="selected">No</option>
					<option value="1">Yes</option>
				<%
						Case "1","" 'default
				%>
					<option value="0">No</option>
					<option value="1" selected="selected">Yes</option>
				<%
					End Select
				%>
				</select> <small><em>(Yes if you want the Post to appear in paged results (e.g. Blog Posts); No for static/stand-alone pages)</em></small>
			</td>
		</tr>
		<tr>
			<td colspan="2"><input type="submit" name="submit" value="Save" /></td>
		</tr>
	</table>
</form>

<%
	Set objUser = Nothing
	Set objPost = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->