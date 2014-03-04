<!--#include virtual="/admin/templatetop.asp"-->
<%
	Set objUser = New User

	objUser.ID = Request.QueryString("id")
	Call objUser.GetAdminUser()

	If objUser.ID <> 0 Then
%>
	<h2>Admin | Edit User</h2>
	<p>Edit the fields below as necessary, then click save to update the user:</p>
<%Else%>
	<h2>Admin | Add User</h2>
	<p>Fill in all fields below and click save to add the user:</p>
<%End If%>

<form method="post" action="/admin/saveuser.asp">
	<table border="0" cellpadding="0" cellspacing="0" width="90%">
		<input type="hidden" name="id" value="<%=objUser.ID%>" />
		<tr>
			<td>Name</td>
			<td><input type="text" name="name" value="<%=objUser.Name%>" size="50" maxlength="50" /> <small><em>(displayed on the site)</em></small></td>
		</tr>
		<tr>
			<td>Email</td>
			<td><input type="text" name="email" value="<%=objUser.Email%>" size="50" maxlength="255" /> <small><em>(Email is used alongside Password for Authentication)</em></small></td>
		</tr>
		<%If objUser.ID <> 0 Then%>
		<tr><td colspan="2"><p>If you want to update your password, enter your new one below - otherwise leave this blank</p></td></tr>
		<%End If%>
		<tr>
			<td>Password</td>
			<td><input type="password" name="password" value="" size="50" maxlength="1024" /></td>
		</tr>
		<tr>
			<td>Confirm Password</td>
			<td><input type="password" name="confirmpassword" value="" size="50" maxlength="1024" /></td>
		</tr>
		<tr>
			<td>URL</td>
			<td><input type="text" name="url" value="<%=objUser.URL%>" size="50" maxlength="1024" /> <small><em>(of your website)</em></small></td>
		</tr>
		<%If objUser.ID <> 0 Then%>
		<tr>
			<td>IP</td>
			<td><input type="text" name="ip" value="<%=objUser.IP%>" readonly="readonly" size="50" maxlength="15" /></td>
		</tr>
		<%End If%>
		<tr>
			<td valign="top">Biography</td>
			<td>
				<div id="wmd-editor" class="wmd-panel">
					<div id="wmd-button-bar"></div>
					<textarea name="wmd-input" id="wmd-input" rows="1" cols="1"><%=objUser.Biography%></textarea>
				</div>
				<script type="text/javascript" src="/wmd/showdown.js"></script>
				<script type="text/javascript" src="/wmd/wmd.js"></script>
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
		<%If objUser.ID <> 0 Then%>
		<tr>
			<td>DateCreated</td>
			<td><%=objUser.DateCreated%></td>
		</tr>
		<tr>
			<td>DateModified</td>
			<td><%=objUser.DateModified%></td>
		</tr>
		<%End If%>
		<tr>
			<td>Active</td>
			<td>
				<select name="active">
				<%
					Select Case cstr(objUser.Active)
						Case "0",""
				%>
					<option value="0" selected="selected">No</option>
					<option value="1">Yes</option>
				<%
						Case "1"
				%>
					<option value="0">No</option>
					<option value="1" selected="selected">Yes</option>
				<%
					End Select
				%>
				</select> <small><em>(only Active Users can login)</em></small>
			</td>
		</tr>
		<tr>
			<td>LoginFails</td>
			<td><input type="text" name="loginfails" value="<%=objUser.LoginFails%>" size="3" maxlength="10" /> <small><em>(how many times this User has failed to login successfully)</em></small></td>
		</tr>
		<tr>
			<td colspan="2"><input type="submit" name="submit" value="Save" /></td>
		</tr>
	</table>
</form>

<%
	Set objUser = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->