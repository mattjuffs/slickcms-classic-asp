<!--#include virtual="/admin/templatetop.asp"-->
<%	
	Set objTag = New Tag

	objTag.ID = Request.QueryString("id")
	Call objTag.GetAdminTag()

	If objTag.ID <> 0 Then
%>
	<h2>Admin | Edit Tag</h2>
	<p>Edit the fields below as necessary, then click save to update the tag:</p>
<%Else%>
	<h2>Admin | Add Tag</h2>
	<p>Fill in all fields below and click save to add your tag:</p>
<%End If%>

<form method="post" action="/admin/savetag.asp">
	<table border="0" cellpadding="0" cellspacing="0" width="90%">
		<input type="hidden" name="id" value="<%=objTag.ID%>" />
		<tr>
			<td>Name</td>
			<td><input type="text" name="name" value="<%=objTag.Name%>" maxlength="255" size="50" /></td>
		</tr>
		<tr>
			<td colspan="2"><input type="submit" name="submit" value="Save" /></td>
		</tr>
	</table>
</form>

<%
	Set objTag = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->