<!--#include virtual="/admin/templatetop.asp"-->
<%	
	Set objCategory = New Category

	objCategory.ID = Request.QueryString("id")
	Call objCategory.GetAdminCategory()

	If objCategory.ID <> 0 Then
%>
	<h2>Admin | Edit Category</h2>
	<p>Edit the fields below as necessary, then click save to update the category:</p>
<%Else%>
	<h2>Admin | Add Category</h2>
	<p>Fill in all fields below and click save to add your category:</p>
<%End If%>

<form method="post" action="/admin/savecategory.asp">
	<table border="0" cellpadding="0" cellspacing="0" width="90%">
		<input type="hidden" name="id" value="<%=objCategory.ID%>" />
		<tr>
			<td>Name</td>
			<td><input type="text" name="name" value="<%=objCategory.Name%>" maxlength="255" size="50" /></td>
		</tr>
		<tr>
			<td>Type</td>
			<td>
			    <select name="type">
			        <option value=""></option>
			        <%If objCategory.CType = "Posts" Then strSelected = "selected" Else strSelected = "" %>
			        <option value="Posts" <%=strSelected%>>Posts</option>
			        <%If objCategory.CType = "Navigation" Then strSelected = "selected" Else strSelected = "" %>
			        <option value="Navigation" <%=strSelected%>>Navigation</option>
			        <%If objCategory.CType = "Users" Then strSelected = "selected" Else strSelected = "" %>
			        <option value="Users" <%=strSelected%>>Users</option>
                    <%If objCategory.CType = "Links" Then strSelected = "selected" Else strSelected = "" %>
			        <option value="Links" <%=strSelected%>>Links</option>
			    </select>
			</td>
		</tr>
		<tr>
			<td valign="top">Description</td>
			<td><textarea name="description" rows="10" cols="50"><%=objCategory.Description%></textarea></td>
		</tr>
		<tr>
			<td colspan="2"><input type="submit" name="submit" value="Save" /></td>
		</tr>
	</table>
</form>

<%
	Set objCategory = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->