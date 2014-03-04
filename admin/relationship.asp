<!--#include virtual="/admin/templatetop.asp"-->
<%
	Set objRelationship = New Relationship

	objRelationship.ID = Request.QueryString("id")
	objRelationship.PostID = Request.QueryString("postid")
	objRelationship.OptionTemplate = "<option value=""[id]"">[name]</option>"
	objRelationship.OptionTemplateSelected = "<option selected=""selected"" value=""[id]"">[name]</option>"
	Call objRelationship.GetAdminRelationship()

	If objRelationship.ID <> 0 Then
%>
	<h2>Admin | Edit Relationship</h2>
	<p>Edit the fields below as necessary, then click save to update the relationship:</p>
<%Else%>
	<h2>Admin | Add Relationship</h2>
	<%If objRelationship.PostID <> 0 Then%>
	    <p>Choose a Category and/or Tag to associate with your Post <em>(you can add additional Relationships afterwards)</em>:</p>
	<%Else%>
	    <p>Fill in all fields below and click save to add your relationship:</p>
	<%End If%>
<%End If%>

<form method="post" action="/admin/saverelationship.asp">
	<table border="0" cellpadding="0" cellspacing="0" width="90%">
		<input type="hidden" name="id" value="<%=objRelationship.ID%>" />
		<tr>
			<td>Category</td>
			<td>
				<select name="categoryid">
					<option value="0"></option>
					<%Call objRelationship.GetCategories()%>
				</select>
			</td>
		</tr>
		<tr>
			<td>Tag</td>
			<td>
				<select name="tagid">
					<option value="0"></option>
					<%Call objRelationship.GetTags()%>
				</select>
			</td>
		</tr>
		<tr>
			<td>Link</td>
			<td>
				<select name="linkid">
					<option value="0"></option>
					<%Call objRelationship.GetLinks()%>
				</select>
			</td>
		</tr>
		<tr>
			<td>Post</td>
			<td>
				<select name="postid">
					<option value="0"></option>
					<%Call objRelationship.GetPosts()%>
				</select>
			</td>
		</tr>
		<tr>
			<td>User</td>
			<td>
				<select name="userid">
					<option value="0"></option>
					<%Call objRelationship.GetUsers()%>
				</select>
			</td>
		</tr>
		<tr>
			<td>Order</td>
			<td><input type="text" name="order" value="<%=objRelationship.Order%>" /></td>
		</tr>
		<tr>
			<td colspan="2"><input type="submit" name="submit" value="Save" /></td>
		</tr>
	</table>
</form>

<p><strong>Please Note</strong>: Only one relationship should be defined at a time, for example - between a Category and a Link, or a Post and a User.</p>

<%
	Set objRelationship = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->