<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Save Category</h2>
<%
	Set objCategory = New Category
	
	objCategory.ID = Request.Form("id")
	objCategory.Name = Request.Form("name")
	objCategory.CType = Request.Form("type")
	objCategory.Description = Request.Form("description")

	objCategory.Save()

	Set objCategory = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->