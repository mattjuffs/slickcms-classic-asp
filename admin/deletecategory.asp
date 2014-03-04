<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Delete Category</h2>
<%
	Set objCategory = New Category
	objCategory.ID = Request.QueryString("id")
	objCategory.Delete()
	Set objCategory = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->