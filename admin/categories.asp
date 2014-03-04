<!--#include virtual="/admin/templatetop.asp"-->
<%
    Set objCategory = New Category
    
    If Request.QueryString("page") <> "" Then
        objCategory.Pagination = Request.QueryString("page")
    End If
%>
<h2>Admin | Categories</h2>
<p><a href="/admin/category.asp">Add a new category</a>, or click on <em>edit</em> below to amend an existing one</p>
<table border="0" cellpadding="2px" cellspacing="0" class="records">
	<tr class="headings">
		<td>Name</td>
		<td>Type</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
<%
	strTemplate = "<tr>" & vbCrLf
	strTemplate = strTemplate & "<td>[name]</td>" & vbCrLf
	strTemplate = strTemplate & "<td>[type]</td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/category.asp?id=[categoryid]"">edit</a></td>" & vbCrLf
	strTemplate = strTemplate & "<td><a href=""/admin/deletecategory.asp?id=[categoryid]"">delete</a></td>" & vbCrLf
	strTemplate = strTemplate & "</tr>"

	objCategory.Template = strTemplate
	Call objCategory.GetAdminCategories()
%>
</table>
<%=objSlickCMS.AdminPaging(objCategory.AdminCategoriesCount,objCategory.Pagination,"categories.asp")%>
<!--#include virtual="/admin/templatebottom.asp"-->
<%
    Set objCategory = Nothing
%>