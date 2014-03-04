<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Save Tag</h2>
<%
	Set objTag = New Tag
	
	objTag.ID = Request.Form("id")
	objTag.Name = Request.Form("name")

	objTag.Save()

	Set objTag = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->