<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Save Link</h2>
<%
	Set objLink = New Link
	
	objLink.ID = Request.Form("id")
	objLink.Name = Request.Form("name")
	objLink.URL = Request.Form("url")
	objLink.Description = Request.Form("description")
	objLink.Published = Request.Form("published")

	objLink.Save()

	Set objLink = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->