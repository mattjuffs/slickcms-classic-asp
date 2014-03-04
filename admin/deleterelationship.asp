<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Delete Relationship</h2>
<%
	Set objRelationship = New Relationship
	objRelationship.ID = Request.QueryString("id")
	objRelationship.Delete()
	Set objRelationship = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->