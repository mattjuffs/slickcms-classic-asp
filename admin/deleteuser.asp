<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Delete User</h2>
<%
	Set objUser = New User
	objUser.ID = Request.QueryString("id")
	objUser.Delete()
	Set objUser = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->