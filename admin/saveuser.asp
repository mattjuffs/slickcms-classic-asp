<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Save User</h2>
<%
	If Request.Form("password") <> Request.Form("confirmpassword") Then
		Session("Error") = "Your passwords don't match!"
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
		Response.End
	Else
		Set objUser = New User
		
		objUser.ID = Request.Form("id")
		objUser.Name = Request.Form("name")
		objUser.Email = Request.Form("email")
		objUser.Password = Request.Form("password")
		objUser.URL = Request.Form("url")
		objUser.IP = Request.ServerVariables("REMOTE_ADDR")
		objUser.Biography = Request.Form("wmd-input")
		objUser.Active = Request.Form("active")
		objUser.LoginFails = Request.Form("loginfails")
	
		objUser.Save()
	
		Set objUser = Nothing
	End If
%>
<!--#include virtual="/admin/templatebottom.asp"-->