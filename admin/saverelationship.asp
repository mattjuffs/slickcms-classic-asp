<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Save Relationship</h2>
<%
	Set objRelationship = New Relationship
	
	objRelationship.ID = Request.Form("id")
	objRelationship.CategoryID = Request.Form("categoryid")
	objRelationship.LinkID = Request.Form("linkid")
	objRelationship.PostID = Request.Form("postid")
	objRelationship.UserID = Request.Form("userid")
	objRelationship.TagID = Request.Form("tagid")
	objRelationship.Order = Request.Form("order")

	objRelationship.Save()

	Set objRelationship = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->