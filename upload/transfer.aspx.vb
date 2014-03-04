Partial Class upload_transfer
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Request.Form("LoggedOn") Is Nothing Then
            Session("LoggedOn") = False
        Else
            Session("LoggedOn") = Request.Form("LoggedOn")
        End If

        If Session("LoggedOn") = "" Then Session("LoggedOn") = False
		Response.Redirect("/upload/upload.aspx")
    End Sub
End Class