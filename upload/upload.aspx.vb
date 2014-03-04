Imports System.IO

Partial Class Upload
    Inherits System.Web.UI.Page
    Dim FilePath As String = Server.MapPath("~") & "\uploads\"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If FileUpload1.HasFile Then
            Try
                If CheckBox1.Checked = True Then
                    'overwrite
                    Call Upload()
                Else
                    If File.Exists(FilePath & FileUpload1.FileName) Then
                        Label1.Text = "ERROR: File already exists, please rename and try again, or choose Overwrite!"
                    Else
                        Call Upload()
                    End If
                End If
            Catch ex As Exception
                Label1.Text = "ERROR: " & ex.Message.ToString()
            End Try
        Else
            Label1.Text = "You have not specified a file."
        End If
    End Sub

    Sub Upload()
        Dim s As New StringBuilder

        FileUpload1.SaveAs(FilePath & FileUpload1.FileName)

        s.Append("<p>File name: " & FileUpload1.PostedFile.FileName & "<br />")
        s.Append("File Size: " & (FileUpload1.PostedFile.ContentLength / 1024).ToString("#.00") & " kb<br />")
        s.Append("Content type: " & FileUpload1.PostedFile.ContentType & "<br />")
        s.Append("<a href=""/uploads/" & FileUpload1.FileName & """ target=""_blank"">View File</a></p>")
        s.Append("<p>Link to your file:<br /><textarea rows=""3"" cols=""35""><a href=""/uploads/" & FileUpload1.FileName & """>Link</a></textarea></p>")
        s.Append("<p>Image of your file:<br /><textarea rows=""3"" cols=""35""><img src=""/uploads/" & FileUpload1.FileName & """ alt="""" title="""" /></textarea><small><em>(remember to specify the <strong>alt</strong> and <strong>title</strong> attributes)<em></small></p>")

        Label1.Text = s.ToString
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("LoggedOn") <> True Then
            Response.Redirect("/admin/login.asp")
            Response.End()
        End If
    End Sub
End Class
