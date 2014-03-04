<!--Begin: Comments Form-->
<%
Select Case objPost.URL
    Case "contact","send-message","search"
        'Exclude comments on certain pages/posts
    Case Else
%>
<hr />
<%
    strCaptcha = objCaptcha.Generate(50, 150) 'height,width
    objComment.PostID = objPost.ID
    Response.Write("<p id=""comments"">There are " & objComment.Count & " comments:</p>")

    If Request.Form("postid") <> "" Then
        Dim strResult
        
        'Set defaults
        If objCaptcha.Process <>  "pass" Then
            strResult = Application("CaptchaFailed")
            Response.Write("<p class=""error"">" & strResult & "</p>")
        Else
            objComment.PostID = Request.Form("postid")
            objComment.UserID = IIf(Session("UserID") = "", 0, Session("UserID"))
            objComment.Name = IIf(Request.Form("name") = "", "Anonymous", Request.Form("name"))
            objComment.Email = IIf(Request.Form("email") = "", "n/a", Request.Form("email"))
            objComment.URL = IIf(Request.Form("website") = "", "n/a", Request.Form("website"))
            objComment.Content = IIf(Request.Form("comment") = "", "n/a", Request.Form("comment"))
            objComment.Published = IIf(Application("ModerateComments") = 1, IIf(Session("LoggedOn") = true, 1, 0), 1)
            objComment.IP = Request.ServerVariables("REMOTE_ADDR")
            objComment.HTTP_USER_AGENT = Request.ServerVariables("HTTP_USER_AGENT")

            strResult = objComment.Save
            Response.Write("<p class=""message"">" & strResult & "</p>")
        End If
        
        'set Session so user can re-use their entered data
        Session("Name") = Request.Form("name")
        Session("Email") = Request.Form("email")
        Session("URL") = Request.Form("website")
        Session("Comment") = Request.Form("comment")
    End If

    objComment.CommentsTemplate = "<div class=""[class]"" id=""[permalink]""><p>[name] said:</p>[content]<p>[date] | <a href=""#[permalink]"">Permalink</a> [admin]</p></div>" & vbCrLf
    objComment.EditTemplate = "| <a href=""/admin/comment.asp?id=[id]"">Edit</a> | <a href=""/admin/deletecomment.asp?id=[id]"">Delete</a>"
    Call objComment.GetComments()
%>
<p><a href="javascript:Toggle('comments-form');">Leave a comment</a></p>
<form method="post" action="/<%=objPost.Url%>/#comments" id="comments-form" style="display: none;" onsubmit="return ValidateComment()">
	<p>
	    <input type="hidden" name="postid" value="<%=objPost.ID%>" />
	    <label for="name"><em>Name (required):</em></label><br />
	    <input type="text" name="name" id="name" maxlength="50" size="50" value="<%=Session("Name")%>" />
    </p>
	<p><label for="email"><em>Email (required, not published):</em></label><br /><input type="text" name="email" id="email" maxlength="255" size="50" value="<%=Session("Email")%>" /></p>
	<p><label for="website"><em>Website (optional):</em></label><br /><input type="text" name="website" id="website" maxlength="255" size="50" value="<%=Session("URL")%>" /></p>
	<p><small>XHTML: You can use these tags: <em>&lt;a href="" title=""&gt; &lt;abbr title=""&gt; &lt;acronym title=""&gt; &lt;blockquote cite=""&gt; &lt;cite&gt; &lt;code&gt; &lt;del datetime=""&gt; &lt;em&gt; &lt;q cite=""&gt; &lt;strike&gt; &lt;strong&gt;</em></small></p>
	<p><label for="comment"><em>Your comments (required):</em></label><br /><textarea name="comment" id="comment" rows="7" cols="75"><%=Session("Comment")%></textarea></p>
	<!--Begin: Captcha-->
	<p>
        <img src="http://www.opencaptcha.com/img/<%=strCaptcha%>" alt="opencaptcha" /><br />
        <label for="answer"><em>Verification (required, for spam protection):</em></label><br />
        <input type="hidden" name="opencaptcha" value="<%=strCaptcha%>" />
        <input type="text" name="answer" id="answer" /> <small>(please enter the text as you see it in the image above, if present)</small>
	</p>
	<!--End: Captcha-->
	<p><input type="submit" value="Submit Your Comments" /> <small>(all comments are currently moderated and will appear once approved by us)</small></p>
	<p class="error" id="comments-error" style="display: none;"></p>
</form>
<%
End Select
%>
<!--End: Comments Form-->