<!--Begin: Contact Form-->
<%
    strCaptcha = objCaptcha.Generate(50, 150) 'height,width

	If Session("EmailSent") <> "" Then
		If Session("EmailSent") = true Then
%>
<p class="message"><%=Application("EmailSent")%></p>
<%
		Else
%>
<p class="error"><%=Application("CaptchaFailed")%></p>
<%
		End If
	End If
	
	Session("EmailSent") = Empty
%>
<form method="post" action="/send-message/" id="contact">
	<p><em>Name (required):</em><br /><input type="text" name="name" id="name" maxlength="50" size="50" value="<%=Session("Name")%>" /></p>
	<p><em>Email (required, for replies):</em><br /><input type="text" name="email" id="email" maxlength="255" size="50" value="<%=Session("Email")%>" /></p>
	<p><em>Your message (required):</em><br /><textarea name="message" id="message" rows="7" cols="75"><%=Session("Comment")%></textarea></p>
	<!--Begin: Captcha-->
	<p>
        <img src="http://www.opencaptcha.com/img/<%=strCaptcha%>" alt="opencaptcha" /><br />
        <label for="answer"><em>Verification (required, for spam protection):</em></label><br />
        <input type="hidden" name="opencaptcha" value="<%=strCaptcha%>" />
        <input type="text" name="answer" /> <small>(please enter the text as you see it in the image above)</small>
	</p>
	<!--End: Captcha-->
	<p><input type="submit" value="Send Your Message" /></p>
</form>
<!--End: Contact Form-->