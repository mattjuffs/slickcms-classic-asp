<div id="footer">
	<ul>
		<%If Session("LoggedOn") = true Then%>
			<li><a href="/admin/logout.asp">Logout <%=Session("Name")%></a></li>
			<li>|</li>
		<%End If%>
		<li><a href="/admin/">Admin</a></li>
		<li>|</li>
		<li>All Content &copy; 2008 - 2010 Slickhouse</li>
		<li>|</li>
		<li><a href="/ms-rl/">Ms-RL</a></li>
		<li>|</li>
		<li><a href="/contact/">Contact SlickCMS</a></li>
		<li>|</li>
		<li><a href="<%=Application("SiteURL")%>rss2.asp?t=posts">RSS 2.0 (Posts)</a></li>
		<li>|</li>
		<li><a href="<%=Application("SiteURL")%>rss2.asp?t=comments">RSS 2.0 (Comments)</a></li>
        <li>|</li>
		<li><a href="<%=Application("SiteURL")%>sitemap.asp">XML Sitemap</a></li>
	</ul>
	<p>
        <a href="http://validator.w3.org/check?uri=referer"><img src="<%=Application("CDN")%>images/valid-xhtml11-blue.png" alt="Valid XHTML 1.1" /></a>
        <a href="http://jigsaw.w3.org/css-validator/check/referer"><img src="<%=Application("CDN")%>images/vcss-blue.gif" alt="Valid CSS" /></a>
        <a href="http://validator.w3.org/feed/check.cgi?url=http%3A//<%=Replace(Application("SiteURL"),"http://","")%>rss2.asp%3Ft%3Dposts"><img src="<%=Application("CDN")%>images/valid-rss.png" alt="Valid RSS" /></a>
	</p>
	<p><a href="http://www.microsoft.com/"><img src="<%=Application("CDN")%>images/microsoft.png" alt="Microsoft logos" title="SlickCMS is built using Microsoft Technologies" /></a></p>
</div>