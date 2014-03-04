<div id="header">
    <%
        Select Case objPost.UrlType
            Case "post","date" 'single item
    %>
	<h2><a href="/"><img src="/images/slickcms_logo_2010.png" alt="slickcms logo" title="<%=Application("SiteName")%>" /></a></h2>
	<%
	        Case Else
	%>
	<h1><a href="/"><img src="/images/slickcms_logo_2010.png" alt="slickcms logo" title="<%=Application("SiteName")%>" /></a></h1>
	<%
	    End Select
	%>
	<form id="search" method="post" action="/search/">
		<p>
		    <input type="text" name="keywords" id="keywords" maxlength="100" />
		    <input type="submit" value="Search" />
        </p>
	</form>
</div>