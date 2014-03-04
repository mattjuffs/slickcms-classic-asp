	</div>
	<div id="footer">
        <p>
            <a href="http://slickcms.slickhouse.com/"><img src="/images/powered_by_slickcms.png" alt="slickcms logo" title="Powered by slickcms v<%=Application("SlickCMS_Version")%>" /></a><br />
            <a href="http://www.microsoft.com/"><img src="<%=Application("CDN")%>images/microsoft.png" alt="Microsoft logos" title="SlickCMS is built using Microsoft Technologies" /></a>
        </p>
	</div>
</div>

</body>
</html>
<%
	Call CloseDatabase()
	
	Set objSlickCMS = Nothing
%>