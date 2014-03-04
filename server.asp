<%
    'used for writing out Server variables in Debug mode
    If Application("Debug") = 1 Then
	    Dim i
    	
	    Response.Write "<table border=""1"">"
	    Response.Write "<tr><td><strong>Server variables</strong></td><td>&nbsp;</td></tr>"
    	
	    For Each i in Request.ServerVariables
		    Response.Write "<tr>"
		    Response.Write "<td>" & i & "</td>"
		    Response.Write "<td>" & Request.ServerVariables(i) & "&nbsp;</td>"
		    Response.Write "</tr>"
	    Next
    	
	    Response.Write "</table>"
	End If
%>