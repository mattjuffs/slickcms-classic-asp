<%
    'used for writing out Session variables in Debug mode
    If Application("Debug") = 1 Then
	    Dim i
	    Dim j

	    j = Session.Contents.Count
    	
	    Response.Write "<table border=""1"">"
	    Response.Write "<tr><td><strong>Session variables</strong></td><td><strong>" & j & "</strong></td></tr>"
    	
	    For Each i in Session.Contents
		    Response.Write "<tr>"
		    Response.Write "<td>" & i & "</td>"
		    Response.Write "<td>" & Session.Contents(i) & "&nbsp;</td>"
		    Response.Write "</tr>"
	    Next
    	
	    Response.Write "</table>"
	End If
%>