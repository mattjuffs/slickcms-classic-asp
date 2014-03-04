<%
Class Statistic
	'public properties - let (set)
	    'none

	'public properties - get (retrieve)
		'none
		
	'private properties
		Private m_strSQL
		Private m_objRS

	'public methods
		Public Function BlogStats(strTemplate)
			m_strSQL = "Execute [Statistics]"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1

			If Not m_objRS.BOF Then
			    'stats from DB			
				strTemplate = Replace(strTemplate,"[posts]",m_objRS.Fields("Posts").Value)
				strTemplate = Replace(strTemplate,"[pages]",m_objRS.Fields("Pages").Value)
				strTemplate = Replace(strTemplate,"[comments]",m_objRS.Fields("Comments").Value)
				strTemplate = Replace(strTemplate,"[categories]",m_objRS.Fields("Categories").Value)
				strTemplate = Replace(strTemplate,"[tags]",m_objRS.Fields("Tags").Value)
				strTemplate = Replace(strTemplate,"[links]",m_objRS.Fields("Links").Value)
				strTemplate = Replace(strTemplate,"[users]",m_objRS.Fields("Users").Value)
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
			'stats from Application
			strTemplate = Replace(strTemplate,"[activevisitors]",Application("ActiveVisitors"))
			strTemplate = Replace(strTemplate,"[totalvisitors]",Application("TotalVisitors"))
			
			BlogStats = strTemplate
		End Function

	'private methods
		'none
End Class
%>