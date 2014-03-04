<%
Class Tag
	'public properties - let (set)
		Public Property Let ID(p_ID)
			Set m_objClean = New Clean
			m_objClean.Data = p_ID
			Call m_objClean.Numeric()
			m_ID = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Name(p_Name)
			Set m_objClean = New Clean
			m_objClean.Data = p_Name
			Call m_objClean.Url()
			Call m_objClean.Maxlength(255)
			m_Name = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Template(p_Template)
			m_Template = p_Template
		End Property
		
		Public Property Let Pagination(p_Pagination)
			Set m_objClean = New Clean
			m_objClean.Data = p_Pagination
			Call m_objClean.Numeric()
			m_Pagination = CInt(m_objClean.Data)
			Set m_objClean = Nothing
			If m_Pagination = "" Then m_Pagination = 0
		End Property

	'public properties - get (retrieve)
		Public Property Get ID()
			ID = m_ID
		End Property
		
		Public Property Get Name()
			Name = m_Name
		End Property
		
		Public Property Get Pagination()
			Pagination = m_Pagination
		End Property
		
		Public Property Get AdminTagsCount()
		    AdminTagsCount = m_AdminTagsCount
        End Property
		
	'private properties
		Private m_strSQL
		Private m_objRS
		Private m_objCmd
		Private m_objClean

		Private m_ID
		Private m_Name
		Private m_Template
		Private m_Url
		
		'pagination variables
		Private m_Start
		Private m_End
        Private m_Pagination
        Private m_AdminTagsCount

	'public methods
		Public Sub GetAdminTags()
			Dim strTemplate
			
			'count the tags for pagination use
			m_strSQL = "Select Count(*) From dbo.Tags"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			If Not m_objRS.BOF Then
			    m_AdminTagsCount = m_objRS.Fields(0).Value
            Else
                m_AdminTagsCount = 0
            End If
            If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
            'pagination
			If m_Pagination = 0 Then m_Pagination = 1
			m_End = (Application("AdminPagination") * m_Pagination)
			m_Start = (m_End - Application("AdminPagination"))+1
			If m_Start = 0 Then m_Start = 1

			m_strSQL = "Execute Admin_SelectTags " & m_Start & "," & m_End
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTemplate = m_Template
				
					m_ID = m_objRS.Fields("TagID").Value
					m_Name = m_objRS.Fields("Name").Value
					
					strTemplate = Replace(strTemplate,"[tagid]",m_ID)
					strTemplate = Replace(strTemplate,"[name]",m_Name)

					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub

		Public Sub GetAdminTag()
			If m_ID <> 0 Then
				m_strSQL = "Execute Admin_SelectTag " & m_ID

				Set m_objRS = Server.CreateObject("ADODB.RecordSet")
				m_objRS.Open m_strSQL,objConn,0,1
				
				If Not m_objRS.BOF Then
					m_ID = m_objRS.Fields("TagID").Value
					m_Name = m_objRS.Fields("Name").Value
				Else
					m_ID = 0
				End If
				
				If m_objRS.State <> 0 Then m_objRS.Close
				Set m_objRS = Nothing
			End If
			
			If m_ID = 0 Then
				m_Name = ""
			End If
		End Sub
		
		Public Sub Save()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			If m_ID = 0 Then
				m_objCmd.CommandText = "Admin_InsertTag"
			Else
				m_objCmd.CommandText = "Admin_UpdateTag"
				m_objCmd.Parameters.Append m_objCmd.CreateParameter("@LinkID", 3, 1, , m_ID)
			End If

			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Name", 200, 1, 255, m_Name)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			If m_ID = 0 Then
				Response.Write(Application("TagSaved"))
			Else
				Response.Write(Application("TagUpdated"))
			End If
		End Sub
		
		Public Sub Delete()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			m_objCmd.CommandText = "Admin_DeleteTag"
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@TagID", 3, 1, , m_ID)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			Response.Write(Application("TagDeleted"))
		End Sub
		
		Public Sub Cloud()
		    'tag cloud - a selection of the most popular tags grouped together
		    Dim strTemplate

            m_strSQL = "Execute SelectTags"

			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTemplate = m_Template

					m_ID = m_objRS.Fields("TagID").Value
					m_Url = FormatUrl(m_objRS.Fields("Name").Value,"output")
					m_Name = m_objRS.Fields("Name").Value
					
					strTemplate = Replace(strTemplate,"[tagid]",m_ID)
					strTemplate = Replace(strTemplate,"[name]",m_Name)
					strTemplate = Replace(strTemplate,"[url]",m_Url)
					strTemplate = Replace(strTemplate,"[postcount]",m_objRS.Fields("PostCount").Value)
					
					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Function GetPostTags(intPostID)
		    'retrieve comma separated list of Tags for a Post
		    Dim strReturn
		    Dim strTemplate
		    
            m_strSQL = "Execute SelectPostTags " & intPostID

			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTemplate = m_Template

					m_Url = FormatUrl(m_objRS.Fields("Name").Value,"output")
					m_Name = m_objRS.Fields("Name").Value

					strTemplate = Replace(strTemplate,"[name]",m_Name)
					strTemplate = Replace(strTemplate,"[url]",m_Url)
					
					'concatenate as strTemplate is to be returned
					strReturn = (strReturn & strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		    
		    GetPostTags = strReturn
		End Function

	'private methods
		Public Function FormatUrl(strUrl,strType)
			'builds up a Url as per the Urls Specification
			If strUrl = "" Then strUrl = "home"
			strUrl = Replace(strUrl," ","-")
			strUrl = lcase(strUrl)
			
			Set m_objClean = New Clean
			m_objClean.Data = strUrl
			Call m_objClean.Url()
			Call m_objClean.MaxLength(255)
			strUrl = m_objClean.Data
			Set m_objClean = Nothing
			
			'format Url specifically for use
			Select Case strType
				Case "input" 'input into SlickCMS
					strUrl = strUrl
				Case "output" 'output from SlickCMS
					If strUrl = "home" Then
						strUrl = Application("SiteUrl")
					Else
   				        strUrl = Application("SiteUrl") & "tag/" & strUrl & "/"
					End If
				Case Else
					strUrl = ""
			End Select

			FormatUrl = strUrl
		End Function
End Class
%>