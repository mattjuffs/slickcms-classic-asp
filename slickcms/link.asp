<%
Class Link
	'public properties - let (set)
		Public Property Let Template(p_Template)
			m_Template = p_Template
		End Property
	
		Public Property Let CategoryID(p_CategoryID)
			Set m_objClean = New Clean
			m_objClean.Data = p_CategoryID
			Call m_objClean.Numeric()
			m_CategoryID = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
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
			Call m_objClean.AlphaNumeric()
			Call m_objClean.MaxLength(255)
			m_Name = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let URL(p_URL)
			Set m_objClean = New Clean
			m_objClean.Data = p_URL
			Call m_objClean.MaxLength(1024)
			m_URL = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Description(p_Description)
			m_Description = p_Description
		End Property
		
		Public Property Let Published(p_Published)
			Set m_objClean = New Clean
			m_objClean.Data = p_Published
			Call m_objClean.Numeric()
			m_Published = CInt(m_objClean.Data)
			Set m_objClean = Nothing
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
		
		Public Property Get URL()
			URL = m_URL
		End Property
		
		Public Property Get Description
			Description = m_Description
		End Property

		Public Property Get DateCreated()
			DateCreated = m_DateCreated
		End Property
		
		Public Property Get DateModified()
			DateModified = m_DateModified
		End Property
		
		Public Property Get Published()
			Published = m_Published
		End Property
		
		Public Property Get Pagination()
			Pagination = m_Pagination
		End Property
		
		Public Property Get AdminLinksCount()
		    AdminLinksCount = m_AdminLinksCount
        End Property

	'private properties
		Private m_strSQL
		Private m_objRS
		Private m_objCmd
		Private m_objClean
		
		Private m_ID
		Private m_Name
		Private m_URL
		Private m_Description
		Private m_DateCreated
		Private m_DateModified
		Private m_Published
		Private m_Template
		Private m_CategoryID
		
		'pagination variables
		Private m_Start
		Private m_End
        Private m_Pagination
        Private m_AdminLinksCount

	'public methods
		Public Sub GetLinks()
		    'retrieves a list of Links for navigation use
			Dim strTemplate
			Dim strName, strURL, strDescription
	
			m_strSQL = "Execute SelectLinks " & m_CategoryID
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strName = m_objRS.Fields("Name").Value
					strURL = m_objRS.Fields("URL").Value
					strDescription = m_objRS.Fields("Description").Value
					
					strTemplate = m_Template
					strTemplate = Replace(strTemplate,"[name]",strName)
					strTemplate = Replace(strTemplate,"[url]",strURL)
					strTemplate = Replace(strTemplate,"[description]",strDescription)
	
					Response.Write(strTemplate & vbCrLf)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Sub GetAdminLinks()
			Dim strTemplate
			Dim strPublished
			
			'count the links for pagination use
			m_strSQL = "Select Count(*) From dbo.Links"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			If Not m_objRS.BOF Then
			    m_AdminLinksCount = m_objRS.Fields(0).Value
            Else
                m_AdminLinksCount = 0
            End If
            If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
            'pagination
			If m_Pagination = 0 Then m_Pagination = 1
			m_End = (Application("AdminPagination") * m_Pagination)
			m_Start = (m_End - Application("AdminPagination"))+1
			If m_Start = 0 Then m_Start = 1

			m_strSQL = "Execute Admin_SelectLinks " & m_Start & "," & m_End
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTemplate = m_Template
				
					m_ID = m_objRS.Fields("LinkID").Value
					m_Name = m_objRS.Fields("Name").Value
					m_URL = m_objRS.Fields("URL").Value
					m_DateModified = m_objRS.Fields("DateModified").Value
					m_Published = m_objRS.Fields("Published").Value
					
					Select Case m_Published
						Case 0
							strPublished = "No"
						Case 1
							strPublished = "Yes"
						Case Else
							strPublished = "No"
					End Select
					
					strTemplate = Replace(strTemplate,"[linkid]",m_ID)
					strTemplate = Replace(strTemplate,"[name]",m_Name)
					strTemplate = Replace(strTemplate,"[url]",m_URL)
					strTemplate = Replace(strTemplate,"[datemodified]",m_DateModified)
					strTemplate = Replace(strTemplate,"[published]",strPublished)
					
					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Sub GetAdminLink()
			If m_ID <> 0 Then
				m_strSQL = "Execute Admin_SelectLink " & m_ID

				Set m_objRS = Server.CreateObject("ADODB.RecordSet")
				m_objRS.Open m_strSQL,objConn,0,1
				
				If Not m_objRS.BOF Then
					m_ID = m_objRS.Fields("LinkID").Value
					m_Name = m_objRS.Fields("Name").Value
					m_URL = m_objRS.Fields("URL").Value
					m_Description = m_objRS.Fields("Description").Value
					m_DateCreated = m_objRS.Fields("DateCreated").Value
					m_DateModified = m_objRS.Fields("DateModified").Value
					m_Published = m_objRS.Fields("Published").Value
				Else
					m_ID = 0
				End If
				
				If m_objRS.State <> 0 Then m_objRS.Close
				Set m_objRS = Nothing
			End If
			
			If m_ID = 0 Then
				m_Name = ""
				m_URL = ""
				m_Description = " " '#442
				m_DateCreated = ""
				m_DateModified = ""
				m_Published = ""
			End If
		End Sub
		
		Public Sub Save()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			If m_ID = 0 Then
				m_objCmd.CommandText = "Admin_InsertLink"
			Else
				m_objCmd.CommandText = "Admin_UpdateLink"
				m_objCmd.Parameters.Append m_objCmd.CreateParameter("@LinkID", 3, 1, , m_ID)
			End If

            If len(m_Description) = 0 Then m_Description = " " '#442

			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Name", 200, 1, 255, m_Name)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@URL", 200, 1, 1024, m_URL)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Description", 200, 1, len(m_Description), m_Description)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Published", 3, 1, , m_Published)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			If m_ID = 0 Then
				Response.Write(Application("LinkSaved"))
			Else
				Response.Write(Application("LinkUpdated"))
			End If
		End Sub
		
		Public Sub Delete()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			m_objCmd.CommandText = "Admin_DeleteLink"
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@LinkID", 3, 1, , m_ID)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			Response.Write(Application("LinkDeleted"))
		End Sub

	'private methods
		'none
End Class
%>