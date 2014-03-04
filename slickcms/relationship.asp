<%
Class Relationship
	'public properties - let (set)
		Public Property Let Template(p_Template)
			m_Template = p_Template
		End Property
		
		Public Property Let ID(p_ID)
			Set m_objClean = New Clean
			m_objClean.Data = p_ID
			Call m_objClean.Numeric()
			m_ID = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let CategoryID(p_CategoryID)
			Set m_objClean = New Clean
			m_objClean.Data = p_CategoryID
			Call m_objClean.Numeric()
			m_CategoryID = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let LinkID(p_LinkID)
			Set m_objClean = New Clean
			m_objClean.Data = p_LinkID
			Call m_objClean.Numeric()
			m_LinkID = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let PostID(p_PostID)
			Set m_objClean = New Clean
			m_objClean.Data = p_PostID
			Call m_objClean.Numeric()
			m_PostID = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let UserID(p_UserID)
			Set m_objClean = New Clean
			m_objClean.Data = p_UserID
			Call m_objClean.Numeric()
			m_UserID = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let TagID(p_TagID)
			Set m_objClean = New Clean
			m_objClean.Data = p_TagID
			Call m_objClean.Numeric()
			m_TagID = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Order(p_Order)
			Set m_objClean = New Clean
			m_objClean.Data = p_Order
			Call m_objClean.Numeric()
			m_Order = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let OptionTemplate(p_OptionTemplate)
			m_OptionTemplate = p_OptionTemplate
		End Property
		
		Public Property Let OptionTemplateSelected(p_OptionTemplateSelected)
			m_OptionTemplateSelected = p_OptionTemplateSelected
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
		
		Public Property Get CategoryID()
			CategoryID = m_CategoryID
		End Property
		
		Public Property Get LinkID()
			LinkID = m_LinkID
		End Property
		
		Public Property Get PostID()
			PostID = m_PostID
		End Property
		
		Public Property Get UserID()
			UserID = m_UserID
		End Property
		
		Public Property Get TagID()
			TagID = m_TagID
		End Property
		
		Public Property Get Order()
			Order = m_Order
		End Property
		
		Public Property Get Pagination()
			Pagination = m_Pagination
		End Property
		
		Public Property Get AdminRelationshipsCount()
		    AdminRelationshipsCount = m_AdminRelationshipsCount
        End Property
		
	'private properties
		Private m_strSQL
		Private m_objRS
		Private m_objCmd
		Private m_objClean
		
		Private m_ID
		Private m_CategoryID
		Private m_LinkID
		Private m_PostID
		Private m_UserID
		Private m_TagID
		Private m_Order
		Private m_Template
		Private m_OptionTemplate
		Private m_OptionTemplateSelected
		
		'pagination variables
		Private m_Start
		Private m_End
        Private m_Pagination
        Private m_AdminRelationshipsCount

	'public methods
		Public Sub GetAdminRelationships()
			Dim strTemplate
			Dim strPublished
			
			'count the relationships for pagination use
			m_strSQL = "Select Count(*) From dbo.Relationships"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			If Not m_objRS.BOF Then
			    m_AdminRelationshipsCount = m_objRS.Fields(0).Value
            Else
                m_AdminRelationshipsCount = 0
            End If
            If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
            'pagination
			If m_Pagination = 0 Then m_Pagination = 1
			m_End = (Application("AdminPagination") * m_Pagination)
			m_Start = (m_End - Application("AdminPagination"))+1
			If m_Start = 0 Then m_Start = 1

			m_strSQL = "Execute Admin_SelectRelationships " & m_Start & "," & m_End			
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTemplate = m_Template
					
					strTemplate = Replace(strTemplate,"[relationshipid]",m_objRS.Fields("RelationshipID").Value)
					strTemplate = Replace(strTemplate,"[categoryid]",m_objRS.Fields("CategoryID").Value)
					strTemplate = Replace(strTemplate,"[categoryname]",m_objRS.Fields("CategoryName").Value)
					strTemplate = Replace(strTemplate,"[linkid]",m_objRS.Fields("LinkID").Value)
					strTemplate = Replace(strTemplate,"[linkname]",m_objRS.Fields("LinkName").Value)
					strTemplate = Replace(strTemplate,"[postid]",m_objRS.Fields("PostID").Value)
					strTemplate = Replace(strTemplate,"[posttitle]",m_objRS.Fields("PostTitle").Value)
					strTemplate = Replace(strTemplate,"[userid]",m_objRS.Fields("UserID").Value)
					strTemplate = Replace(strTemplate,"[username]",m_objRS.Fields("UserName").Value)
					strTemplate = Replace(strTemplate,"[tagid]",m_objRS.Fields("TagID").Value)
					strTemplate = Replace(strTemplate,"[tagname]",m_objRS.Fields("TagName").Value)
					strTemplate = Replace(strTemplate,"[order]",m_objRS.Fields("Order").Value)
					
					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Sub GetAdminRelationship()
			If m_ID <> 0 Then
				m_strSQL = "Execute Admin_SelectRelationship " & m_ID

				Set m_objRS = Server.CreateObject("ADODB.RecordSet")
				m_objRS.Open m_strSQL,objConn,0,1
				
				If Not m_objRS.BOF Then
					m_ID = m_objRS.Fields("RelationshipID").Value
					m_CategoryID = m_objRS.Fields("CategoryID").Value
					m_LinkID = m_objRS.Fields("LinkID").Value
					m_PostID = m_objRS.Fields("PostID").Value
					m_UserID = m_objRS.Fields("UserID").Value
					m_TagID = m_objRS.Fields("TagID").Value
					m_Order = m_objRS.Fields("Order").Value
				Else
					m_ID = 0
				End If
				
				If m_objRS.State <> 0 Then m_objRS.Close
				Set m_objRS = Nothing
			End If
			
			If m_ID = 0 Then
				m_ID = 0
				m_CategoryID = 0
				m_LinkID = 0
				m_UserID = 0
				m_TagID = 0
				m_Order = 0
			End If
		End Sub
		
		Public Sub GetCategories()
		    'used within the Admin
			Dim intID, strName
			m_strSQL = "Select CategoryID, Name From dbo.Categories Order By Name Asc"
			
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					intID = m_objRS.Fields("CategoryID").Value
					strName = m_objRS.Fields("Name").Value

					If m_CategoryID = intID Then
						Response.Write(Replace(Replace(m_OptionTemplateSelected,"[id]",intID),"[name]",strName))
					Else
						Response.Write(Replace(Replace(m_OptionTemplate,"[id]",intID),"[name]",strName))
					End If
				
					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Sub GetTags()
		    'used within the Admin
			Dim intID, strName
			m_strSQL = "Select TagID, Name From dbo.Tags Order By Name Asc"
			
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					intID = m_objRS.Fields("TagID").Value
					strName = m_objRS.Fields("Name").Value

					If m_TagID = intID Then
						Response.Write(Replace(Replace(m_OptionTemplateSelected,"[id]",intID),"[name]",strName))
					Else
						Response.Write(Replace(Replace(m_OptionTemplate,"[id]",intID),"[name]",strName))
					End If
				
					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Sub GetLinks()
		    'used within the Admin
			Dim intID, strName
			m_strSQL = "Select LinkID, Name From dbo.Links Order By Name Asc"
			
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					intID = m_objRS.Fields("LinkID").Value
					strName = m_objRS.Fields("Name").Value

					If m_LinkID = intID Then
						Response.Write(Replace(Replace(m_OptionTemplateSelected,"[id]",intID),"[name]",strName))
					Else
						Response.Write(Replace(Replace(m_OptionTemplate,"[id]",intID),"[name]",strName))
					End If
				
					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Sub GetPosts()
		    'used within the Admin
			Dim intID, strName
			m_strSQL = "Select PostID, Title From dbo.Posts Order By Title Asc"
			
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					intID = m_objRS.Fields("PostID").Value
					strName = m_objRS.Fields("Title").Value

					If m_PostID = intID Then
						Response.Write(Replace(Replace(m_OptionTemplateSelected,"[id]",intID),"[name]",strName))
					Else
						Response.Write(Replace(Replace(m_OptionTemplate,"[id]",intID),"[name]",strName))
					End If
				
					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Sub GetUsers()
		    'used within the Admin
			Dim intID, strName
			m_strSQL = "Select UserID, Name From dbo.Users Order By Name Asc"
			
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					intID = m_objRS.Fields("UserID").Value
					strName = m_objRS.Fields("Name").Value

					If m_UserID = intID Then
						Response.Write(Replace(Replace(m_OptionTemplateSelected,"[id]",intID),"[name]",strName))
					Else
						Response.Write(Replace(Replace(m_OptionTemplate,"[id]",intID),"[name]",strName))
					End If
				
					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub

		Public Sub Save()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			If m_ID = 0 Then
				m_objCmd.CommandText = "Admin_InsertRelationship"
			Else
				m_objCmd.CommandText = "Admin_UpdateRelationship"
				m_objCmd.Parameters.Append m_objCmd.CreateParameter("@RelationshipID", 3, 1, , m_ID)
			End If

			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@CategoryID", 3, 1, , m_CategoryID)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@LinkID", 3, 1, , m_LinkID)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@PostID", 3, 1, , m_PostID)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@UserID", 3, 1, , m_UserID)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@TagID", 3, 1, , m_TagID)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Order", 3, 1, , m_Order)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			If m_ID = 0 Then
				Response.Write(Application("RelationshipSaved"))
			Else
				Response.Write(Application("RelationshipUpdated"))
			End If
		End Sub

		Public Sub Delete()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			m_objCmd.CommandText = "Admin_DeleteRelationship"
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@RelationshipID", 3, 1, , m_ID)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			Response.Write(Application("RelationshipDeleted"))
		End Sub

	'private methods
		'none
End Class
%>