<%
Class Category
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
		
		Public Property Let CType(p_Type)
		    'Type is a reserved word, hence the C prefix
			Set m_objClean = New Clean
			m_objClean.Data = p_Type
			Call m_objClean.AlphaNumeric()
			Call m_objClean.Maxlength(255)
			m_Type = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Description(p_Description)
			m_Description = p_Description
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
		
		Public Property Get CType()
			CType = m_Type
		End Property
		
		Public Property Get Description()
			Description = m_Description
		End Property
		
		Public Property Get Pagination()
			Pagination = m_Pagination
		End Property
		
		Public Property Get AdminCategoriesCount()
		    AdminCategoriesCount = m_AdminCategoriesCount
        End Property
		
	'private properties
		Private m_strSQL
		Private m_objRS
		Private m_objCmd
		Private m_objClean

		Private m_ID
		Private m_Name
		Private m_Type
		Private m_Description
		Private m_Template
		Private m_Url
		
		'pagination variables
		Private m_Start
		Private m_End
        Private m_Pagination
        Private m_AdminCategoriesCount

	'public methods
		Public Sub GetAdminCategories()
			Dim strTemplate
			
			'count the categories for pagination use
			m_strSQL = "Select Count(*) From dbo.Categories"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			If Not m_objRS.BOF Then
			    m_AdminCategoriesCount = m_objRS.Fields(0).Value
            Else
                m_AdminCategoriesCount = 0
            End If
            If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
            'pagination
			If m_Pagination = 0 Then m_Pagination = 1
			m_End = (Application("AdminPagination") * m_Pagination)
			m_Start = (m_End - Application("AdminPagination"))+1
			If m_Start = 0 Then m_Start = 1

			m_strSQL = "Execute Admin_SelectCategories " & m_Start & "," & m_End
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTemplate = m_Template
				
					m_ID = m_objRS.Fields("CategoryID").Value
					m_Name = m_objRS.Fields("Name").Value
					m_Type = m_objRS.Fields("Type").Value
					
					strTemplate = Replace(strTemplate,"[categoryid]",m_ID)
					strTemplate = Replace(strTemplate,"[name]",m_Name)
					strTemplate = Replace(strTemplate,"[type]",m_Type)
					
					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub

		Public Sub GetAdminCategory()
			If m_ID <> 0 Then
				m_strSQL = "Execute Admin_SelectCategory " & m_ID

				Set m_objRS = Server.CreateObject("ADODB.RecordSet")
				m_objRS.Open m_strSQL,objConn,0,1
				
				If Not m_objRS.BOF Then
					m_ID = m_objRS.Fields("CategoryID").Value
					m_Name = m_objRS.Fields("Name").Value
					m_Type = m_objRS.Fields("Type").Value
					m_Description = m_objRS.Fields("Description").Value
				Else
                    m_ID = 0
                    m_Name = ""
                    m_Type = ""
                    m_Description = ""
				End If
				
				If m_objRS.State <> 0 Then m_objRS.Close
				Set m_objRS = Nothing
			End If
		End Sub
		
		Public Sub Save()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			If m_ID = 0 Then
			    'insert new record
				m_objCmd.CommandText = "Admin_InsertCategory"
			Else
			    'update existing record
				m_objCmd.CommandText = "Admin_UpdateCategory"
				m_objCmd.Parameters.Append m_objCmd.CreateParameter("@LinkID", 3, 1, , m_ID)
			End If

			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Name", 200, 1, 255, m_Name)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Type", 200, 1, 255, m_Type)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Description", 200, 1, len(m_Description), m_Description)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			If m_ID = 0 Then
				Response.Write(Application("CategorySaved"))
			Else
				Response.Write(Application("CategoryUpdated"))
			End If
		End Sub
		
		Public Sub Delete()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			m_objCmd.CommandText = "Admin_DeleteCategory"
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@CategoryID", 3, 1, , m_ID)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			Response.Write(Application("CategoryDeleted"))
		End Sub
		
		Public Sub Categories()
		    'retrieves a list of Categories for navigation
		    Dim strTemplate

            m_strSQL = "Execute SelectCategories"

			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTemplate = m_Template

					m_ID = m_objRS.Fields("CategoryID").Value
					m_Description = m_objRS.Fields("Description").Value
					m_Url = FormatUrl(m_objRS.Fields("Name").Value,"output")
					m_Name = m_objRS.Fields("Name").Value
					
					strTemplate = Replace(strTemplate,"[categoryid]",m_ID)
					strTemplate = Replace(strTemplate,"[name]",m_Name)
					strTemplate = Replace(strTemplate,"[description]",m_Description)
					strTemplate = Replace(strTemplate,"[url]",m_Url)
					strTemplate = Replace(strTemplate,"[postcount]",m_objRS.Fields("PostCount").Value)
					
					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Function GetPostCategories(intPostID)
		    'retrieve comma separated list of Categories for a Post
		    Dim strReturn
		    Dim strTemplate
		    
            m_strSQL = "Execute SelectPostCategories " & intPostID

			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTemplate = m_Template

					m_Description = m_objRS.Fields("Description").Value
					m_Url = FormatUrl(m_objRS.Fields("Name").Value,"output")
					m_Name = m_objRS.Fields("Name").Value

					strTemplate = Replace(strTemplate,"[name]",m_Name)
					strTemplate = Replace(strTemplate,"[description]",m_Description)
					strTemplate = Replace(strTemplate,"[url]",m_Url)
					
					'concatenate as strTemplate is to be returned
					strReturn = (strReturn & strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		    
		    GetPostCategories = strReturn
		End Function

	'private methods
		Private Function FormatUrl(strUrl,strType)
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
   				        strUrl = Application("SiteUrl") & "category/" & strUrl & "/"
					End If
				Case Else
					strUrl = ""
			End Select

			FormatUrl = strUrl
		End Function
End Class
%>