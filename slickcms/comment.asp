<%
Class Comment
	'public properties - let (set)
		Public Property Let ID(p_ID)
			Set m_objClean = New Clean
			m_objClean.Data = p_ID
			Call m_objClean.Numeric()
			m_ID = CInt(m_objClean.Data)
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
        
        Public Property Let Name(p_Name)
			Set m_objClean = New Clean
			m_objClean.Data = p_Name
			Call m_objClean.AlphaNumeric()
			Call m_objClean.MaxLength(50)
			m_Name = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Email(p_Email)
			Set m_objClean = New Clean
			m_objClean.Data = p_Email
			Call m_objClean.Email()
			Call m_objClean.MaxLength(255)
			m_Email = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let URL(p_URL)
			Set m_objClean = New Clean
			m_objClean.Data = p_URL
			Call m_objClean.MaxLength(1024)
			m_URL = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Content(p_Content)
			m_Content = p_Content
		End Property
		
		Public Property Let Published(p_Published)
			Set m_objClean = New Clean
			m_objClean.Data = p_Published
			Call m_objClean.Numeric()
			m_Published = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let CommentsTemplate(p_CommentsTemplate)
			m_CommentsTemplate = p_CommentsTemplate
		End Property
		
		Public Property Let EditTemplate(p_EditTemplate)
		    m_EditTemplate = p_EditTemplate
		End Property
		
		Public Property Let IP(p_IP)			
			Set m_objClean = New Clean
			m_objClean.Data = p_IP
			Call m_objClean.NumericPlus()
			Call m_objClean.MaxLength(15)
			m_IP = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let HTTP_USER_AGENT(p_HTTP_USER_AGENT)			
			Set m_objClean = New Clean
			m_objClean.Data = p_HTTP_USER_AGENT
			Call m_objClean.MaxLength(1024)
			m_HTTP_USER_AGENT = m_objClean.Data
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
		
		Public Property Get PostID()
		    PostID = m_PostID
        End Property
        
        Public Property Get UserID()
            UserID = m_UserID
        End Property
        
        Public Property Get Name()
            Name = m_Name
        End Property
        
        Public Property Get Email()
            Email = m_Email
        End Property
        
        Public Property Get URL()
            URL = m_URL
        End Property
        
        Public Property Get Content()
            Content = m_Content
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
		
		Public Property Get IP()
		    IP = m_IP
        End Property
        
        Public Property Get HTTP_USER_AGENT()
            HTTP_USER_AGENT = m_HTTP_USER_AGENT
        End Property
		
		Public Property Get Count()
		    Dim intCount

		    m_strSQL = "Select COUNT(*) As [Total] From [Comments] Where [PostID] = " & m_PostID & " And [Published] = 1"
		    
            Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
		        intCount = m_objRS.Fields("Total").Value
            Else
                intCount = 0
            End If
            
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
			Count = intCount
        End Property
        
		Public Property Get Pagination()
			Pagination = m_Pagination
		End Property
		
		Public Property Get AdminLinksCount()
		    AdminLinksCount = m_AdminLinksCount
        End Property

	'private properties
        Private m_ID
        Private m_PostID
        Private m_UserID
        Private m_Name
        Private m_Email
        Private m_URL
        Private m_Content
        Private m_DateCreated
        Private m_DateModified
        Private m_Published
        Private m_IP
        Private m_HTTP_USER_AGENT
        Private m_CommentsTemplate
        Private m_EditTemplate
        
		'pagination variables
		Private m_Start
		Private m_End
        Private m_Pagination
        Private m_AdminLinksCount

		Private m_strSQL
		Private m_objRS
		Private m_objClean
		Private m_objCmd
		Private m_objSlickCMS
		
		Private m_PostTitle
		Private m_PostUrl

	'public methods
		Public Function Save()
		    Dim strReturn

            Set m_objCmd = Server.CreateObject("ADODB.Command")

		    m_objCmd.ActiveConnection = objConn
		    m_objCmd.CommandType = 4

		    If m_ID = 0 Then
                Call GetPostData()
		        
		        If Application("CommentEmails") = 1 Then
                    Call SendNotification()
                End If
		    
		        strReturn = Application("CommentSaved")
			    m_objCmd.CommandText = "InsertComment"
			    m_objCmd.Parameters.Append m_objCmd.CreateParameter("@IP", 200, 1, 15, m_IP)
			    m_objCmd.Parameters.Append m_objCmd.CreateParameter("@HTTP_USER_AGENT", 200, 1, 1024, m_HTTP_USER_AGENT)
		    Else
                strReturn = Application("CommentUpdated")
			    m_objCmd.CommandText = "Admin_UpdateComment"
			    m_objCmd.Parameters.Append m_objCmd.CreateParameter("@CommentID", 3, 1, , m_ID)
		    End If

            m_objCmd.Parameters.Append m_objCmd.CreateParameter("@PostID", 3, 1, , m_PostID)
            m_objCmd.Parameters.Append m_objCmd.CreateParameter("@UserID", 3, 1, , m_UserID)
            m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Name", 200, 1, 50, m_Name)
            m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Email", 200, 1, 255, m_Email)
            m_objCmd.Parameters.Append m_objCmd.CreateParameter("@URL", 200, 1, 1024, m_URL)
            m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Content", 200, 1, len(m_Content), m_Content)
		    m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Published", 3, 1, , m_Published)
		    m_objCmd.Execute

		    Set m_objCmd = Nothing

            Save = strReturn
		End Function
		
		Public Function GetComments()
		    Dim strTemplate, strEditTemplate
		    Dim intAlternate
		    
		    m_strSQL = "Execute SelectComments " & m_PostID	    
		    intAlternate = 0
		    
            Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
				    strTemplate = m_CommentsTemplate
				    
					If intAlternate = 1 Then
						strTemplate = Replace(strTemplate,"[class]","comment-alt")
						intAlternate = 0
					Else
						strTemplate = Replace(strTemplate,"[class]","comment")
						intAlternate = intAlternate + 1
					End If

                    m_ID = m_objRS.Fields("CommentID").Value
                    m_Name = m_objRS.Fields("Name").Value
                    m_URL = m_objRS.Fields("URL").Value
                    m_Content = m_objRS.Fields("Content").Value
                    m_DateCreated = m_objRS.Fields("DateCreated").Value

					If Session("LoggedOn") = true Then
					    'add the EditTemplate to allow an Admin to carry out Administration tasks on the Comment
                        strEditTemplate = m_EditTemplate
                        strEditTemplate = Replace(strEditTemplate,"[id]",m_ID)
					    strTemplate = Replace(strTemplate,"[admin]",strEditTemplate)
					Else
					    strTemplate = Replace(strTemplate,"[admin]","")
					End If
                    
                    If m_URL <> "n/a" Then
                        strTemplate = Replace(strTemplate,"[name]","<a href=""" & m_URL & """>" & m_Name & "</a>")
                    Else
                        strTemplate = Replace(strTemplate,"[name]",m_Name)
                    End If

                    strTemplate = Replace(strTemplate,"[date]",m_DateCreated)
                    strTemplate = Replace(strTemplate,"[content]",m_Content)
                    strTemplate = Replace(strTemplate,"[permalink]","comment-" & m_ID)

					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
			m_ID = 0
		End Function
		
		Function GetAdminComments()
            Dim strTemplate
			Dim strPublished
			Dim strPageable
			Dim strPostTitle
			
			'count the comments for pagination use
			m_strSQL = "Select Count(*) From dbo.Comments"
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

			m_strSQL = "Execute Admin_SelectComments " & m_Start & "," & m_End
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTemplate = m_CommentsTemplate
					
					m_ID = m_objRS.Fields("CommentID").Value
					m_PostID = m_objRS.Fields("PostID").Value
					strPostTitle = m_objRS.Fields("PostTitle").Value
					m_UserID = m_objRS.Fields("UserID").Value
					m_Name = m_objRS.Fields("Name").Value
					m_Email = m_objRS.Fields("Email").Value
					m_URL = m_objRS.Fields("URL").Value
					m_IP = m_objRS.Fields("IP").Value
					m_Content = Server.HTMLEncode(m_objRS.Fields("Content").Value) 'to prevent XSS
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
					
					strTemplate = Replace(strTemplate,"[commentid]",m_ID)
					strTemplate = Replace(strTemplate,"[postid]",m_PostID)
					strTemplate = Replace(strTemplate,"[posttitle]",strPostTitle)
                    strTemplate = Replace(strTemplate,"[userid]",m_UserID)
                    strTemplate = Replace(strTemplate,"[name]",m_Name)
                    strTemplate = Replace(strTemplate,"[email]",m_Email)
                    strTemplate = Replace(strTemplate,"[url]",IIf(m_URL = "n/a","",m_URL))
                    strTemplate = Replace(strTemplate,"[ip]",m_IP)
                    strTemplate = Replace(strTemplate,"[content]",m_Content)
					strTemplate = Replace(strTemplate,"[datemodified]",m_DateModified)
					strTemplate = Replace(strTemplate,"[published]",strPublished)
					
					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
			m_ID = 0
		End Function
		
		Public Function GetAdminComment()
			m_strSQL = "Execute Admin_SelectComment " & m_ID
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				m_ID = m_objRS.Fields("CommentID").Value
				m_PostID = m_objRS.Fields("PostID").Value
				m_UserID = m_objRS.Fields("UserID").Value
				m_Name = m_objRS.Fields("Name").Value
				m_Email = m_objRS.Fields("Email").Value
				m_URL = m_objRS.Fields("URL").Value
				m_IP = m_objRS.Fields("IP").Value
				m_HTTP_USER_AGENT = m_objRS.Fields("HTTP_USER_AGENT").Value
				m_Content = m_objRS.Fields("Content").Value
				m_DateCreated = m_objRS.Fields("DateCreated").Value
				m_DateModified = m_objRS.Fields("DateModified").Value
				m_Published = m_objRS.Fields("Published").Value
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Function
		
		Public Sub Delete()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			m_objCmd.CommandText = "Admin_DeleteComment"
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@CommentID", 3, 1, , m_ID)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			Response.Write(Application("CommentDeleted"))
		End Sub
		
		Public Sub RecentComments()
		    'returns a list of the 5 most recent comments
            Dim strTemplate
		    
		    m_strSQL = "Execute RecentComments"
		    
            Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
			    Set m_objClean = New Clean
			
				Do While Not m_objRS.EOF
				    strTemplate = m_CommentsTemplate

                    m_ID = m_objRS.Fields("CommentID").Value
                    m_PostTitle = m_objRS.Fields("PostTitle").Value
                    m_PostUrl = objPost.FormatURL(m_objRS.Fields("PostURL").Value,"output") & "#comment-" & m_ID
                    m_Name = m_objRS.Fields("Name").Value
                    m_Content = m_objRS.Fields("Content").Value
                    m_DateCreated = m_objRS.Fields("DateCreated").Value
                    
                    m_objClean.Data = m_Content
                    Call m_objClean.StripHTML()
                    m_Content = m_objClean.Data

                    strTemplate = Replace(strTemplate,"[posttitle]",m_PostTitle)
                    strTemplate = Replace(strTemplate,"[url]",m_PostUrl)
                    strTemplate = Replace(strTemplate,"[name]",m_Name)
                    strTemplate = Replace(strTemplate,"[date]",m_DateCreated)
                    strTemplate = Replace(strTemplate,"[content]",m_Content)
                    strTemplate = Replace(strTemplate,"[permalink]","comment-" & m_ID)

					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
				
				Set m_objClean = Nothing
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
			m_ID = 0
		End Sub

	'private methods
		Private Sub GetPostData()
	        'retrieve the Post Title and Url that the comment is for
            m_strSQL = "Select [Title],[URL] = CASE [Pageable] When 1 Then Replace(convert(varchar, [DateCreated], 111), '/', '-') + '/' + [URL] Else [URL] End From dbo.Posts Where PostID = " & m_PostID

            Set m_objRS = Server.CreateObject("ADODB.RecordSet")
	        m_objRS.Open m_strSQL,objConn,0,1

	        If Not m_objRS.BOF Then
                m_PostTitle = m_objRS.Fields("Title").Value
                m_PostUrl = Application("SiteURL") & m_objRS.Fields("URL").Value & "/"
            Else
                m_PostTitle = "[Could not locate Post Title]"
                m_PostUrl = "[Could not locate Post Url]"
            End If
            
	        If m_objRS.State <> 0 Then m_objRS.Close
	        Set m_objRS = Nothing
		End Sub
		
		Private Sub SendNotification()
            'send an email to notify of new comment
            Dim strBody

            strBody = Application("NewCommentEmailBody")
            strBody = Replace(strBody,"[post]",m_PostTitle)
            strBody = Replace(strBody,"[url]",m_PostUrl)
            strBody = Replace(strBody,"[name]",m_Name)
            strBody = Replace(strBody,"[ip]",m_IP)
            strBody = Replace(strBody,"[id]",m_ID) 'not populated at this stage - in future, return ID from Stored Procedure and then send the notification email?
            strBody = Replace(strBody,"[email]",m_Email)
            strBody = Replace(strBody,"[website]",m_URL)
            strBody = Replace(strBody,"[comment]",m_Content)

            Set m_objSlickCMS = New SlickCMS
            Call m_objSlickCMS.SendNotification("New Comment", strBody)
            Set m_objSlickCMS = Nothing
		End Sub
End Class
%>