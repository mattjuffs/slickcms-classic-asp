<%
Class Post
	'public properties - let (set)
		Public Property Let Url(p_Url)
			m_Url = FormatUrl(p_Url,"input")
			m_RawUrl = m_Url
		End Property
		
		Public Property Let UrlType(p_UrlType)
		    m_UrlType = p_UrlType
		End Property

		Public Property Let ID(p_ID)
			Set m_objClean = New Clean
			m_objClean.Data = p_ID
			Call m_objClean.Numeric()
			m_ID = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
	
		Public Property Let UserID(p_UserID)
			Set m_objClean = New Clean
			m_objClean.Data = p_UserID
			Call m_objClean.Numeric()
			m_UserID = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Author(p_Author)
			Set m_objClean = New Clean
			m_objClean.Data = p_Author
			Call m_objClean.AlphaNumeric()
			Call m_objClean.MaxLength(50)
			m_Author = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Title(p_Title)
			Set m_objClean = New Clean
			m_objClean.Data = p_Title
			Call m_objClean.AlphaNumeric()
			Call m_objClean.Maxlength(255)
			m_Title = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Summary(p_Summary)
			m_Summary = p_Summary
		End Property
		
		Public Property Let Content(p_Content)
			m_Content = p_Content
		End Property
		
		Public Property Let Search(p_Search)
			Set m_objClean = New Clean
			m_objClean.Data = p_Search
			Call m_objClean.AlphaNumeric()
			Call m_objClean.Maxlength(255)
			m_Search = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Published(p_Published)
			Set m_objClean = New Clean
			m_objClean.Data = p_Published
			Call m_objClean.Numeric()
			m_Published = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property

		Public Property Let PostsTemplate(p_PostsTemplate)
			m_PostsTemplate = p_PostsTemplate
		End Property
		
		Public Property Let NavigationTemplate(p_NavigationTemplate)
			m_NavigationTemplate = p_NavigationTemplate
		End Property
		
		Public Property Let SearchTemplate(p_SearchTemplate)
			m_SearchTemplate = p_SearchTemplate
		End Property
		
		Public Property Let Keywords(p_Keywords)
			Set m_objClean = New Clean
			m_objClean.Data = p_Keywords
			Call m_objClean.SQL()
			Call m_objClean.MaxLength(100)
			m_Keywords = m_objClean.Data
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
		
		Public Property Let Pageable(p_Pageable)
			Set m_objClean = New Clean
			m_objClean.Data = p_Pageable
			Call m_objClean.Numeric()
			m_Pageable = CInt(m_objClean.Data)
			Set m_objClean = Nothing
			
			If m_Pageable = "" Then m_Pageable = 0
		End Property
		
        Public Property Let Year(p_Year)
			Set m_objClean = New Clean
			m_objClean.Data = p_Year
			Call m_objClean.Numeric()
			m_Year = CInt(m_objClean.Data)
			Set m_objClean = Nothing
			If m_Year = "" Then m_Year = 0
		End Property

        Public Property Let Month(p_Month)
			Set m_objClean = New Clean
			m_objClean.Data = p_Month
			Call m_objClean.Numeric()
			m_Month = CInt(m_objClean.Data)
			Set m_objClean = Nothing
			If m_Month = "" Then m_Month = 0
		End Property

        Public Property Let Day(p_Day)
			Set m_objClean = New Clean
			m_objClean.Data = p_Day
			Call m_objClean.Numeric()
			m_Day = CInt(m_objClean.Data)
			Set m_objClean = Nothing
			If m_Day = "" Then m_Day = 0
		End Property
		
        Public Property Let ArchivesTemplate(p_ArchivesTemplate)
			m_ArchivesTemplate = p_ArchivesTemplate
		End Property
		
	'public properties - get (retrieve)
		Public Property Get Url()
		    Select Case m_UrlType
		        Case "date"
		            'prefix the post with the date
		            If m_DateCreated <> "" Then
		                Url = ConvertDate(m_DateCreated) & "/" & m_Url
		            Else
		                Url = m_Url
		            End If
		        Case Else
			        Url = m_Url
            End Select
		End Property
		
		Public Property Get UrlType()
		    UrlType = m_UrlType
		End Property
			
		Public Property Get ID()
			ID = m_ID
		End Property
		
		Public Property Get UserID()
			UserID = m_UserID
		End Property
		
		Public Property Get Author()
			Author = m_Author
		End Property
		
		Public Property Get Title()
			Title = m_Title
		End Property
		
		Public Property Get Summary()
			Summary = m_Summary
		End Property
		
		Public Property Get Content()
			Content = m_Content
		End Property
		
		Public Property Get Search()
			Search = m_Search
		End Property
		
		Public Property Get DateCreated()
			DateCreated = HumanDate(m_DateCreated)
		End Property
		
		Public Property Get DateModified()
			DateModified = HumanDate(m_DateModified)
		End Property
		
		Public Property Get Published()
			Published = m_Published
		End Property
		
		Public Property Get Keywords()
			Keywords = m_Keywords
		End Property
		
		Public Property Get Pagination()
			Pagination = m_Pagination
		End Property
		
		Public Property Get Pageable()
			Pageable = m_Pageable
		End Property
		
		Public Property Get AdminPostsCount()
		    AdminPostsCount = m_AdminPostsCount
        End Property

	'private properties
		Private m_Url
		Private m_RawUrl 'internal class use only
		Private m_UrlType
		Private m_ID
		Private m_UserID
		Private m_Author
		Private m_Title
		Private m_Summary
		Private m_Content
		Private m_Search
		Private m_DateCreated
		Private m_DateModified
		Private m_Published
		Private m_PostsTemplate
		Private m_NavigationTemplate
		Private m_SearchTemplate
		Private m_Keywords
		Private m_Pagination
		Private m_Pageable
		Private m_Year
		Private m_Month
		Private m_Day
		Private m_ArchivesTemplate
		Private m_CategoryName
		Private m_Tag
		Private m_AdminPostsCount

		Private m_strSQL
		Private m_objRS
		Private m_objClean
		Private m_objCmd
		
		'pagination variables
		Private m_Start
		Private m_End

	'public methods
		Public Sub Navigation(CategoryID)
		    'retrieves a list of Posts used for navigation
			Dim strTemplate
			Dim strUrl
			Dim strPage

			m_strSQL = "Execute Navigation " & CategoryID
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strUrl = m_objRS.Fields("Url").Value
					strPage = m_objRS.Fields("Title").Value
					m_DateCreated = m_objRS.Fields("DateCreated").Value
					
					If strUrl <> "404" Then
						strTemplate = m_NavigationTemplate
						strTemplate = Replace(strTemplate,"[url]",FormatUrl(strUrl,"output"))
						strTemplate = Replace(strTemplate,"[title]",strPage)
						If m_Title = strPage Then
							strTemplate = Replace(strTemplate,"<li>","<li class=""selected"">")
						End If

						Response.Write(strTemplate)
					End If
					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub

		Public Function GetPost()
		    'retrieves a Post for use on a single page
			m_strSQL = "Execute SelectPost '" & m_Url & "'"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				m_ID = m_objRS.Fields("PostID").Value
				m_Author = m_objRS.Fields("Author").Value
				m_Title = m_objRS.Fields("Title").Value
				m_Summary = m_objRS.Fields("Summary").Value
				m_Content = m_objRS.Fields("Content").Value
				m_DateCreated = m_objRS.Fields("DateCreated").Value
				m_DateModified = m_objRS.Fields("DateModified").Value
			Else
				m_ID = 0
				m_Author = ""
				m_Title = ""
				m_Summary = ""
				m_Content = ""
				m_DateCreated = ""
				m_DateModified = ""
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
			If m_Url = "search" Then
				m_Content = SearchPosts()
			End If
		End Function

        Public Sub GetPosts(strType)
            'retrieves a list of Posts for use on the site (paginated)
			Dim strTemplate
			Dim intComments
			
			intPosts = 0

			'pagination
			If m_Pagination = 0 Then m_Pagination = 1
			m_End = (Application("Pagination") * m_Pagination)
			m_Start = (m_End - Application("Pagination"))+1
			If m_Start = 0 Then m_Start = 1
			
			Select Case strType
			    Case "posts"
			        m_strSQL = "Execute SelectPosts " & m_Start & "," & m_End
                Case "archives"
                    m_strSQL = "Execute SelectPostsArchive " & m_Year & "," & m_Month & "," & m_Day & "," & m_Start & "," & m_End
                Case "categories"
                    m_CategoryName = Replace(m_RawUrl,"-"," ")
                    m_strSQL = "Execute SelectPostsCategory '" & m_CategoryName & "'," & m_Start & "," & m_End
			    Case "tags"
			        m_Tag = Replace(m_RawUrl,"-"," ")
			        m_strSQL = "Execute SelectPostsTag '" & m_Tag & "'," & m_Start & "," & m_End
            End Select
            
            If m_strSQL = "" Then
                'trigger an error to the webmaster, as this would be a result of their code - not user input
                Response.Write("strType was not provided or is not valid!")
                Response.End
            End If
			
			If Application("Debug") = 1 Then
			    Response.Write("<!--GetPosts() SQL: " & m_strSQL & "-->")
            End If

			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					intPosts = (intPosts + 1)
					strTemplate = m_PostsTemplate

					m_ID = m_objRS.Fields("PostID").Value
					m_Author = m_objRS.Fields("Author").Value
					m_Title = m_objRS.Fields("Title").Value
					m_Summary = m_objRS.Fields("Summary").Value
					m_Content = m_objRS.Fields("Content").Value
					m_DateCreated = HumanDate(m_objRS.Fields("DateCreated").Value)
					m_DateModified = HumanDate(m_objRS.Fields("DateModified").Value)
					m_Url = FormatUrl(m_objRS.Fields("Url").Value,"output")
					
					strTemplate = Replace(strTemplate,"[postid]",m_ID)
					strTemplate = Replace(strTemplate,"[author]",m_Author)
					strTemplate = Replace(strTemplate,"[title]",m_Title)
					strTemplate = Replace(strTemplate,"[url]",m_Url)
					strTemplate = Replace(strTemplate,"[summary]",m_Summary)
					strTemplate = Replace(strTemplate,"[content]",m_Content)
					strTemplate = Replace(strTemplate,"[datecreated]",m_DateCreated)
					strTemplate = Replace(strTemplate,"[datemodified]",m_DateModified)
					
					objComment.PostID = m_ID
					intComments = objComment.Count()
					strTemplate = Replace(strTemplate,"[comments]",intComments)

					strCategories = objCategory.GetPostCategories(m_ID)
                    If Right(strCategories,2) = ", " Then strCategories = Left(strCategories,(Len(strCategories)-2))
                    strTemplate = Replace(strTemplate,"[categories]",strCategories)
                    
					strTags = objTag.GetPostTags(m_ID)
                    If Right(strTags,2) = ", " Then strTags = Left(strTags,(Len(strTags)-2))
                    strTemplate = Replace(strTemplate,"[tags]",strTags)
					
					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
			End If

			If Application("Debug") = 1 Then
			    Response.Write("<!--Page:" & m_Pagination & "|Start:" & m_Start & "|End:" & m_End & "|intPosts:" & intPosts & "-->" & vbCrLf)
            End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Function Paging(strPagination)
			'this should be updated to move the HTML out of the class and into a template
			Dim intOlder, intNewer, intPage
			Dim strOlder, strNewer
			Dim strPrefix
			Dim strYear, strMonth, strDay 'string versions of date parts
			
			'calculate pages
			intPage = m_Pagination
			If intPage = 0 Then intPage = 1
			intOlder = (intPage + 1)
			intNewer = (intPage - 1)
			
			'build up the prefix
			strPrefix = "/"

			Select Case m_UrlType
                Case "archive"
                    'ensure date parts are full length
                    strYear = m_Year
                    strMonth = IIf((Len(m_Month)=1), "0" & m_Month, m_Month)
                    strDay = IIf((Len(m_Day)=1), "0" & m_Day, m_Day)

			        If m_Year <> 0 Then
			            strPrefix = (strPrefix & strYear & "/") 'e.g. /2010/
			        End If
        			
			        If m_Month <> 0 Then
			            strPrefix = (strPrefix & strMonth & "/") 'e.g. /12/
			        End If
        			
			        If m_Day <> 0 Then
			            strPrefix = (strPrefix & strDay & "/") 'e.g. /31/
			        End If
                Case "tag"
                    strPrefix = strPrefix & "tag/" & m_RawUrl & "/"
                Case "category"
			        strPrefix = strPrefix & "category/" & m_RawUrl & "/"
			End Select

			strOlder = "<a href=""" & strPrefix & "page/" & intOlder & "/"">Older Posts</a>"
			
			'cms/blog/blog with homepage
			If Application("Homepage") = 1 And intNewer = 1 Then
			    strNewer = "<a href=""" & strPrefix & "page/" & intNewer & "/"">Newer Posts</a>"
			ElseIf Application("Homepage") = 0 And intNewer = 1 Then
				strNewer = "<a href=""" & strPrefix & """>Newer Posts</a>"
			Else
				strNewer = "<a href=""" & strPrefix & "page/" & intNewer & "/"">Newer Posts</a>"
			End If
			
			'on homepage
			If intNewer = 0 Then strNewer = ""			
			
            'on the final page
            If TotalPosts() <= (intPage * Application("Pagination")) Then strOlder = ""

			strPagination = Replace(strPagination,"[older]",strOlder)
			strPagination = Replace(strPagination,"[newer]",strNewer)
			
			'hide the pagination altogether if not required
			If strOlder = "" And strNewer = "" Then strPagination = ""
			
			Paging = strPagination
		End Function
		
		Public Sub GetAdminPosts()
			Dim strTemplate
			Dim strPublished
			Dim strPageable
			
			intPosts = 0
			
			'count the posts for pagination use
			m_strSQL = "Select Count(*) From dbo.Posts"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			If Not m_objRS.BOF Then
			    intPosts = m_objRS.Fields(0).Value
            End If
            If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
			m_AdminPostsCount = intPosts
			
            'pagination
			If m_Pagination = 0 Then m_Pagination = 1
			m_End = (Application("AdminPagination") * m_Pagination)
			m_Start = (m_End - Application("AdminPagination"))+1
			If m_Start = 0 Then m_Start = 1
			
			m_strSQL = "Execute Admin_SelectPosts " & m_Start & "," & m_End
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTemplate = m_PostsTemplate
					
					m_ID = m_objRS.Fields("PostID").Value
					m_Author = NullCheck(m_objRS.Fields("Author").Value)
					m_UserID = m_objRS.Fields("UserID").Value
					m_Title = m_objRS.Fields("Title").Value
					m_DateCreated = m_objRS.Fields("DateCreated").Value
					m_DateModified = m_objRS.Fields("DateModified").Value
					m_Published = m_objRS.Fields("Published").Value
					m_Pageable = m_objRS.Fields("Pageable").Value
					
					Select Case m_Published
						Case 1
							strPublished = "Yes"
						Case Else
							strPublished = "No"
					End Select
					
					Select Case m_Pageable
						Case 1
							strPageable = "Yes"
						Case Else
							strPageable = "No"
					End Select
					
					strTemplate = Replace(strTemplate,"[postid]",m_ID)
					strTemplate = Replace(strTemplate,"[author]",m_Author)
					strTemplate = Replace(strTemplate,"[title]",m_Title)
					strTemplate = Replace(strTemplate,"[datecreated]",m_DateCreated)
					strTemplate = Replace(strTemplate,"[datemodified]",m_DateModified)
					strTemplate = Replace(strTemplate,"[published]",strPublished)
					strTemplate = Replace(strTemplate,"[pageable]",strPageable)
					
					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Sub GetAdminPost()
			If m_ID <> 0 Then
				m_strSQL = "Execute Admin_SelectPost " & m_ID
				Set m_objRS = Server.CreateObject("ADODB.RecordSet")
				m_objRS.Open m_strSQL,objConn,0,1
				
				If Not m_objRS.BOF Then
					m_ID = m_objRS.Fields("PostID").Value
					m_UserID = m_objRS.Fields("UserID").Value
					m_Author = m_objRS.Fields("Author").Value
					m_Title = m_objRS.Fields("Title").Value
					m_Summary = m_objRS.Fields("Summary").Value
					m_Content = m_objRS.Fields("Content").Value
					m_Search = m_objRS.Fields("Search").Value
					m_DateCreated = m_objRS.Fields("DateCreated").Value
					m_DateModified = m_objRS.Fields("DateModified").Value
					m_Published = m_objRS.Fields("Published").Value
					m_Pageable = m_objRS.Fields("Pageable").Value
					m_Url = m_objRS.Fields("Url").Value
				Else
					m_ID = 0
				End If
				
				If m_objRS.State <> 0 Then m_objRS.Close
				Set m_objRS = Nothing
			End If
			
			If m_ID = 0 Then
				m_UserID = 0
				m_Author = ""
				m_Title = ""
				m_Summary = ""
				m_Content = ""
				m_Search = ""
				m_DateCreated = ""
				m_DateModified = ""
				m_Published = ""
				m_Pageable = ""
				m_Url = ""
			End If
		End Sub
		
		Public Sub Save()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			If m_ID = 0 Then
			    'new Post, insert a new record
				m_objCmd.CommandText = "Admin_InsertPost"
			Else
			    'existing Post, update the record
				m_objCmd.CommandText = "Admin_UpdatePost"
				m_objCmd.Parameters.Append m_objCmd.CreateParameter("@PostID", 3, 1, , m_ID)
			End If

			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@UserID", 3, 1, , m_UserID)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Title", 200, 1, 255, m_Title)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Summary", 200, 1, len(m_Summary), m_Summary)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Content", 200, 1, len(m_Content), m_Content)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Search", 200, 1, 255, m_Search)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Published", 3, 1, , m_Published)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Pageable", 3, 1, , m_Pageable)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Url", 200, 1, 255, m_Url)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			If m_ID = 0 Then
                '#399 get ID
                GetPostID()
				Response.Write(Application("PostSaved"))
			Else
				Response.Write(Application("PostUpdated"))
			End If
		End Sub
		
		Public Sub Delete()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			m_objCmd.CommandText = "Admin_DeletePost"
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@PostID", 3, 1, , m_ID)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			Response.Write(Application("PostDeleted"))
		End Sub
		
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
   				        strUrl = Application("SiteUrl") & strUrl & "/"
					End If
				Case Else
					strUrl = ""
			End Select

			FormatUrl = strUrl
		End Function
		
		Public Function SearchPosts()
		    'used for Search functionality, builds up a string to use as the Post content
			Dim strReturn, strResults, strTitle, strSummary, strUrl

			strReturn = Replace(m_Content,"[keywords]",m_Keywords)
			strReturn = strReturn & vbCrLf
			
			m_strSQL = "Execute SearchPosts '" & m_Keywords & "'"
			
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTitle = m_objRS.Fields("Title").Value
					strSummary = m_objRS.Fields("Summary").Value
					strUrl = FormatUrl(m_objRS.Fields("Url").Value,"output")
					
					'ensure a Summary is displayed (use the Title if blank)
					If strSummary = "" Then strSummary = strTitle

					strResults = strResults & m_SearchTemplate
					strResults = Replace(strResults,"[title]",strTitle)
					strResults = Replace(strResults,"[summary]",strSummary)
					strResults = Replace(strResults,"[url]",strUrl)

					m_objRS.MoveNext
				Loop
			Else
				strResults = Application("NoSearchResults")
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing

			strReturn = strReturn & strResults

			SearchPosts = strReturn
		End Function
		
		Public Function TotalPosts()
		    'counts posts as paging is done in sql server itself, not in application
			Dim intReturn
            
            If m_UrlType = "archive" Then
                m_CategoryName = ""
                m_Tag = ""
            ElseIf m_UrlType = "tag" Then
                m_CategoryName = ""
                m_Tag = Replace(m_RawUrl,"-"," ")
            ElseIf m_UrlType = "category" Then
                m_CategoryName = Replace(m_RawUrl,"-"," ")
                m_Tag = ""
            End If

			m_strSQL = "Execute SelectTotalPosts '" & m_Tag & "','" & m_CategoryName & "'," & m_Year & "," & m_Month & "," & m_Day
			
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				intReturn = m_objRS.Fields("TotalPosts").Value
			Else
				intReturn = 0
			End If

			TotalPosts = intReturn
		End Function
		
		Public Sub Archives(strType)
		    'retrieves a list of Yearly, Monthly or Daily Archives for navigation
		    Dim strTemplate

		    strType = lcase(strType)

		    Select Case strType
		        Case "yearly","monthly","daily"
                    m_strSQL = "Execute SelectArchives '" & strType & "'"
			        Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			        m_objRS.Open m_strSQL,objConn,0,1
        			
			        If Not m_objRS.BOF Then
				        Do While Not m_objRS.EOF
				            strTemplate = m_ArchivesTemplate
    				        
				            strTemplate = Replace(strTemplate,"[url]", "/" & m_objRS.Fields("Archive").Value & "/")
				            strTemplate = Replace(strTemplate,"[archive]", Replace(m_objRS.Fields("Archive").Value,"/","-"))
				            strTemplate = Replace(strTemplate,"[postcount]",m_objRS.Fields("PostCount").Value)
    				        
				            Response.Write(strTemplate)
    				        
				            m_objRS.MoveNext
                        Loop
			        End If
    			    
                    If m_objRS.State <> 0 Then m_objRS.Close
                    Set m_objRS = Nothing
		        Case Else
		            'invalid type, do nothing
		    End Select
		End Sub

	'private methods
		Private Function NullCheck(strData)
			Dim strReturn

			If Len(strData) < 1 Then
				strReturn = ""
			Else
				strReturn = strData
			End If
			
			NullCheck = strReturn
		End Function
		
		Private Function ConvertDate(dDate)
		    'converts date to YYYY-MM-DD format
		    Dim dReturn
            Dim dYear
            Dim dMonth
            Dim dDay

            dYear = DatePart("yyyy", dDate)
            dMonth = DatePart("m", dDate)
            dDay = DatePart("d", dDate)

            'ensure month/day is in MM/DD format
            If dMonth < 10 Then dMonth = "0" & dMonth
            If dDay < 10 Then dDay = "0" & dDay

            dReturn = dYear & "-" & dMonth & "-" & dDay
		    
		    ConvertDate = dReturn
		End Function

        Private Sub GetPostID()
            m_strSQL = "Execute Admin_SelectPostID"
			
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				m_ID = m_objRS.Fields("PostID").Value
			Else
				m_ID = 0
			End If
        End Sub
End Class
%>