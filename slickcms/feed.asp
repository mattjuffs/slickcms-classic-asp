<%
Class Feed	
	'public properties - let (set)
        Public Property Let Version(p_Version)
            Set m_objClean = New Clean
			m_objClean.Data = p_Version
			Call m_objClean.NumericPlus()
			Call m_objClean.MaxLength(10)
			m_Version = lcase(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
        Public Property Let FType(p_Type)
            'prefixed with F as Type is a reserved word
            Set m_objClean = New Clean
			m_objClean.Data = p_Type
			Call m_objClean.Alpha()
			Call m_objClean.MaxLength(8)
			m_Type = m_objClean.Data
			Set m_objClean = Nothing
		End Property

	'public properties - get (retrieve)
        Public Property Get XML()
			XML = m_XML
		End Property
		
	'private properties
		Private m_XML
		Private m_Version
		Private m_Type
		
        Private m_strSQL
		Private m_objRS
		Private m_objClean
		Private m_objCmd
		Private m_objPost
		Private m_URL

	'public methods
		Public Sub RSS()
		    If m_Version = "2.0" Then
		        Select Case m_Type
		            Case "posts"
		                Call RSS_Posts()
                    Case "comments"
                        Call RSS_Comments()
                    Case Else
                        Response.Write("Invalid Type")
                End Select
            ElseIf m_Version = "0.92" Then
                '#332
            Else
                Response.Write("Invalid Version")
            End If
		End Sub
		
		Public Sub Atom()
		    '#333
		End Sub
		
		Public Sub Sitemap()
		    Set m_objPost = New Post
		    
		    'header
		    m_XML = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
            m_XML = m_XML & "<urlset xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">" & vbCrLf

            Response.Write(m_XML)
            m_XML = ""
		    
		    'items
            Set m_objRS = Server.CreateObject("ADODB.RecordSet")
            m_strSQL = "Execute Feed_Sitemap"
            m_objRS.Open m_strSQL, Application("ConnectionString"), 0, 1
            
            If Not m_objRS.BOF Then
                Do While Not m_objRS.EOF
                    m_URL = m_objPost.FormatURL(m_objRS.Fields("loc").Value, "output")

                    m_XML = m_XML & "<url>" & vbCrLf
                    m_XML = m_XML & "<loc>" & m_URL & "</loc>" & vbCrLf
                    m_XML = m_XML & "<lastmod>" & m_objRS.Fields("lastmod").Value & "</lastmod>" & vbCrLf
                    m_XML = m_XML & "<changefreq>" & m_objRS.Fields("changefreq").Value & "</changefreq>" & vbCrLf
                    m_XML = m_XML & "<priority>" & IIf((m_objRS.Fields("priority").Value = "1"),"1.0","0.5") & "</priority>" & vbCrLf
                    m_XML = m_XML & "</url>" & vbCrLf
                    
                    Response.Write(m_XML)
                    m_XML = ""
                    
                    m_objRS.MoveNext
                Loop
            Else
                'no posts
            End If
            
            If m_objRS.State <> 0 Then m_objRS.Close
            Set m_objRS = Nothing
		    
		    'footer
		    m_XML = "</urlset>" & vbCrLf
            Response.Write(m_XML)
            m_XML = ""

		    Set m_objPost = Nothing
		End Sub

	'private methods
        Private Sub RSS_Posts()
            Set m_objPost = New Post

            'header
            m_XML = "<?xml version=""1.0""?>" & vbCrLf
            m_XML = m_XML & "<rss version=""2.0"" xmlns:atom=""http://www.w3.org/2005/Atom"">" & vbCrLf
            m_XML = m_XML & "<channel>" & vbCrLf
            m_XML = m_XML & "<title>" & Application("SiteName") & " Posts</title>" & vbCrLf
            m_XML = m_XML & "<link>" & Application("SiteURL") & "</link>" & vbCrLf
            m_XML = m_XML & "<description>" & Application("SiteName") & " Posts RSS 2.0 Feed</description>" & vbCrLf
            m_XML = m_XML & "<language>en</language>" & vbCrLf
            m_XML = m_XML & "<pubDate>" & RSS_Date(now(), "GMT") & "</pubDate>" & vbCrLf
            m_XML = m_XML & "<generator>Weblog Editor 2.0</generator>" & vbCrLf
            m_XML = m_XML & "<ttl>60</ttl>" & vbCrLf
            m_XML = m_XML & "<atom:link href=""" & Application("SiteURL") & "rss2.asp?t=posts"" rel=""self"" type=""application/rss+xml"" />" & vbCrLf
            
            Response.Write(m_XML)
            m_XML = ""

            'items      
            Set m_objRS = Server.CreateObject("ADODB.RecordSet")
            m_strSQL = "Execute Feed_RSS20"
            m_objRS.Open m_strSQL, Application("ConnectionString"), 0, 1
            
            If Not m_objRS.BOF Then
                Do While Not m_objRS.EOF
                    m_URL = m_objPost.FormatURL(m_objRS.Fields("URL").Value, "output")

                    m_XML = m_XML & "<item>" & vbCrLf
                    m_XML = m_XML & "<title>" & m_objRS.Fields("Title").Value & "</title>" & vbCrLf
                    m_XML = m_XML & "<link>" & m_URL & "</link>" & vbCrLf
                    m_XML = m_XML & "<guid>" & m_URL & "</guid>" & vbCrLf
                    m_XML = m_XML & "<description><![CDATA[" & CleanDescription(m_objRS.Fields("Content").Value) & "]]></description>" & vbCrLf
                    m_XML = m_XML & "<comments>" & m_URL & "#comments" & "</comments>" & vbCrLf
                    m_XML = m_XML & "<pubDate>" & m_objRS.Fields("DateCreated").Value & "</pubDate>" & vbCrLf
                    m_XML = m_XML & "</item>" & vbCrLf
                    
                    Response.Write(m_XML)
                    m_XML = ""
                    
                    m_objRS.MoveNext
                Loop
            Else
                'no posts
            End If
            
            If m_objRS.State <> 0 Then m_objRS.Close
            Set m_objRS = Nothing

            'footer
            m_XML = "</channel>" & vbCrLf
            m_XML = m_XML & "</rss>" & vbCrLf
            Response.Write(m_XML)
            m_XML = ""

            Set m_objPost = Nothing
		End Sub

        Private Sub RSS_Comments()
            Set m_objPost = New Post

            'header
            m_XML = "<?xml version=""1.0""?>" & vbCrLf
            m_XML = m_XML & "<rss version=""2.0"" xmlns:atom=""http://www.w3.org/2005/Atom"">" & vbCrLf
            m_XML = m_XML & "<channel>" & vbCrLf
            m_XML = m_XML & "<title>" & Application("SiteName") & " Comments</title>" & vbCrLf
            m_XML = m_XML & "<link>" & Application("SiteURL") & "</link>" & vbCrLf
            m_XML = m_XML & "<description>" & Application("SiteName") & " Comments RSS 2.0 Feed</description>" & vbCrLf
            m_XML = m_XML & "<language>en</language>" & vbCrLf
            m_XML = m_XML & "<pubDate>" & RSS_Date(now(), "GMT") & "</pubDate>" & vbCrLf
            m_XML = m_XML & "<generator>SlickCMS " & Application("SlickCMS_Version") & "</generator>" & vbCrLf
            m_XML = m_XML & "<ttl>60</ttl>" & vbCrLf
            m_XML = m_XML & "<atom:link href=""" & Application("SiteURL") & "rss2.asp?t=comments"" rel=""self"" type=""application/rss+xml"" />" & vbCrLf
            
            Response.Write(m_XML)
            m_XML = ""

            'items      
            Set m_objRS = Server.CreateObject("ADODB.RecordSet")
            m_strSQL = "Execute Feed_RSS20_Comments"
            m_objRS.Open m_strSQL, Application("ConnectionString"), 0, 1
            
            If Not m_objRS.BOF Then
                Do While Not m_objRS.EOF
                    m_URL = Application("SiteUrl") & m_objRS.Fields("URL").Value

                    m_XML = m_XML & "<item>" & vbCrLf
                    m_XML = m_XML & "<title>" & m_objRS.Fields("Title").Value & "</title>" & vbCrLf
                    m_XML = m_XML & "<link>" & m_URL & "</link>" & vbCrLf
                    m_XML = m_XML & "<guid>" & m_URL & "</guid>" & vbCrLf
                    m_XML = m_XML & "<description><![CDATA[" & CleanDescription(m_objRS.Fields("Content").Value) & "]]></description>" & vbCrLf
                    m_XML = m_XML & "<pubDate>" & m_objRS.Fields("DateCreated").Value & "</pubDate>" & vbCrLf
                    m_XML = m_XML & "</item>" & vbCrLf
                    
                    Response.Write(m_XML)
                    m_XML = ""
                    
                    m_objRS.MoveNext
                Loop
            Else
                'no comments
            End If
            
            If m_objRS.State <> 0 Then m_objRS.Close
            Set m_objRS = Nothing

            'footer
            m_XML = "</channel>" & vbCrLf
            m_XML = m_XML & "</rss>" & vbCrLf
                
            Response.Write(m_XML)
            m_XML = ""

            Set m_objPost = Nothing
		End Sub

		Private Function RSS_Date(dDate, offset)
            Dim dDay, dDays, dMonth, dYear
            Dim dHours, dMinutes, dSeconds

            dDate = CDate(dDate)
            dDay = WeekdayName(Weekday(dDate),true)
            dDays = Day(dDate)
            dMonth = MonthName(Month(dDate), true)
            dYear = Year(dDate)
            dHours = zeroPad(Hour(dDate), 2)
            dMinutes = zeroPad(Minute(dDate), 2)
            dSeconds = zeroPad(Second(dDate), 2)

            RSS_Date = dDay & ", " & dDays & " " & dMonth & " " & dYear & " "& dHours & ":" & dMinutes & ":" & dSeconds & " " & offset
        End Function 

        Private Function zeroPad(m, t)
           zeroPad = String((t - Len(m)),"0") & m
        End Function

        Private Function CleanDescription(str)
            'removes whitespace from the <description>
            str = Replace(str, vbCrLf, "")
            CleanDescription = str
        End Function
End Class
%>