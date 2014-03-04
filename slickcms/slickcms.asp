<!--
    METADATA
    TYPE="typelib"
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"
    NAME="CDO for Windows 2000 Library"
-->
<%
	'SlickCMS
	'(C) Copyright MMIX Matthew Juffs (Slickhouse.com) - released under the Microsoft Reciprocal License (Ms-RL)
	
	'SlickCMS is a Classic ASP (VBScript) Content Management System framework that contains a collection of classes
	'Presentation			XHTML/CSS/ASP page(s)
	'Business				SlickCMS (objects/classes)
	'Data					MS SQL Server Database
%>
<!--#include virtual="/slickcms/category.asp"-->
<!--#include virtual="/slickcms/comment.asp"-->
<!--#include virtual="/slickcms/feed.asp"-->
<!--#include virtual="/slickcms/link.asp"-->
<!--#include virtual="/slickcms/post.asp"-->
<!--#include virtual="/slickcms/statistic.asp"-->
<!--#include virtual="/slickcms/user.asp"-->
<!--#include virtual="/slickcms/clean.asp"-->
<!--#include virtual="/slickcms/image.asp"-->
<!--#include virtual="/slickcms/relationship.asp"-->
<!--#include virtual="/slickcms/md5.asp"-->
<!--#include virtual="/slickcms/captcha.asp"-->
<!--#include virtual="/slickcms/url.asp"-->
<!--#include virtual="/slickcms/tag.asp"-->
<!--#include virtual="/slickcms/global.asp"-->
<%
Class SlickCMS
	'public properties - let (set)
		'none

	'public properties - get (retrieve)
		Public Property Get Months()
			Months = m_Months
		End Property
		
		Public Property Get Weeks()
			Weeks = m_Weeks
		End Property
		
		Public Property Get Days()
			Days = m_Days
		End Property
		
		Public Property Get Hours()
			Hours = m_Hours
		End Property
		
		Public Property Get Minutes()
			Minutes = m_Minutes
		End Property
		
		Public Property Get Seconds()
			Seconds = m_Seconds
		End Property
	
	'private properties
		Private m_strSQL
		Private m_objRS
		Private m_objClean

		Private m_Months
		Private m_Weeks
		Private m_Days
		Private m_Hours
		Private m_Minutes
		Private m_Seconds

	'public methods
		Public Sub CalculateDate()
			'specific to OurWedding2009.co.uk - but could be adapted for other uses
			Dim dMonths, dWeeks, dDays, dHours, dMinutes, dSeconds
			m_strSQL = "Execute CalculateDate"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				m_Months = m_objRS.Fields("Months").Value
				m_Weeks = m_objRS.Fields("Weeks").Value
				m_Days = m_objRS.Fields("Days").Value
				m_Hours = m_objRS.Fields("Hours").Value
				m_Minutes = m_objRS.Fields("Minutes").Value
				m_Seconds = m_objRS.Fields("Seconds").Value
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Sub SendNotification(strSubject, strBody)
		    'used for sending email notifications to all users
		    Dim strEmail

            m_strSQL = "Select [Email] From dbo.Users Where [Active] = 1"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			'loop through email(s) and send an email notification to each
			If Not m_objRS.BOF Then
			    Do While Not m_objRS.EOF
			        strEmail = m_objRS.Fields("Email").Value
			        
			        Call SendEmail(Application("Email"), strEmail, strSubject, strBody)
			        
			        m_objRS.MoveNext
			    Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub
		
		Public Sub SendMessage(strEmail, strMessage, strName)
		    'used for the Contact form
			Dim strFrom
			Dim strBody
			Dim strTo
			Dim strSubject
			
			Set m_objClean = New Clean
		
			'get user input from form (via parameters)
			m_objClean.Data = strEmail
			m_objClean.Email()
			m_objClean.MaxLength(255)
			strFrom = m_objClean.Data

			m_objClean.Data = strMessage
			m_objClean.Encode()
			strMessage = m_objClean.Data

			m_objClean.Data = strName
			m_objClean.AlphaNumeric()
			m_objClean.MaxLength(50)
			strName = m_objClean.Data
		
			strTo = Application("Email")
			strSubject = Application("Subject")
			
			If Application("Debug") = 1 Then
			    Response.Write(strFrom & vbCrLf & strTo & vbCrLf & strName & vbCrLf & strMessage & vbCrLf & strSubject & vbCrLf & strHTMLBody)
            End If
			
			'build up the HTML email body
			strBody = Application("EmailBody")
			strBody = Replace(strBody,"[name]",strName)
			strBody = Replace(strBody,"[from]",strFrom)
			strBody = Replace(strBody,"[message]",strMessage)

			Call SendEmail(strFrom, strTo, strSubject, strBody)
			
			Set m_objClean = Nothing
		End Sub
		
		Public Function AddThis(strURL,strTitle)
			'string concatenation is acceptable here, due to the small size of string
			Dim strAddThis
			
			strURL = objPost.FormatURL(strURL,"output")			
			strTitle = strTitle & " | " & Application("SiteName")

			If Application("AddThis") = 1 Then
				strAddThis = vbCrLf
				strAddThis = strAddThis & "<!-- AddThis Bookmark Button BEGIN -->" & vbCrLf

				Select Case Application("AddThis_Version")
					Case 12
						strAddThis = strAddThis & "<script type=""text/javascript"">" & vbCrLf
						strAddThis = strAddThis & vbTab & "addthis_url = '" & strURL & "/';" & vbCrLf
						strAddThis = strAddThis & vbTab & "addthis_title = '" & strTitle & "';" & vbCrLf
						strAddThis = strAddThis & vbTab & "addthis_pub = '" & Application("AddThis_Account") & "';" & vbCrLf
						strAddThis = strAddThis & "</script>" & vbCrLf
						strAddThis = strAddThis & "<script type=""text/javascript"" src=""http://s7.addthis.com/js/addthis_widget.php?v=12"" ></script>" & vbCrLf
					Case 20
						strAddThis = strAddThis & "<script type=""text/javascript"">" & vbCrLf
						strAddThis = strAddThis & vbTab & "var addthis_pub = '" & Application("AddThis_Account") & "';" & vbCrLf
						strAddThis = strAddThis & "</script>" & vbCrLf
						strAddThis = strAddThis & "<a href=""http://www.addthis.com/bookmark.php?v=20"" onmouseover=""return addthis_open(this, '', '" & strURL & "', '" & strTitle & "')"" onmouseout=""addthis_close()"" onclick=""return addthis_sendto()"">" & vbCrLf
						
						'AddThis hosted button
						'strAddThis = strAddThis & vbTab & "<img src=""http://s7.addthis.com/static/btn/lg-share-en.gif"" width=""125"" height=""16"" alt=""Bookmark and Share"" style=""border:0;""/>" & vbCrLf
						
						'local hosted button
						strAddThis = strAddThis & vbTab & "<img src=""/images/lg-share-en.gif"" width=""125"" height=""16"" alt=""Bookmark and Share"" style=""border:0;""/>" & vbCrLf

						strAddThis = strAddThis & "</a>" & vbCrLf
						strAddThis = strAddThis & "<script type=""text/javascript"" src=""http://s7.addthis.com/js/200/addthis_widget.js""></script>" & vbCrLf
					Case 25
                        strAddThis = strAddThis & "<div class=""addthis_toolbox addthis_default_style"">" & vbCrLf
                        strAddThis = strAddThis & "<a href=""http://www.addthis.com/bookmark.php?v=250&amp;username=xa-4ca4f3320aef5a96"" class=""addthis_button_compact"">Share</a>" & vbCrLf
                        strAddThis = strAddThis & "</div>" & vbCrLf
                        strAddThis = strAddThis & "<script type=""text/javascript"" src=""http://s7.addthis.com/js/250/addthis_widget.js#username=xa-4ca4f3320aef5a96""></script>" & vbCrLf
				End Select
				
				strAddThis = strAddThis & "<!--URL:" & strURL & "|Title:" & strTitle & "-->" & vbCrLf
				strAddThis = strAddThis & "<!-- AddThis Bookmark Button END -->" & vbCrLf
			Else
				strAddThis = "<!--AddThis is switched off-->"
			End If
			
			AddThis = strAddThis
		End Function
		
		Public Function GoogleAnalytics()
			'string concatenation is acceptable here, due to the small size of string
			Dim strGA
			
			strGA = vbCrLf
			
			If Application("GoogleAnalyticsID") <> "" Then
				strGA = strGA & "<!--Google Analytics BEGIN-->" & vbCrLf
				strGA = strGA & "<script type=""text/javascript"">" & vbCrLf
				strGA = strGA & "<!-- //" & vbCrLf
				strGA = strGA & vbTab & "var gaJsHost = ((""https:"" == document.location.protocol) ? ""https://ssl."" : ""http://www."");" & vbCrLf
				strGA = strGA & vbTab & "document.write(unescape(""%3Cscript src='"" + gaJsHost + ""google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E""));" & vbCrLf
				strGA = strGA & "// -->" & vbCrLf
				strGA = strGA & "</script>" & vbCrLf
				strGA = strGA & "<script type=""text/javascript"">" & vbCrLf
				strGA = strGA & vbTab & "var pageTracker = _gat._getTracker(""" & Application("GoogleAnalyticsID") & """);" & vbCrLf
				strGA = strGA & vbTab & "pageTracker._trackPageview();" & vbCrLf
				strGA = strGA & "</script>" & vbCrLf
				strGA = strGA & "<!--Google Analytics END-->" & vbCrLf
			End If
			
			GoogleAnalytics = strGA
		End Function

		Public Sub SendEmail(strFrom, strTo, strSubject, strBody) 'move to /slickcms/messaging.asp?
			'sends an email using CDO.Message
			Dim cdoMessage
			Dim cdoConfig

			strFrom = ValidateEmail(strFrom)
			strTo = ValidateEmail(strTo)
            
            If Application("Debug") = 1 Then
			    Response.Write(strFrom & "<br />" & strTo)
            End If
			
			If (Len(strFrom) = 0) Or (Len(strTo) = 0) Then
				Session("EmailSent") = false
				Exit Sub
			End If

            Set cdoConfig = CreateObject("CDO.Configuration")  
            Set cdoMessage = Server.CreateObject("CDO.Message")

            With cdoConfig.Fields  
                .Item(cdoSendUsingMethod) = 2
                .Item(cdoSMTPServer) = Application("Mail_SMTP")
                .Item(cdoSMTPServerPort) = Application("Mail_Port")
                
                'use SMTP authentication if details provided
                If Application("Mail_User") <> "" And Application("Mail_Password") <> "" Then
                    .Item(cdoSMTPAuthenticate) = 1
                    .Item(cdoSendUsername) = Application("Mail_User")
                    .Item(cdoSendPassword) = Application("Mail_Password")
                Else
                    .Item(cdoSMTPAuthenticate) = 0
                End If

                .Update  
            End With 
			
            With cdoMessage 
                Set .Configuration = cdoConfig 
                .From = strFrom
                .To = strTo
                .Subject = strSubject
                
                If Application("Mail_Type") = "PLAIN" Then
				    .TextBody = strBody
			    Else
				    .HTMLBody = strBody
			    End If
                
                .Send 
            End With

			Set cdoMessage = Nothing
			Set cdoConfig = Nothing 
			
			If err.number <> 0 Then
				Session("EmailSent") = false
			Else
				Session("EmailSent") = true
			End If
		End Sub
		
		Public Function CSS()
		    'determines which additional style sheet to call, depending on the user's browser - which allows for browser specific styles
			Dim strCSS
			Dim strMedia
			Dim strUserAgent
			Dim strReturn
			
			strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
			strMedia = "screen"
			strCSS = ""

			If InStr(strUserAgent,"MSIE 6.0")>0 Then
				strCSS = "ie6.css"
			ElseIf InStr(strUserAgent,"MSIE 7.0")>0 Then
				strCSS = "ie7.css"
			ElseIf InStr(strUserAgent,"MSIE 8.0")>0 Then
				strCSS = "ie8.css"
			ElseIf InStr(strUserAgent,"Chrome")>0 Then
				strCSS = "chrome.css"	
			ElseIf InStr(strUserAgent,"Safari")>0 Then
				strCSS = "safari.css"
			ElseIf InStr(strUserAgent,"Opera")>0 Then
				strCSS = "opera.css"
			ElseIf InStr(strUserAgent,"Firefox")>0 Then
				strCSS = "firefox.css"
			End If

            If Application("Debug") = 1 Then
    			Response.Write("<!--User Agent:" & Request.ServerVariables("HTTP_USER_AGENT") & "-->" & vbcrlf)
            End If

			If strCSS <> "" Then
				strReturn = ("<link href=""" & Application("CDN") & "css/" & strCSS & """ rel=""stylesheet"" type=""text/css"" media=""" & strMedia & """ />")
			End If
			
			CSS = strReturn
		End Function
		
		Public Sub Log(strMessage)
		    'logs a message to the file system
			Dim objFS
			Dim objFile
		
			Set objFS = Server.CreateObject("Scripting.FileSystemObject")
			Set objFile = objFS.OpenTextFile(Application("ErrorLog"), 8, True)
		
			objFile.WriteLine(strMessage)
			objFile.Close
		
			Set objFile = Nothing
			Set objFS = Nothing
		End Sub
		
        Public Function AdminPaging(intRecords, intPagination, strPage)
		    Dim intPages 'total number of pages
		    Dim intPage
		    Dim intRemainder
		    Dim strPaging
		    
		    'calculate how many pages for pagination e.g. 21/5 = 4.2
		    intPages = Round(intRecords / Application("AdminPagination"))
		    
		    'check if there's any remainder e.g. 21-(4*5) = 1
		    intRemainder = (intRecords - (intPages * Application("AdminPagination")))
		    
		    'if there's any remainder, then add a final page
		    If intRemainder > 0 Then intPages = (intPages + 1)
            
            'start of paging string
            strPaging = "<div class=""pagination"">" & vbCrLf
            
            'add the Previous link
            If intPagination <> 1 And intPages > 1 Then
                strPaging = strPaging & ("<span><a href=""/admin/" & strPage & "?page=" & (intPagination - 1) & """>&#171; Previous</a></span> ") & vbCrLf
            Else
                strPaging = strPaging & ("<span class=""disabled""><a href=""javascript:void(0);"">&#171; Previous</a></span> ") & vbCrLf
            End If
		    
		    'loop through pages adding a link to the paging string
		    For intPage = 1 To intPages
                'add a class for styling the current page
                If CInt(intPagination) = CInt(intPage) Then
                    strPaging = strPaging & ("<span class=""current""><a href=""javascript:void(0);"">" & intPage & "</a></span>") & vbCrLf
                Else
                    strPaging = strPaging & ("<span><a href=""/admin/" & strPage & "?page=" & intPage & """>" & intPage & "</a></span> ") & vbCrLf
                End If
		    Next
		    
		    'add the Next link
		    If intPagination <> intPages Then
		        strPaging = strPaging & ("<span><a href=""/admin/" & strPage & "?page=" & (intPagination + 1) & """>Next &#187;</a></span> ") & vbCrLf
            Else
                strPaging = strPaging & ("<span class=""disabled""><a href=""javascript:void(0);"">Next &#187;</a></span> ") & vbCrLf
            End If
		    
		    'end of paging string
		    strPaging = strPaging & "</div>" & vbCrLf
		    
		    'hide pagination if not applicable
		    If intPages = 1 Then strPaging = ""

		    AdminPaging = strPaging
		End Function
		
        Public Sub Meta()
            'used for generating the <title> and <meta> tags of the HTML document
            Select Case objPost.UrlType
                Case "post","date"
                    strTitle = objPost.Title
                    strDescription = objPost.Summary
                Case "archive"
                    strDescription = "All Posts published in "
                    If dMonth <> "" And dDay = "" Then
                        'monthly archive
                        strTitle = dYear & "-" & dMonth & " archive"
                        strDescription = strDescription & dYear & "-" & dMonth
                    ElseIf dDay <> "" Then
                        strTitle = dYear & "-" & dMonth & "-" & dDay & " archive"
                        strDescription = strDescription & dYear & "-" & dMonth & "-" & dDay
                    Else
                        strTitle = dYear & " archive"
                        strDescription = strDescription & dYear
                    End If
                Case "tag"
                    strTitle = objPost.Url & " tag"
                    strDescription = "All posts tagged with " & objPost.Url
                Case "category"
                    strTitle = objPost.Url & " category"
                    strDescription = "All posts categorised under " & objPost.Url
                Case "pagination"
                    If objPost.Pagination = 0 Then
                        strTitle = "Home"
                    Else
                        strTitle = "Page " & objPost.Pagination
                    End If
            End Select
            
            If objPost.Pagination <> 0 And objPost.UrlType <> "pagination" Then
                'tag on the Pagination if not already present and required
                strTitle = strTitle & " | Page " & objPost.Pagination
                strDescription = strDescription & ", Page " & objPost.Pagination
            End If

            'append the global suffix for the <title>
            strTitle = strTitle & " | " & Application("TitleTag")
            
            'if blank, set a generic global description
            If strDescription = "" Then strDescription = Application("DescriptionTag")
            
            'ensure title is < 65 characters in length and description is < 150 (for SEO standards)
            If Len(strTitle)>65 Then strTitle = Left(strTitle,62) & "..."
            If Len(strDescription)>150 Then strDescription = Left(strDescription,147) & "..."
        End Sub
		
	'private methods
		Function ValidateEmail(strEmail)
			Dim strReturn

			Set m_objClean = New Clean

			m_objClean.Data = strEmail
			m_objClean.Email()
			m_objClean.MaxLength(1024)
			strReturn = m_objClean.Data

			Set m_objClean = Nothing

			If InStr(strReturn,"@")=0 Then strReturn = ""
			If InStr(strReturn,".")=0 Then strReturn = ""
			If InStr(strReturn,"@")=Len(strReturn) Then strReturn = ""
			If InStr(strReturn,".")=Len(strReturn) Then strReturn = ""
			
			ValidateEmail = strReturn
		End Function
End Class
%>