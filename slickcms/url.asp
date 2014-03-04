<%
    'SlickCMS UrlHandler for SlickCMSRewriteURL

    Sub UrlHandler()
        'local variable declarations
        Dim m_objClean
        Dim m_objSlickCMS 'for logging
        Dim strUrlTemp 'temporary string used for evaluating
        Dim count 'for debug
        Dim aUrlTemp 'temporary array
        Dim bPhoneHome
        
        'initialise
        strUrl = ""
        intPagination = 0
        strUrlType = ""
        count = 0
        bPhoneHome = false

        'retrieve the raw Url and format to lowercase
        If Application("IIS_Version") >= 7 Then
            'IIS7+ using the Rewrite Module
            strUrl = lcase(Request.ServerVariables("HTTP_X_ORIGINAL_URL"))
        ElseIf Application("IIS_Version") = 6 Then
            'IIS6 using custom error
            strUrl = lcase(Request.ServerVariables("QUERY_STRING"))
            strUrl = Replace(strUrl,"404;","")
            strUrl = Replace(strUrl,":80","")
            strUrl = Replace(strUrl,Application("SiteURL"),"")
            If Left(strUrl,1) <> "/" Then strUrl = "/" & strUrl 'prefix with /
        Else
            strUrlType = "post"
            strUrl = "404"
            Exit Sub
        End If
        
        '#405 - redirect to homepage (possible hack attempts, or unwanted URLs
        If InStr(strUrl,"/cgi-bin/")>0 Then bPhoneHome = true
        If InStr(strUrl,"/www.whatismyip.com/")>0 Then bPhoneHome = true
        If InStr(strUrl,"/wp-includes/")>0 Then bPhoneHome = true
        If InStr(strUrl,".php")>0 Then bPhoneHome = true
        If InStr(strUrl,"sourcedir")>0 Then bPhoneHome = true
        
        If bPhoneHome = true Then
            Set m_objSlickCMS = New SlickCMS
            Call m_objSlickCMS.Log("Phoning Home: " & strUrl)
            Set m_objSlickCMS = Nothing

            Response.Redirect("/")
            Response.End
        End If
        
        '#405 - remove trackback from URL (for now)
        If InStr(strUrl,"/trackback/")>0 Then
            strUrl = Replace(strUrl,"/trackback/","/")
        End If
        
        '#419 - remove feed from URL
        If Right(strUrl,6) = "/feed/" Then
            strUrl = Left(strUrl,Len(strUrl)-5) 'keep the trailing slash
        End If
        
        If Application("Debug") = 1 Then
            Response.Write("URL:" & strUrl & "<br />")
        End If
               
        'force a trailing slash on the Url - #356
        If Right(strUrl,1) <> "/" And InStr(strUrl,".asp") = 0 Then
            'log
            Set m_objSlickCMS = New SlickCMS
            Call m_objSlickCMS.Log("Potential 301 redirect: " & strUrl)
            Set m_objSlickCMS = Nothing
            
            '301 redirect
            'Response.Status="301 Moved Permanently"
            'Response.AddHeader "Location", strUrl & "/"
            'Response.End
        End If

        If InStr(strUrl,"/page/")>0 Then
            'pagination present, set the page number
            aUrlTemp = Split(strUrl,"/page/")
            
            intPagination = aUrlTemp(UBound(aUrlTemp))
            intPagination = Replace(intPagination,"/","") 'remove trailing slash
            
            'return Url without pagination elements
            strUrlTemp = aUrlTemp(LBound(aUrlTemp)) & "/"
            aUrl = Split(strUrlTemp,"/")
            
            '#405 - deal with legacy search parameters (e.g. http://slickcms/page/2/?s=arty&submit_y=1)
            If InStr(intPagination,"?")>0 Then
                aUrlTemp = Split(intPagination,"?")
                intPagination = aUrlTemp(LBound(aUrlTemp))
            End If
        Else
            'split the Url into an Array
            aUrl = Split(strUrl,"/")
        End If
       
        'ascertain what type of Url it is by how many parts there are to it
        Select Case UBound(aUrl)
            Case 1, -1
                'http://slickcms/ - root/homepage
                If Application("Pagination") <> 0 Then
                    strUrlType = "pagination"
                Else
                    strUrlType = "post"
                End If
                strUrl = "home"
                
                'to allow for blogs to have a homepage
                If Application("Homepage") = 1 And intPagination = 0 Then
                    strUrlType = "post"
                Else
                    strUrlType = "pagination"
                End If
                strUrl = "home"
            Case 2
                Set m_objClean = New Clean
                m_objClean.Data = aUrl(1)
                Call m_objClean.Numeric()
                strUrlTemp = m_objClean.Data
                Set m_objClean = Nothing

                'if the numeric string is 4 digits, then it's considered a year
                If Len(strUrlTemp) = 4 Then
                    'http://slickcms/2009/ - yearly archives
                    strUrlType = "archive"
                    dYear = aUrl(1)
                    strUrl = dYear
                Else
                    'http://slickcms/ms-rl/ - posts
                    strUrlType = "post"
                    strUrl = aUrl(1)
                End If
            Case 3
                If aUrl(1) = "tag" Then
                    'http://slickcms/tag/computers/ - tags
                    strUrlType = "tag"
                    strUrl = aUrl(2)
                ElseIf aUrl(1) = "category" Then
                    'http://slickcms/category/name-of-category/ - categories
                    strUrlType = "category"
                    strURL = aUrl(2)
                ElseIf aUrl(1) = "posts" Then
                    'http://slickcms/posts/ms-rl/ - old style posts
                    strUrlType = "post"
                    strUrl = aUrl(2)
                ElseIf aUrl(1) = "page" Then
                    'http://slickcms/page/2/ - pagination pages, should already be caught by pagination check
                    'strUrlType = "pagination"
                    'strUrl = ""
                    'intPagination = aUrl(2)
                    Response.Write("you should never get here")
                    Response.End
                ElseIf InStr(aUrl(1),"-")>1 And UBound(aUrl)>2 Then
                    'http://slickcms/2009-12-15/wmd-test/ - date based posts
                    strUrlType = "date"
                    strUrl = aUrl(2)
                ElseIf Len(aUrl(1)) = 4 And Len(aUrl(2)) = 2 Then
                    'http://slickcms/2009/12/ - monthly archives
                    strUrlType = "archive"
                    dYear = aUrl(1)
                    dMonth = aUrl(2)
                    strUrl = dYear & "-" & dMonth
                End If
            Case 4
                'http://slickcms/2009/12/15/ - daily archives
                strUrlType = "archive"
                dYear = aUrl(1)
                dMonth = aUrl(2)
                dDay = aUrl(3)
                strUrl = dYear & "-" & dMonth & "-" & dDay

                '#405 - comments incorrectly formed (e.g. http://slickcms/2009-07-26/ms-rl/comment-3/)
                If InStr(aUrl(3),"comment-") Then
                    strUrlType = "date"
                    strUrl = aUrl(2)
                End If
            Case Else
                'unrecognised format
                strUrlType = "post"
                strUrl = "404"
        End Select

        If Application("Debug") = 1 Then
            Response.Write(LBound(aUrl) & ":" & UBound(aUrl) & "<br />")
            For Each strUrlTemp In aUrl
                Response.Write("aUrl(" & count & "):" & strUrlTemp & "<br />")
                count = (count + 1)
            Next
            Response.Write("<br />")
            Response.Write("Year:" & dYear & "<br />")
            Response.Write("Month:" & dMonth & "<br />")
            Response.Write("Day:" & dDay & "<br /><br />")
            Response.Write("URL:" & strUrl & "<br />")
            Response.Write("Type:" & strUrlType & "<br />")
            Response.Write("Pagination:" & intPagination)
        End If
        
        'set URL parameters of Post object
        objPost.Url = strUrl
        objPost.UrlType = strUrlType
        objPost.Year = dYear
        objPost.Month = dMonth
        objPost.Day = dDay
        objPost.Pagination = intPagination
    End Sub
%>