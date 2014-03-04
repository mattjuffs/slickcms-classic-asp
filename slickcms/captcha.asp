<%
Class Captcha
    'Currently uses OpenCaptcha: http://opencaptcha.com/
    'This class is very specific and could be adapted in future to work with multiple captcha service providers
    
	'public properties - let (set)
	    'none

	'public properties - get (retrieve)
		'none
		
	'private properties
		'none

	'public methods
        Function Generate(intHeight, intWidth)
            'generates a captcha
            Dim strImageName
            Dim strRandom
            Dim strSize
            Dim strSite

            'name of the site
            strSite = lcase(Replace(Application("SiteName")," ",""))
            
            'random string using date and seconds since midnight
            strRandom = Replace(GetDate() & CStr(Timer()),".","")
            
            'size of the captcha
            strSize = "-" & CStr(intHeight) & "-" & CStr(intWidth)
            
            'piece together
            strImageName = strRandom & strSite & strSize & ".jpgx"
            
            Generate = strImageName
        End Function

        Function Process()
            'processes a captcha in form post
            Dim strURL
            Dim strImageName
            Dim strAnswer
            Dim strResult
            
            strImageName = Request.Form("opencaptcha")

            '#449 - Captcha provider may be down, so bypass
            If strImageName <> "" Then
                strAnswer = Request.Form("answer")
                strURL = "http://www.opencaptcha.com/validate.php?img=" & strImageName & "&ans=" & strAnswer
                strResult = GetResult(strURL)
                If strResult = "" Then strResult = "pass"
            Else
                strResult = "pass"
            End If
            
            Process = strResult
        End Function

	'private methods
        Function GetDate()
            'gets the current date in a specific format
	        Dim strDate
	        Dim intYear
	        Dim intMonth
	        Dim intDay
	        Dim intHour
	        Dim intMinute
	        Dim strReturn

	        strDate = Now()
	        intYear = DatePart("yyyy", strDate)
	        intMonth = DatePart("m", strDate)
	        intDay = DatePart("d", strDate)
	        intHour = DatePart("h", strDate)
	        intMinute = DatePart("n", strDate)

	        'ensure month/day is in MM/DD format
	        If intMonth < 10 Then intMonth = "0" & intMonth
	        If intDay < 10 Then intDay = "0" & intDay
        	
            strReturn = intYear & intMonth & intDay & intHour & intMinute
            
            GetDate = strReturn
        End Function
        
        Function GetResult(strURL)
            'posts the captcha answer and returns the pass/fail result
            Dim objXMLHTTP, strResponse, intStatus

            On Error Resume Next
            
            Response.Buffer = True

            Set objXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")

            objXMLHTTP.setTimeouts 5000, 10000, 10000, 10000
            objXMLHTTP.Open "GET", strURL, False
            objXMLHTTP.Send
            intStatus = objXMLHTTP.status

            If err.number <> 0 Or intStatus <> 200 Then
                strResponse = ""
            Else
                strResponse = objXMLHTTP.responseText
            End If

            Set objXMLHTTP = Nothing

            GetResult = strResponse
        End Function
End Class
%>