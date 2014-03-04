<%
    'Not to be confused with global.asa, this page is used for Global variable, function and sub declarations

    'object declarations
    Dim objPost
	Dim objLink
	Dim objSlickCMS
	Dim objImage
	Dim objComment
	Dim objCaptcha
	Dim objCategory
	Dim objTag
	Dim objStatistic
	Dim objConn

    'global declarations
	Dim strAlbum
	Dim strTemplate
	Dim strPagination
	Dim strCaptcha
	Dim strCategories
	Dim strTags
	Dim strTitle, strDescription 'meta tags
	Dim intPosts

    'UrlHandler() declarations
    Dim strUrl
    Dim aUrl 'array for Url
    Dim strUrlType
    Dim dYear, dMonth, dDay 'date parts
    Dim intPagination

    'global functions
	Function OpenDatabase()
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open Application("ConnectionString")
	End Function
	
	Function CloseDatabase()
		objConn.Close
		Set objConn = Nothing
	End Function
	
	'2009-10-01 #63
	Function IIf(bExpression, strTrue, strFalse)
	    If bExpression Then
	        IIf = strTrue
        Else
            IIf = strFalse
        End If
	End Function
	
	Function ThemeFile(strFile)
	    'used for retrieving Theme files from within the Theme's folder
	    'thanks to: http://www.motobit.com/tips/detpg_read-write-binary-files/
	    'postponed from #67
        Dim strReturn
        Dim BinaryStream
        Dim strFilePath
        Dim objFSO, bFileExists
        
        'theme files are HTML
        strFile = strFile & ".html"
        
        'could update to use Session("Theme") for visitor to change their theme
	    strFilePath = Server.MapPath("/themes/" & Application("Theme") & "/" & strFile)
	    
	    'determine if the file exists
	    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	    If objFSO.FileExists(strFilePath) Then
	        bFileExists = true
        Else
            bFileExists = false
        End If
        Set objFSO = Nothing

        If bFileExists = true Then
            'create Stream object
            Set BinaryStream = CreateObject("ADODB.Stream")

            'specify stream type - we want to get binary data.
            BinaryStream.Type = 2

            'specify charset for the source text (unicode) data
            'see: http://msdn.microsoft.com/en-us/library/ms526296(EXCHG.10).aspx for full list of CharSets
            BinaryStream.CharSet = "utf-8"

            'open the stream
            BinaryStream.Open

            'load the file data from disk To stream object
            BinaryStream.LoadFromFile strFilePath

            'open the stream And get binary data from the object
            strReturn = BinaryStream.ReadText
            
            Set BinaryStream = Nothing
        Else
            'return an empty string
            strReturn = ""
        End If
        
        'return string of file contents
        ThemeFile = strReturn
    End Function
    
    Function HumanDate(dDate)
        If IsDate(dDate) Then
            'returns the date in a user friendly format, e.g. Tuesday 26th January 2010, 21:45pm
            Dim dYear
            Dim dMonth
            Dim dMonthName
            Dim dDay
            Dim dDayName
            Dim dDaySuffix
            Dim dTime
            Dim dTimeSuffix
            Dim dReturn

            dYear = DatePart("yyyy",dDate)
            dMonth = DatePart("m",dDate)
            dMonthName = MonthName(dMonth)
            dDay = DatePart("d",dDate)
            dDayName = weekdayname(weekday(dDate))

            Select Case dDay
                Case 1,21,31
                    dDaySuffix = "st"
                Case 2,22
                    dDaySuffix = "nd"
                Case 3,23
                    dDaySuffix = "rd"
                Case 4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,24,25,26,27,28,29,30
                    dDaySuffix = "th"
            End Select
            
            dTime = FormatDateTime(Now(),4)

            If DatePart("h",dDate) >= 12 Then
                dTimeSuffix = "pm"
            Else
                dTimeSuffix = "am"
            End If

            dReturn = dDayName & " " & dDay & dDaySuffix & " " & dMonthName & " " & dYear & ", " & dTime & dTimeSuffix
            
            HumanDate = dReturn
        Else
            HumanDate = dDate
        End If
    End Function
%>