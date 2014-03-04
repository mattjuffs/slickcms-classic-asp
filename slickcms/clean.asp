<%
Class Clean
	'Used for data validation, to prevent SQL injection, XSS and errors etc.

	'public properties - let (set)
		Public Property Let Data(p_Data)
			m_Data = p_Data
		End Property

	'public properties - get (retrieve)		
		Public Property Get Data()
			Data = m_Data
		End Property

	'private properties
		Private m_Data 'data that the methods process
		Private m_objRegEx 'regular expression object

	'public methods
        Public Sub Alpha()
			'Alpha whitelist
			If m_Data <> "" Then
				Set m_objRegEx = new RegExp
				m_objRegEx.IgnoreCase = True
				m_objRegEx.Global = True
				m_objRegEx.Pattern = "[^-a-zA-Z]"
				m_Data = cstr(m_Data)
				m_Data = m_objRegEx.Replace(m_Data,"")
				Set m_objRegEx = Nothing
			Else
				m_Data = ""
			End If
		End Sub
		
		Public Sub Numeric_Old()
		    'no longer used (replace with Numeric test below)
			'Numeric whitelist
			If m_Data <> "" Then
				Set m_objRegEx = new RegExp
				m_objRegEx.IgnoreCase = True
				m_objRegEx.Global = True
				m_objRegEx.Pattern = "[^0-9-]"
				m_Data = CStr(m_Data)				
				m_Data = m_objRegEx.Replace(m_Data,"")				
				If m_Data = "" Then m_Data = 0
				If m_Data = "-" Then m_Data = 0
				m_Data = CInt(m_Data)
				Set m_objRegEx = Nothing
			Else
				m_Data = 0
			End If
		End Sub
		
		Public Sub Numeric()
			'positive/negative number regex test
			If m_Data <> "" Then
				Set m_objRegEx = new RegExp
				m_objRegEx.IgnoreCase = True
				m_objRegEx.Global = True
				m_objRegEx.Pattern = "^-{0,1}\d*\.{0,1}\d+$"
				If m_objRegEx.Test(m_Data) = True Then
				    m_Data = m_Data
				Else
				    m_Data = 0
				End If
				Set m_objRegEx = Nothing
			Else
				m_Data = 0
			End If
		End Sub

		Public Sub AlphaNumeric()
			'AlphaNumeric whitelist
			If m_Data <> "" Then
				Set m_objRegEx = new RegExp
				m_objRegEx.IgnoreCase = True
				m_objRegEx.Global = True
				m_objRegEx.Pattern = "[^-a-zA-Z0-9-;,.' ]"
				m_Data = cstr(m_Data)
				m_Data = m_objRegEx.Replace(m_Data,"")
				Set m_objRegEx = Nothing
			Else
				m_Data = ""
			End If
		End Sub
	
		Public Sub NumericPlus()
			'NumericPlus whitelist
			If m_Data <> "" Then
				Set m_objRegEx = new RegExp
				m_objRegEx.IgnoreCase = True
				m_objRegEx.Global = True
				m_objRegEx.Pattern = "[^0-9\(\)\-ext. ]"
				m_Data = cstr(m_Data)
				m_Data = m_objRegEx.Replace(m_Data,"")
				Set m_objRegEx = Nothing
			Else
				m_Data = ""
			End If
		End Sub
	
		Public Sub Email()
			'Email whitelist
			If m_Data <> "" Then
				Set m_objRegEx = new RegExp
				m_objRegEx.IgnoreCase = True
				m_objRegEx.Global = True
				m_objRegEx.Pattern = "[^-a-zA-Z0-9@.!$&*-=^`|~#%'+/?_{}]"
				m_Data = cstr(m_Data)
				m_Data = m_objRegEx.Replace(m_Data,"")
				Set m_objRegEx = Nothing
			Else
				m_Data = ""
			End If
		End Sub
		
		Public Sub MaxLength(intLength)
			'checks if a string is less than a specified length, otherwise it chops off the excess
			If len(m_Data) > intLength Then
				m_Data = left(m_Data,intLength)
			Else
				m_Data = m_Data
			End If
		End Sub
	
		Public Sub Encode()
			'HTML encodes a string
			If m_Data <> "" Then
				m_Data = Server.HTMLEncode(m_Data)
			Else
				m_Data = ""
			End If
		End Sub
		
		Public Sub StripHTML()
            Set m_objRegEx = New RegExp

            m_objRegEx.IgnoreCase = True
            m_objRegEx.Global = True
            m_objRegEx.Pattern = "<(.|\n)+?>"

            m_Data = m_objRegEx.Replace(m_Data, "")
            m_Data = Replace(m_Data, "<", "&lt;")
            m_Data = Replace(m_Data, ">", "&gt;")

            Set m_objRegEx = Nothing
		End Sub
		
		Public Sub SQL()
		    'For dynamic SQL Parameters
			If m_Data <> "" Then
				Set m_objRegEx = new RegExp
				m_objRegEx.IgnoreCase = True
				m_objRegEx.Global = True
				m_objRegEx.Pattern = "[^a-zA-Z0-9, ]"
				m_Data = cstr(m_Data)
				m_Data = m_objRegEx.Replace(m_Data,"")
				Set m_objRegEx = Nothing
			Else
				m_Data = ""
			End If
		End Sub
		
        Public Sub Url()
			'AlphaNumeric whitelist
			If m_Data <> "" Then
				Set m_objRegEx = new RegExp
				m_objRegEx.IgnoreCase = True
				m_objRegEx.Global = True
				m_objRegEx.Pattern = "[^-a-zA-Z0-9-/ ]"
				m_Data = cstr(m_Data)
				m_Data = m_objRegEx.Replace(m_Data,"")
				Set m_objRegEx = Nothing
			Else
				m_Data = ""
			End If
		End Sub
End Class
%>