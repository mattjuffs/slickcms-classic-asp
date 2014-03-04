<%
Class Image
	'public properties - let (set)
		Public Property Let Album(p_Album)
			Set m_objClean = New Clean
			m_objClean.Data = Replace(p_Album,"%20"," ")
			Call m_objClean.AlphaNumeric()
			m_Album = m_objClean.Data
			m_Album = m_objClean.MaxLength(255)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Template(p_Template)
			m_Template = p_Template
		End Property

	'public properties - get (retrieve)
		Public Property Get Album()
			Album = m_Album
		End Property
		
	'private properties
		Private m_Album
		Private m_Template
		
		Private m_strSQL
		Private m_objRS
		Private m_objClean

	'public methods
		Public Sub Load()
			Dim strImage, strAlt, strTitle, strClass
			Dim strLarge, strThumb

			m_strSQL = "Select Name, Alt, Title, Orientation From Images Where Album = '" & m_Album & "' Order By Orientation Asc"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strImage = m_objRS.Fields("Name").Value
					strAlt = m_objRS.Fields("Alt").Value
					strTitle = m_objRS.Fields("Title").Value
					strClass = lcase(m_objRS.Fields("Orientation").Value)
					strLarge = "/images/" & m_Album & "/" & strImage
					strThumb = "/images/" & m_Album & "/thumbnails/" & strImage

					strReturn = Replace(m_Template,"[large]",strLarge)
					strReturn = Replace(strReturn,"[thumb]",strThumb)
					strReturn = Replace(strReturn,"[alt]",strAlt)
					strReturn = Replace(strReturn,"[title]",strTitle)
					strReturn = Replace(strReturn,"[class]",strClass)

					Response.Write(strReturn)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
		End Sub

	'private methods
		'none
End Class
%>