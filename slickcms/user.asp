<%
Class User
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
		
		Public Property Let Password(p_Password)
			If Len(p_Password)>0 Then
				Set m_objClean = New Clean
				m_objClean.Data = p_Password
				Call m_objClean.MaxLength(1024)
				m_Password = m_objClean.Data
				Set m_objClean = Nothing

				m_Password = MD5(p_Password)
			Else
				m_Password = " " 'bypass updating password
			End If
		End Property
		
		Public Property Let URL(p_URL)
			Set m_objClean = New Clean
			m_objClean.Data = p_URL
			Call m_objClean.MaxLength(1024)
			m_URL = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let IP(p_IP)			
			Set m_objClean = New Clean
			m_objClean.Data = p_IP
			Call m_objClean.NumericPlus()
			Call m_objClean.MaxLength(15)
			m_IP = m_objClean.Data
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Biography(p_Biography)
			m_Biography = p_Biography
		End Property
		
		Public Property Let Active(p_Active)
			Set m_objClean = New Clean
			m_objClean.Data = p_Active
			Call m_objClean.Numeric()
			m_Active = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let LoginFails(p_LoginFails)
			Set m_objClean = New Clean
			m_objClean.Data = p_LoginFails
			Call m_objClean.Numeric()
			m_LoginFails = CInt(m_objClean.Data)
			Set m_objClean = Nothing
		End Property
		
		Public Property Let Template(p_Template)
			m_Template = p_Template
		End Property
		
		Public Property Let AuthorsTemplate(p_AuthorsTemplate)
			m_AuthorsTemplate = p_AuthorsTemplate
		End Property
		
		Public Property Let AuthorsTemplateSelected(p_AuthorsTemplateSelected)
			m_AuthorsTemplateSelected = p_AuthorsTemplateSelected
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

		Public Property Get Email()
			Email = m_Email
		End Property
		
		Public Property Get URL()
			URL = m_URL
		End Property
		
		Public Property Get IP()
			IP = m_IP
		End Property
		
		Public Property Get Biography()
			Biography = m_Biography
		End Property
		
		Public Property Get DateCreated()
			DateCreated = m_DateCreated
		End Property
		
		Public Property Get DateModified()
			DateModified = m_DateModified
		End Property
		
		Public Property Get Active()
			Active = m_Active
		End Property
		
		Public Property Get LoginFails()
			LoginFails = m_LoginFails
		End Property
		
		Public Property Get Pagination()
			Pagination = m_Pagination
		End Property
		
		Public Property Get AdminUsersCount()
		    AdminUsersCount = m_AdminUsersCount
        End Property
		
	'private properties
		Private m_ID
		Private m_Name
		Private m_Email
		Private m_Password
		Private m_URL
		Private m_IP
		Private m_Biography
		Private m_DateCreated
		Private m_DateModified
		Private m_Active
		Private m_LoginFails
		Private m_Template
		Private m_AuthorsTemplate
		Private m_AuthorsTemplateSelected
		
		Private m_objSlickCMS
		Private m_objCmd
		Private m_strSQL
		Private m_objRs
		Private m_objClean
		
		'pagination variables
		Private m_Start
		Private m_End
        Private m_Pagination
        Private m_AdminUsersCount

	'public methods
		Public Sub CheckLogin()
		    'ensure user is logged on, otherwise redirect
			If Session("LoggedOn") <> true Then
				Session("Error") = Application("CheckLoginFailed")
				Response.Redirect("/admin/login.asp")
				Response.End
			End If
		End Sub
		
		Public Sub Login()
            'authenticates a user's credentials
			Dim rs, strSQL, strRedirect, intAdmin

			strSQL = "Execute dbo.[Login] '" & m_Email & "','" & m_Password & "'"	
			Set rs = Server.CreateObject("ADODB.RecordSet")
			
			rs.Open strSQL,objConn,0,1
			
			If Not rs.BOF Then
				If rs.Fields("LoginFails") < Application("LockLimit") Then
					Session("Name") = rs.Fields("Name").Value

					'2009-10-01 added for #63
					Session("UserID") = rs.Fields("UserID").Value
					Session("Email") = rs.Fields("Email").Value
					Session("URL") = rs.Fields("URL").Value
					
					Session("LoggedOn") = true
					Session("Message") = Replace(Application("LogOnSuccess"),"[name]",Session("Name"))
					strRedirect = "/admin/"

					'for #179
					'If rs.Fields("Admin").Value = "1" Then Session("Admin") = "1"	
				Else
					Session("LoggedOn") = false
					Session("Error") = Application("AccountLocked")
					strRedirect = "/admin/login.asp"
				End If
			Else
				Session("LoggedOn") = false
				Session("Error") = Application("LogOnFail")
				
				Call IncrementLoginCounter()
				
				Set m_objSlickCMS = New SlickCMS
				Call m_objSlickCMS.Log("Invalid email/password:" & m_Email & "/" & m_Password)
				Set m_objSlickCMS = Nothing

				strRedirect = "/admin/login.asp"
			End If
			
			If rs.State <> 0 Then rs.Close
			Set rs = Nothing
			
			Response.Redirect(strRedirect)
			Response.End
		End Sub
		
		Public Sub LogOut()
			'clear all session variables and takes the logged out user to the homepage
			Session.Contents.RemoveAll()
			Session("Message") = Application("LoggedOut")
			Response.Redirect("/")
			Response.End
		End Sub
		
		Public Sub GetAdminUsers()
		    'retrieves a list of Users for use within the Admin
			Dim strTemplate
			Dim strActive

			'count the users for pagination use
			m_strSQL = "Select Count(*) From dbo.Users"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			If Not m_objRS.BOF Then
			    m_AdminUsersCount = m_objRS.Fields(0).Value
            Else
                m_AdminUsersCount = 0
            End If
            If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing
			
            'pagination
			If m_Pagination = 0 Then m_Pagination = 1
			m_End = (Application("AdminPagination") * m_Pagination)
			m_Start = (m_End - Application("AdminPagination"))+1
			If m_Start = 0 Then m_Start = 1

			m_strSQL = "Execute Admin_SelectUsers " & m_Start & "," & m_End
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF
					strTemplate = m_Template
					
					m_ID = m_objRS.Fields("UserID").Value
					m_Name = m_objRS.Fields("Name").Value
					m_Email = m_objRS.Fields("Email").Value
					m_URL = m_objRS.Fields("URL").Value
					m_DateModified = m_objRS.Fields("DateModified").Value
					m_Active = m_objRS.Fields("Active").Value
					m_LoginFails = m_objRS.Fields("LoginFails").Value
					
					Select Case m_Active
						Case 0
							strActive = "No"
						Case 1
							strActive = "Yes"
						Case Else
							strActive = "No"
					End Select
					
					If m_LoginFails > Application("LockLimit") Then
						strTemplate = Replace(strTemplate,"[class]"," class=""locked""")
					Else
						strTemplate = Replace(strTemplate,"[class]","")
					End If
					
					strTemplate = Replace(strTemplate,"[userid]",m_ID)
					strTemplate = Replace(strTemplate,"[name]",m_Name)
					strTemplate = Replace(strTemplate,"[email]",m_Email)
					strTemplate = Replace(strTemplate,"[url]",m_URL)
					strTemplate = Replace(strTemplate,"[datemodified]",m_DateModified)
					strTemplate = Replace(strTemplate,"[active]",strActive)
					strTemplate = Replace(strTemplate,"[loginfails]",m_LoginFails)

					Response.Write(strTemplate)

					m_objRS.MoveNext
				Loop
			End If
			
			If m_objRS.State <> 0 Then m_objRS.Close
			Set m_objRS = Nothing	
		End Sub
		
		Public Sub GetAdminUser()
			If m_ID <> 0 Then
				m_strSQL = "Execute Admin_SelectUser " & m_ID
				Set m_objRS = Server.CreateObject("ADODB.RecordSet")
				m_objRS.Open m_strSQL,objConn,0,1
				
				If Not m_objRS.BOF Then
					m_ID = m_objRS.Fields("UserID").Value
					m_Name = m_objRS.Fields("Name").Value
					m_Email = m_objRS.Fields("Email").Value
					m_URL = m_objRS.Fields("URL").Value
					m_IP = m_objRS.Fields("IP").Value
					m_Biography = m_objRS.Fields("Biography").Value
					m_DateCreated = m_objRS.Fields("DateCreated").Value
					m_DateModified = m_objRS.Fields("DateModified").Value
					m_Active = m_objRS.Fields("Active").Value
					m_LoginFails = m_objRS.Fields("LoginFails").Value
				Else
					m_ID = 0
				End If
				
				If m_objRS.State <> 0 Then m_objRS.Close
				Set m_objRS = Nothing
			End If
			
			If m_ID = 0 Then
				m_Name = ""
				m_Email = ""
				m_URL = ""
				m_IP = ""
				m_Biography = ""
				m_DateCreated = ""
				m_DateModified = ""
				m_Active = 0
				m_LoginFails = 0
			End If
		End Sub
		
		Public Sub Save()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			If m_ID = 0 Then
			    'adding new User via Admin
				m_objCmd.CommandText = "Admin_InsertUser"
			Else
			    'updating existing User via Admin
				m_objCmd.CommandText = "Admin_UpdateUser"
				m_objCmd.Parameters.Append m_objCmd.CreateParameter("@UserID", 3, 1, , m_ID)
			End If

			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Name", 200, 1, 50, m_Name)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Email", 200, 1, 255, m_Email)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Password", 200, 1, 32, m_Password)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@URL", 200, 1, 1024, m_URL)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@IP", 200, 1, 15, m_IP)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Biography", 200, 1, len(m_Biography), m_Biography)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Active", 3, 1, , m_Active)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@LoginFails", 3, 1, , m_LoginFails)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			If m_ID = 0 Then
				Response.Write(Application("UserSaved"))
			Else
				Response.Write(Application("UserUpdated"))
			End If
		End Sub
		
		Public Sub Delete()
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4

			m_objCmd.CommandText = "Admin_DeleteUser"
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@UserID", 3, 1, , m_ID)

			m_objCmd.Execute
			Set m_objCmd = Nothing

			Response.Write(Application("UserDeleted"))
		End Sub
		
		Public Sub Register()
		    Dim strBody

			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4
			m_objCmd.CommandText = "Admin_InsertUser"
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Name", 200, 1, 50, m_Name)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Email", 200, 1, 255, m_Email)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Password", 200, 1, 32, m_Password)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@URL", 200, 1, 1024, m_URL)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@IP", 200, 1, 15, m_IP)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Biography", 200, 1, 3, "n/a")
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Active", 3, 1, , 0)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@LoginFails", 3, 1, , 0)

			m_objCmd.Execute
			Set m_objCmd = Nothing
			
			'send an email to notify of user registration
			strBody = Replace(Application("UserRegisteredEmailBody"),"[name]",m_Name)
	        strBody = Replace(strBody,"[email]",m_Email)
	        
            Set m_objSlickCMS = New SlickCMS
            Call m_objSlickCMS.SendNotification("New User registration", strBody)
            Set m_objSlickCMS = Nothing
			
			Session("Message") = Application("UserRegisterSuccess")
			Response.Redirect("/")
			Response.End
		End Sub
		
		Public Sub ResetPassword()
		    'generates an 8 character random password and resets a User's password to it, then emails the User the new password
			Dim strBody
			m_Password = RandomPassword(8)
			
			strBody = Replace(Application("UserPasswordResetEmail"),"[password]",m_Password)
			
			Call UpdatePassword()
			
			Set m_objSlickCMS = New SlickCMS
			Call m_objSlickCMS.SendEmail(Application("Email"), m_Email, "Password Reset", strBody)
			Set m_objSlickCMS = Nothing

			If Session("EmailSent") = true Then
				Session("Message") = Application("UserPasswordReset")
			Else
				Session("Error") = Application("UserPasswordResetError")
			End If
		End Sub
		
		Public Sub GetAuthors()
		    'retrieves a list of Users for specifying a Post's author within the Admin
			m_strSQL = "Select [UserID], [Name] From dbo.Users Where [Active] = 1 Order By [Name] Asc"
			Set m_objRS = Server.CreateObject("ADODB.RecordSet")
			m_objRS.Open m_strSQL,objConn,0,1
			
			If Not m_objRS.BOF Then
				Do While Not m_objRS.EOF					
					If m_ID = m_objRS.Fields("UserID").Value Then
						Response.Write(Replace(Replace(m_AuthorsTemplateSelected,"[id]",m_objRS.Fields("UserID").Value),"[name]",m_objRS.Fields("Name").Value))
					Else
						Response.Write(Replace(Replace(m_AuthorsTemplate,"[id]",m_objRS.Fields("UserID").Value),"[name]",m_objRS.Fields("Name").Value))
					End If					

					m_objRS.MoveNext
				Loop
			End If

			Set m_objRS = Nothing
		End Sub

	'private methods
		Private Sub IncrementLoginCounter()
		    'increments the login counter if a User fails to login
			If Len(m_Email)>0 Then
				Set m_objCmd = Server.CreateObject("ADODB.Command")
	
				m_objCmd.ActiveConnection = objConn
				m_objCmd.CommandType = 4
	
				m_objCmd.CommandText = "Admin_IncrementLoginCounter"
				m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Email", 200, 1, 255, m_Email)
	
				m_objCmd.Execute
				Set m_objCmd = Nothing
			End If
		End Sub
		
		Private Function RandomPassword(myLength)
		    'generates a random password
			Dim X, Y, strPW
			
			Const minLength = 8
			Const maxLength = 32
			
			If myLength = 0 Then
				Randomize
				myLength = Int((maxLength * Rnd) + minLength)
			End If
			
			For X = 1 To myLength
				'Randomize the type of this character
				Y = Int((3 * Rnd) + 1) '(1) Numeric, (2) Uppercase, (3) Lowercase
				
				Select Case Y
					Case 1
						'Numeric character
						Randomize
						strPW = strPW & CHR(Int((9 * Rnd) + 48))
					Case 2
						'Uppercase character
						Randomize
						strPW = strPW & CHR(Int((25 * Rnd) + 65))
					Case 3
						'Lowercase character
						Randomize
						strPW = strPW & CHR(Int((25 * Rnd) + 97))
				End Select
			Next
			
			RandomPassword = strPW
		End Function
		
		Private Sub UpdatePassword()
		    'updates a User's password with an MD5 hash of it - no plaintext passwords are stored within SlickCMS
			m_Password = MD5(m_Password)
			
			Set m_objCmd = Server.CreateObject("ADODB.Command")

			m_objCmd.ActiveConnection = objConn
			m_objCmd.CommandType = 4
			m_objCmd.CommandText = "Admin_UpdatePassword"
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Email", 200, 1, 255, m_Email)
			m_objCmd.Parameters.Append m_objCmd.CreateParameter("@Password", 200, 1, 32, m_Password)

			m_objCmd.Execute

			Set m_objCmd = Nothing
		End Sub
End Class
%>