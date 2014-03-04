<div id="content">
	<%
		'messages			
		If (Session("EmailSent") <> Empty) And (Session("EmailSent") = false) Then
			Response.Write("<p class=""error"">Sorry, we were unable to send your message - please ensure all fields are filled in!</p>")
			Session("EmailSent") = Empty
		End If

		If Session("Message") <> Empty Then
			Response.Write("<p class=""message"">" & Session("Message") & "</p>")
			Session("Message") = Empty
		End If
		
		If Session("Error") <> Empty Then
			Response.Write("<p class=""error"">" & Session("Error") & "</p>")
			Session("Error") = Empty
		End If
		
		'used for meta data Category/Tag lists
		objCategory.Template = "<a href=""[url]"" title=""[description]"">[name]</a>, "
		objTag.Template = "<a href=""[url]"" title=""View all Posts tagged with [name]"">[name]</a>, "
       
        Select Case objPost.UrlType
            Case "post","date" 'single item
                Response.Write("<h1> " & objPost.Title & "</h1>")
                Response.Write(objPost.Content)

                If objPost.URL = "contact" Then
                %>
                    <!--#include virtual="/contact.asp"-->
                <%
                End If

                'photos
                Select Case objPost.Url
                    Case "album-one"
	                    strAlbum = "albumone"
                    Case "album-two"
	                    strAlbum = "albumtwo"
                    Case Else
	                    strAlbum = ""
                End Select

                If strAlbum <> "" Then
                    Set objImage = New Image
                    objImage.Album = strAlbum
                    objImage.Template = "<div class=""thumbnail""><a href=""[large]""><img src=""[thumb]"" alt=""[alt]"" title=""[title]"" class=""[class]"" /></a></div>"
                    objImage.Load()
                    Set objImage = Nothing
                End If

                strCategories = objCategory.GetPostCategories(objPost.ID)
                If Right(strCategories,2) = ", " Then strCategories = Left(strCategories,(Len(strCategories)-2))
                
                strTags = objTag.GetPostTags(objPost.ID)
                If Right(strTags,2) = ", " Then strTags = Left(strTags,(Len(strTags)-2))
                %>
                <div class="info">
                    <p><strong>Author:</strong> <%=objPost.Author%></p>
                    <p><strong>Created:</strong> <%=objPost.DateCreated%></p>
                    <p><strong>Last updated:</strong> <%=objPost.DateModified%></p>
                    <p><strong>Categories:</strong> <%=strCategories%></p>
                    <p><strong>Tags:</strong> <%=strTags%></p>
                    <%If Session("LoggedOn") = true Then%>
                    <p><a href="/admin/post.asp?id=<%=objPost.ID%>">Edit Post</a></p>
                    <%End If%>
                    <%=objSlickCMS.AddThis(objPost.Url,objPost.Title)%>
                </div>
                <!--#include virtual="/comments.asp"-->
                <%
            Case "archive","tag","category","pagination" 'multiple items
                strTemplate = "<h2><a href=""[url]"">[title]</a></h2>" & vbCrLf

			    If Application("FullPosts") = 0 Then
			        strTemplate = strTemplate & "[summary]" & vbCrLf
                Else
			        strTemplate = strTemplate & "[content]" & vbCrLf
                End If

			    strTemplate = strTemplate & "<div class=""info"">" & vbCrLf
			    strTemplate = strTemplate & "<p><strong>Author:</strong> [author]</p>" & vbCrLf
			    strTemplate = strTemplate & "<p><strong>Created:</strong> [datecreated]</p>" & vbCrLf
			    strTemplate = strTemplate & "<p><strong>Last updated:</strong> [datemodified]</p>" & vbCrLf
			    strTemplate = strTemplate & "<p><strong>Comments:</strong> <a href=""[url]#comments"">[comments]</a></p>" & vbCrLf
			    strTemplate = strTemplate & "<p><strong>Categories:</strong> [categories]</p>" & vbCrLf
			    strTemplate = strTemplate & "<p><strong>Tags:</strong> [tags]</p>" & vbCrLf
			    If Session("LoggedOn") = true Then
			        strTemplate = strTemplate & "<p><a href=""/admin/post.asp?id=[postid]"">Edit Post</a></p>"
			    End If
			    strTemplate = strTemplate & "</div>" & vbCrLf

			    objPost.PostsTemplate = strTemplate
        End Select

        'write out the Posts depending on what section we're in
        Select Case objPost.UrlType
            Case "archive"
	            Call objPost.GetPosts("archives")
            Case "tag"
                Call objPost.GetPosts("tags")
            Case "category"
                Call objPost.GetPosts("categories")
	        Case "pagination"
	            Call objPost.GetPosts("posts")
        End Select
        
        'show pagination if required
        Select Case objPost.UrlType
            Case "archive","tag","category","pagination" 'multiple items
                strPagination = "<div class=""pagination""><span class=""older"">[older]</span><span class=""newer"">[newer]</span></div>"
		        strPagination = objPost.Paging(strPagination)
		        Response.Write(strPagination)
        End Select
	%>
</div>