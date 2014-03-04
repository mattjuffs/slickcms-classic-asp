<div id="sidebar">
    <h2>Pages</h2>
    <ul id="navigation">
        <%
	        objPost.NavigationTemplate = "<li><a href=""[url]"">[title]</a></li>"
	        Call objPost.Navigation(1) 'CategoryID
        %>
    </ul>
    
    <h2>Links</h2>
    <ul>
        <%
            'links - adapt for header/footer navigation and sidebar links
            Set objLink = New Link
            objLink.Template = "<li><a href=""[url]"" title=""[description]"">[name]</a></li>"
            objLink.CategoryID = 22
            objLink.GetLinks()
            Set objLink = Nothing
        %>
    </ul>

    <h2>Recent comments</h2>
    <ul>
        <%
            objComment.CommentsTemplate = "<li><a href=""[url]"" title=""[posttitle], [date]""><strong>[name]</strong>: [content]</a></li>"
            Call objComment.RecentComments()
        %>
    </ul>

    <%
        objPost.ArchivesTemplate = "<li><a href=""[url]"" title=""[archive] Archive"">[archive] ([postcount])</a></li>"
    %>
    <h2>Yearly Archives</h2>
    <ul>
        <%
            Call objPost.Archives("yearly")
        %>
    </ul>
    <h2>Monthly Archives</h2>
    <ul>
        <%
            Call objPost.Archives("monthly")
        %>
    </ul>
    <h2>Daily Archives</h2>
    <ul>
        <%
            Call objPost.Archives("daily")
        %>
    </ul>

    <h2>Categories</h2>
    <ul>
        <%
            objCategory.Template = "<li><a href=""[url]"" title=""[description]"">[name] ([postcount])</a></li>"
            Call objCategory.Categories()
        %>
    </ul>
    
    <h2>Tag Cloud</h2>
    <p id="tagcloud">
    <%
        objTag.Template = "<a href=""[url]"" title=""[postcount] posts"" class=""tag_[postcount]"">[name]</a> "
        Call objTag.Cloud()
    %>
    </p>

    <h2>Blog Stats</h2>
    <%=objStatistic.BlogStats("<p>There are currently [posts] posts, [comments] comments and [links] links contained within [categories] categories and [tags] tags.</p>")%>

    <br />
</div>