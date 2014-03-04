<!--#include virtual="/admin/templatetop.asp"-->
<h2>Admin | Dashboard</h2>
<p>Use the menu on the left to choose from the following areas:</p>
<ul>
	<li><em>Posts</em> - add or edit a post</li>
	<li><em>Comments</em> - moderate comments left by visitors</li>
	<li><em>Links</em> - add or edit a link</li>
	<li><em>Categories</em> - add or edit a category</li>
	<li><em>Tags</em> - add or edit a tag</li>
	<li><em>Relationships</em> - add or edit relationships between everything</li>
	<li><em>Users</em> - add or edit a user</li>
	<li><em>Feeds</em> - view site feeds</li>
	<li><em>Site</em> - return to the website</li>
</ul>

<h3>Uploads</h3>
<p>You can also <a href="/upload/">Upload</a> a file/image and view all of your <a href="/uploads/">Uploads</a>.</p>

<h3>Site Stats</h3>
<%
    Set objStatistic = New Statistic
    Response.Write(objStatistic.BlogStats("<p>There are currently [posts] posts, [comments] comments and [links] links contained within [categories] categories and [tags] tags, with [users] active users.</p><p>You are visitor number 1 of [activevisitors] current visitors; there have been a total of [totalvisitors] visitors since the last reset.</p>"))
    Set objStatistic = Nothing
%>
<!--#include virtual="/admin/templatebottom.asp"-->