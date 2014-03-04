<%@ Page Language="VB" AutoEventWireup="false" CodeFile="upload.aspx.vb" Inherits="Upload" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>SlickCMS File Upload Utility</title>
    <link href="/css/admin_screen.css" rel="stylesheet" type="text/css" media="screen" />
</head>
<body>
    <form id="uploadform" runat="server">
        <div id="fileupload">
        	<h1>File Upload</h1>
			<ol>
				<li>Choose a file to upload</li>
				<li>Tick Overwrite if you want to overwrite the existing file</li>
				<li>Click Upload File</li>
				<li>You can then copy/paste the sample code to integrate your file</li>
			</ol>
            <p><asp:FileUpload ID="FileUpload1" runat="server" /></p>
            <p>Overwrite? <asp:CheckBox ID="CheckBox1" runat="server" /></p>
            <p><asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="Upload File" /></p>
            <p><asp:Label ID="Label1" runat="server"></asp:Label></p>
            <p><small><a href="/admin/">Return to the admin</a></small></p>
        </div>
    </form>
</body>
</html>