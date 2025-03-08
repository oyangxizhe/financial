<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="UpLoadFile._Default" %>
<%@ Import Namespace="System"%>
<%@ Import Namespace="System.IO"%>
<%@ Import Namespace="System.Net"%>
<%@ Import NameSpace="System.Web"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>无标题页</title>

</head>
<body>
    <form id="form1" method ="post"  runat="server">
    <div>
    
    </div>
    </form>


<Script language="C#" runat="server">
void Page_Load(object sender, EventArgs e) {
	
	foreach(string f in Request.Files.AllKeys) {
		HttpPostedFile file = Request.Files[f];
		file.SaveAs("c:\\inetpub\\test\\UploadedFiles\\" + file.FileName);
	}	
}

</Script>



</body>
</html>
