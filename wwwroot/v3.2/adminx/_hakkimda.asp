<% if session("admin") Then %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Blog Ekle</title>
<link rel="stylesheet" href="adminstil.css">
		<script type="text/javascript" src="scripts/wysiwyg.js"></script>
		<script type="text/javascript" src="scripts/wysiwyg-settings.js"></script>
		<!-- 
			Attach the editor on the textareas
		-->
		<script type="text/javascript">
			// Use it to attach the editor to all textareas with full featured setup
			//WYSIWYG.attach('all', full);
			
			// Use it to attach the editor directly to a defined textarea
			WYSIWYG.attach('hakkimda'); // default setup
			
			// Use it to display an iframes instead of a textareas
			//WYSIWYG.display('all', full);  
		</script>
</head>
<!--#include file="db.asp"-->
<!--#include file="../inc.asp"-->
<%
if (Request.QueryString("ayarlari"))="kaydet" then

Set ayar = Server.CreateObject("ADODB.Recordset")
SQL = "Select hakkimda from ayar"
ayar.Open SQL,data,1,3

ayar("hakkimda")=request.form("hakkimda")
ayar.update

ayar.Close
Set ayar = Nothing

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End if
if not hakkimda="" then
hakkimda=Replace(hakkimda,"&lt;","&amp;lt;")
hakkimda=Replace(hakkimda,"&gt;","&amp;gt;")
end if
%>
<body background="images/arka.gif">

<table border="0" width="100%" id="table1" cellpadding="0" style="border-collapse: collapse">
	<tr>
		<td width="10"><img border="0" src="images/bas.gif" width="10" height="25"></td>
		<td width="1215" background="images/bg.gif"><center><font class="baslik">Hakkımda</font></center></td>
		<td width="11"><img border="0" src="images/son.gif" width="15" height="25"></td>
	</tr>
</table>
<div align="center">
<table class="tablo" width="99%">
	<tr>
		<td></td>
		<td><a href="_yorum.asp?id=0">HAKKIMDA YAPILAN YORUMLAR</a></td>
	</tr>
<form action="?ayarlari=kaydet" method="post">
	<tr>
		<td width="186"></td>
		<td width="1037">
<textarea name="hakkimda" id="hakkimda"><%=hakkimda%></textarea>
		</td>
	</tr>
	<tr>
		<td width="186"></td>
		<td width="1037"><input type="submit" value="Kaydet" class="dugme"></td>
	</tr>
</form>
</table>
</div>
</body>

</html>
<% End if %>