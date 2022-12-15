<%
if not session("admin") Then
response.redirect ("admin.asp")
Else
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

<title>Efendy Blog Admin Paneli</title>
</head>

<frameset rows="52,*" framespacing="0" frameborder="0" border="0">
<frame name="ust" src="ust.asp" marginwidth="0" marginheight="0" scrolling="no" noresize>
<frame name="alt" src="alt.asp">
<noframes>
<body>
<p>Taray&#305;c&#305;n&#305;z Çerçeve Desteklemiyor.</p>
</body>
</noframes>
</frameset>

</html>
<% End if %>