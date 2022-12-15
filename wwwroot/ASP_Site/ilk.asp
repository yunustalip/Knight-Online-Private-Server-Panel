<html>
<head>
<title>Sayfa Başlığı</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9"></head>


<% 
' burası açıklamadır
'response.write("deneme<br>")
'response.write("deneme<br>")
'response.write("deneme<br>")
'response.write("deneme<br>")
response.write("deneme<br>")

dim cc
DIM dd
cc=5
if cc > 1 and _
cc <10 then
response.Write("if denemesi")
end if
%>


<table border="1">
<tr><td><%response.Write("tablo içi")%></td></tr>
</table> 



<body>
</body>
</html>
