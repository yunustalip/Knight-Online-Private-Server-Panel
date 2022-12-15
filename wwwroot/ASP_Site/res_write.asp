<%
dim xx,yy
xx="deneme"
yy=2536
response.write "--------ASP-------<br>"
response.write "xx:" & xx & "<br>"
'response.write xx
response.write ("yy:" & yy & "<br>")

satir= "<table border=" 
satir = satir & chr(34) & "1" & chr(34)
satir=satir & "><tr><td>tablo hücresi</td></tr></table>"
response.write satir

response.write "<table border=""1""><tr><td>"
response.write "tablo hücresi</td></tr></table>"


Response.Write("<table border=""1""><tr><td>tablo hücresi</td></tr></table>") 


%> 
-------HTML----------<br>

xx deðeri : <%=xx%> <br>
yy deðeri : <%=yy%> <br>

<table border="1"><tr><td>tablo hücresi</td></tr></table>