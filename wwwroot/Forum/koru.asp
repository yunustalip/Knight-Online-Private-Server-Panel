
<!-- ÜYE OLMADAN GÝRÝLEMEYECEK SAYFALAR ÝÇÝN INCLUDE EDÝN -->
<%
Response.Buffer = True 
If Session("uyelogin")=True <> True Then 
Response.Redirect "default.asp?part=uyegorev&gorev=girisform" 
End If
%>

