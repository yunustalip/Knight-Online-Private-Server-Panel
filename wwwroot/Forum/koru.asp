
<!-- �YE OLMADAN G�R�LEMEYECEK SAYFALAR ���N INCLUDE ED�N -->
<%
Response.Buffer = True 
If Session("uyelogin")=True <> True Then 
Response.Redirect "default.asp?part=uyegorev&gorev=girisform" 
End If
%>

