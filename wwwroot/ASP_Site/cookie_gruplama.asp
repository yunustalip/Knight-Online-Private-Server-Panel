<%
response.Cookies("Ziyaret�i")("Adi")="Bedri"
response.Cookies("Ziyaret�i")("Soyadi")="Akay"
response.Cookies("Ziyaret�i")("Email")="Bedri@yasalegitim.com"
response.Cookies("Ziyaret�i")("yasi")=25
response.Cookies("Font")="Arial"
response.Cookies("�ablon")="�ablon 1"
response.Cookies("Ad�")="Bedrettin" 
response.Cookies("Ziyaretci")=25
%>
<%="----------------------<br>"%>
<%
for each cerez in request.cookies
response.write cerez & "=" & chr(34)&  request.Cookies(cerez) & chr(34) &"<br>"
next
%>
<%="----------------------<br>"%>
