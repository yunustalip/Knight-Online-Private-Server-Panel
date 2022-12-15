<%
response.Cookies("Ziyaretçi")("Adi")="Bedri"
response.Cookies("Ziyaretçi")("Soyadi")="Akay"
response.Cookies("Ziyaretçi")("Email")="Bedri@yasalegitim.com"
response.Cookies("Ziyaretçi")("yasi")=25
response.Cookies("Font")="Arial"
response.Cookies("Þablon")="Þablon 1"
response.Cookies("Adý")="Bedrettin" 
response.Cookies("Ziyaretci")=25
%>
<%="----------------------<br>"%>
<%
for each cerez in request.cookies
response.write cerez & "=" & chr(34)&  request.Cookies(cerez) & chr(34) &"<br>"
next
%>
<%="----------------------<br>"%>
