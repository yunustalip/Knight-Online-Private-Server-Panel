<%
response.Cookies("Ziyaretci")("Adi")="Bedri"
response.Cookies("Ziyaretci")("Soyadi")="Akay"
response.Cookies("Ziyaretci")("Email")="Bedri@yasalegitim.com"
response.Cookies("Ziyaretci")("yasi")=25
response.Cookies("Font")="Arial"
response.Cookies("Sablon")="Þablon 1"
response.Cookies("Adi")="Bedrettin"

%>

<%
for each cerez in request.cookies
response.write cerez & "=" & request.Cookies(cerez) & "<br>"
next
%>
<%="----------------------<br>"%>
