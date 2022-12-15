<%
response.Cookies("Ziyaretçi")("Adi")="Bedri"
response.Cookies("Ziyaretçi")("Soyadi")="Akay"
response.Cookies("Ziyaretçi")("Email")="Bedri@yasalegitim.com"
response.Cookies("Ziyaretçi")("yasi")=25
response.Cookies("Font")="Arial"
response.Cookies("Font").Expires = "30/01/2006"
response.Cookies("Þablon")="Þablon 1"
response.Cookies("Adý")="Bedrettin" 
response.Cookies("Ziyaretci")=25
response.write request.Cookies("Adý") & "<br>"
response.write request.Cookies("Ziyaretçi")("Soyadi") & "<br>"
%>
<%="----------------------<br>"%>
<%for each cerez in request.cookies

if not request.cookies(cerez).haskeys then
response.write cerez & "=" & chr(34)&  request.Cookies(cerez) & chr(34) &"<br>"
else

   for each anahtar in request.cookies(cerez)
   response.write cerez & "." & anahtar & "=" & chr(34)&  request.cookies(cerez)(anahtar) & chr(34) &"<br>"
   next
end if

next

%>
<%="----------------------<br>"%>
