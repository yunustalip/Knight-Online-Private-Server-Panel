<%
response.Cookies("Ziyaret�i")("Adi")="Bedri"
response.Cookies("Ziyaret�i")("Soyadi")="Akay"
response.Cookies("Ziyaret�i")("Email")="Bedri@yasalegitim.com"
response.Cookies("Ziyaret�i")("yasi")=25
response.Cookies("Font")="Arial"
response.Cookies("Font").Expires = "30/01/2006"
response.Cookies("�ablon")="�ablon 1"
response.Cookies("Ad�")="Bedrettin" 
response.Cookies("Ziyaretci")=25
response.write request.Cookies("Ad�") & "<br>"
response.write request.Cookies("Ziyaret�i")("Soyadi") & "<br>"
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
