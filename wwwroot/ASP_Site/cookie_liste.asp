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
