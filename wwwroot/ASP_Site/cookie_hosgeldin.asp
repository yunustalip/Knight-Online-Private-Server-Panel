<%= "Hoþgeldin, " & request.Cookies("Kullanici") & "<br>"%>
<%
dogumtarihi = request.Cookies("dogumtarihi")
dogumgun = day(dogumtarihi)
dogumay = month(dogumtarihi)
sistemgun = day(date)
sistemay = month(date)
if  (dogumgun=sistemgun) and (dogumay=sistemay) then
response.write "doðum gününüz kutlu olsun <br>"
end if

%>



<%="----------------------<br>"%>

<%
for each cerez in request.cookies
response.write cerez & "=" & request.Cookies(cerez) & "<br>"
next
%>
<%="----------------------<br>"%>
