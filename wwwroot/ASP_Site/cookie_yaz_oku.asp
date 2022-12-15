<%

response.Cookies("Kullanici")="Bedri"
response.Cookies("Email")="Bedri@yasalegitim.com"
response.Cookies("yasi")=25
response.Cookies("dogumtarihi")="17/02/1985"
%>
<%=request.Cookies("Kullanici") & "<br>"%>
<%=request.Cookies("Email") & "<br>"%>
<%=request.Cookies("yasi") & "<br>"%>
<%=request.Cookies("dogumtarihi") & "<br>"%>
<%="----------------------<br>"%>

<%
for each cerez in request.cookies
response.write cerez & "=" & request.Cookies(cerez) & "<br>"
next
%>
<%="----------------------<br>"%>
