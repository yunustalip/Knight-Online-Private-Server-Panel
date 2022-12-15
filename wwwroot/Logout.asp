<%
Session("username")=""
Session("login")=""
Session("yetki")=""
Session.abandon
Response.Redirect("login.asp")
%>