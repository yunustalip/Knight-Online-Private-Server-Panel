<%
Session("say") = Session("say") + 1
Response.Cookies("deneme") = "5"
Response.Write Session("say")
Response.Write "<br />"
Response.Write Session.SessionID
%>  