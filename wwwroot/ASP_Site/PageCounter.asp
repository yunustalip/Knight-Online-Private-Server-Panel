<%

Set SayfaSayac = Server.CreateObject("MSWC.PageCounter")
SayfaSayac.PageHit

%>
<% 
'SayfaSayac.reset
'sayfasayac.reset ("/pagecount/index.asp")
%>
<%=sayfasayac.hits & "<br>"%>
<%=sayfasayac.hits ("/pagecount/index.asp") & "<br>"%>

