<%
'response.Cookies("Ziyaret�i").Expires = Now - 1
'response.Cookies("Font").Expires = Now - 1
'response.Cookies("�ablon").expires=Now - 1
'response.Cookies("Ad�").Expires = Now - 1

for each cerez in request.cookies
response.cookies(cerez).expires = now-1
next

%>
