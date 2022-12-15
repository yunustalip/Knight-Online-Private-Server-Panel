<%
'response.Cookies("Ziyaretçi").Expires = Now - 1
'response.Cookies("Font").Expires = Now - 1
'response.Cookies("Þablon").expires=Now - 1
'response.Cookies("Adý").Expires = Now - 1

for each cerez in request.cookies
response.cookies(cerez).expires = now-1
next

%>
