<% 
ipadres=request.ServerVariables("REMOTE_ADDR")
'response.write ipadres&"<br>"

if ipadres="127.0.0.2" then
response.status= "do�ru ki�i ba�land�"
else
response.status= "do�ru ki�i de�ilsiniz"
response.write response.status
response.End
end if
%>

<table><tr><td>ho�geldiniz.</td></tr></table>

