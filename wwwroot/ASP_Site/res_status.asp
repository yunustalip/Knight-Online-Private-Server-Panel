<% 
ipadres=request.ServerVariables("REMOTE_ADDR")
'response.write ipadres&"<br>"

if ipadres="127.0.0.2" then
response.status= "doðru kiþi baðlandý"
else
response.status= "doðru kiþi deðilsiniz"
response.write response.status
response.End
end if
%>

<table><tr><td>hoþgeldiniz.</td></tr></table>

