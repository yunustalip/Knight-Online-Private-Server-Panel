<% if Session("durum")="esp" Then 
resimid=Request.Querystring("resim")
if isnumeric(resimid)=true Then
resimid=resimid&".gif"
Response.Write resimid
End If
End If
%>