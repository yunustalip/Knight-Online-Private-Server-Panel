<%
id=Request.Querystring("id")


Set adocon = Server.CreateObject("ADODB.Connection")
adocon.open= "driver={SQL Server};server=localhost;uid=radyolarimadmin;pwd=864327142358;database=radyolarimdb" 

Set musiclist =adocon.execute("SELECT albumcover FROM musiclist where id="&id)
if not musiclist.eof Then
Response.ContentType = "image/png"
response.binarywrite musiclist("albumcover")

End If
%> 