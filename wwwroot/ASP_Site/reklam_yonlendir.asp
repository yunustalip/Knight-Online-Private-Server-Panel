<%

url=request.QueryString("url")
if url <> "" then response.redirect url

Set reklamlar=Server.CreateObject("MSWC.Adrotator")
response.write reklamlar.getadvertisement("adrotator.txt")
%>
