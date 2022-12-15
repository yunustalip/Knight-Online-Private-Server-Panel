
<%
'Response.Expires=0
Session.LCID = 1055
Session.CodePage = 1254
DefaultLCID = Session.LCID 
DefaultCodePage = Session.CodePage
part  = request.querystring("part")
if part="ana" or part="" then 
Server.Execute("ana.asp") 
Else
part=part&".asp"
Server.Execute(part) 
End If

%>