<%
set FSO = createobject("scripting.filesystemobject")

response.write fso.GetSpecialFolder (0) & "<br>"
response.write fso.GetSpecialFolder (1) & "<br>"
response.write fso.GetSpecialFolder (2) & "<br>"
%>
 
