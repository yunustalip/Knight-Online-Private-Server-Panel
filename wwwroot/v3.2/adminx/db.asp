<!--#include file="../ayar.asp"-->
<%
Set data=Server.CreateObject("ADODB.Connection") 
data.Open"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("../"&mdbisim)
%>