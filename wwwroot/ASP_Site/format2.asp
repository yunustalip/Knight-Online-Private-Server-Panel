
<%
'FormatDateTime(Deðer,Ayar)
' 0  : sistem ayarlarýný kullan
' 1  : tarihi uzun olarak göster
' 2  : tarihi kýsa göster
' 3  : saati ss:ddd:ss göster
' 4  : saati ss:dd göster
%>
<%=  formatdatetime(now,0)& "<br>"%>
<%=  formatdatetime(now,1)& "<br>"%>
<%=  formatdatetime(now,2)& "<br>"%>
<%=  formatdatetime(now,3)& "<br>"%>
<%=  formatdatetime(now,4)& "<br>"%>
<%=  formatdatetime(date,0)& "<br>"%>
<%=  formatdatetime(date,1)& "<br>"%>
<%=  formatdatetime(date,2)& "<br>"%>
<%=  formatdatetime(date,3)& "<br>"%>
<%=  formatdatetime(date,4)& "<br>"%>
<%=  date& "<br>"%>
<%=  now& "<br>"%>