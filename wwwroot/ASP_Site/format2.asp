
<%
'FormatDateTime(De�er,Ayar)
' 0  : sistem ayarlar�n� kullan
' 1  : tarihi uzun olarak g�ster
' 2  : tarihi k�sa g�ster
' 3  : saati ss:ddd:ss g�ster
' 4  : saati ss:dd g�ster
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